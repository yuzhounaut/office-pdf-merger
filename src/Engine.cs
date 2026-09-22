using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficePdf {
    public sealed class Engine {
        readonly string dir;
        readonly Job job;
        readonly Status status;
        readonly Stopwatch watch = Stopwatch.StartNew();
        string qpdf;
        public Engine(string directory) {
            dir = directory; job = Files.Read<Job>(Path.Combine(dir, "job.json"));
            using (var current = Process.GetCurrentProcess()) status = new Status { State = "running", Total = job.Items.Count, Items = job.Items, WorkerPid = current.Id, WorkerStartTicks = current.StartTime.ToUniversalTime().Ticks };
        }
        void Save(string message) {
            status.Message = message; status.ElapsedSeconds = watch.Elapsed.TotalSeconds;
            status.EtaSeconds = status.Completed > 0 && status.Completed < status.Total ? (double?)(watch.Elapsed.TotalSeconds * (status.Total - status.Completed) / status.Completed) : null;
            status.Updated = DateTime.UtcNow.ToString("o");
            Files.Write(status, Path.Combine(dir, "status.json"));
        }
        void Log(string message) { File.AppendAllText(Path.Combine(dir, "progress.log"), DateTime.Now.ToString("HH:mm:ss") + "  " + message + Environment.NewLine, Files.Utf8); Save(message); }
        void CheckCancel() { if (File.Exists(Path.Combine(dir, "cancel"))) throw new OperationCanceledException(); }
        string RunProcess(string exe, string arguments, int timeout, string ownerFile) {
            CheckCancel();
            var info = Files.StartInfo(exe, arguments); info.RedirectStandardOutput = true; info.RedirectStandardError = true;
            info.StandardOutputEncoding = Files.Utf8; info.StandardErrorEncoding = Files.Utf8;
            using (var process = new Process { StartInfo = info }) {
                process.Start(); Task<string> output = process.StandardOutput.ReadToEndAsync(), error = process.StandardError.ReadToEndAsync();
                Stopwatch elapsed = Stopwatch.StartNew();
                try {
                    while (!process.WaitForExit(250)) {
                        CheckCancel();
                        int limit = ownerFile != null && !File.Exists(ownerFile) ? Math.Max(10, timeout) : timeout;
                        if (elapsed.Elapsed.TotalSeconds > limit) throw new TimeoutException("单文件处理超过 " + timeout + " 秒，已停止。请检查文件能否正常打开，或调大超时。" );
                        Save(status.Message);
                    }
                    process.WaitForExit();
                    string stdout = output.Result, stderr = error.Result;
                    if (process.ExitCode == 3 && exe == qpdf) { status.Warnings++; Log("PDF 修复警告：" + stderr.Trim()); }
                    else if (process.ExitCode != 0) throw new InvalidOperationException((stderr + "\n" + stdout).Trim() + " (退出码 " + process.ExitCode + ")");
                    return stdout;
                } finally {
                    if (!process.HasExited) { try { process.Kill(); process.WaitForExit(5000); } catch { } }
                    CleanupOffice(ownerFile);
                }
            }
        }
        static void CleanupOffice(string path) {
            if (path == null || !File.Exists(path)) return;
            try {
                Owner owner = Files.Read<Owner>(path);
                using (Process p = Process.GetProcessById(owner.Pid)) {
                    if (p.StartTime.ToUniversalTime().Ticks == owner.StartTicks && new[] { "WINWORD", "POWERPNT" }.Contains(p.ProcessName.ToUpperInvariant())) { p.Kill(); p.WaitForExit(5000); }
                }
            } catch (ArgumentException) { } catch (InvalidOperationException) { } catch (System.ComponentModel.Win32Exception) { }
        }
        string Q(params string[] arguments) { return RunProcess(qpdf, string.Join(" ", arguments.Select(Files.Quote)), Math.Max(60, job.TimeoutSeconds), null); }
        int Validate(string pdf) {
            Q("--check", pdf);
            int count;
            if (!int.TryParse(Q("--show-npages", pdf).Trim(), out count) || count < 1) throw new InvalidDataException("PDF 没有有效页面。");
            return count;
        }
        void Merge(List<string> inputs, string target) {
            string argsfile = Path.Combine(dir, "merge-" + Guid.NewGuid().ToString("N") + ".args");
            var args = new List<string> { "--empty", "--pages" };
            foreach (string input in inputs) { args.Add(input); args.Add("1-z"); }
            args.Add("--"); args.Add(target);
            // UTF-8 response files avoid the Windows command-line length limit.
            File.WriteAllLines(argsfile, args, Files.Utf8);
            Q("@" + argsfile);
        }
        public int Run() {
            string work = Path.Combine(dir, "work");
            try {
                Save("正在准备内置 PDF 引擎");
                if (job.Items.Count == 0) throw new InvalidOperationException("没有待处理文件。");
                if (job.TimeoutSeconds < 1 || job.TimeoutSeconds > 3600) throw new InvalidOperationException("超时必须在 1–3600 秒之间。");
                Directory.CreateDirectory(work);
                RuntimeFiles.Verify();
                qpdf = Path.Combine(Files.BaseDirectory, "qpdf", "bin", "qpdf.exe"); Q("--version");
                Directory.CreateDirectory(job.OutputDirectory);
                var converted = new List<string>();
                var groups = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < job.Items.Count; i++) {
                    CheckCancel(); Item item = job.Items[i]; item.State = "转换中";
                    Log("[" + (i + 1) + "/" + job.Items.Count + "] " + item.Path);
                    string pdf = Path.Combine(work, i.ToString("D6") + ".pdf");
                    string request = Path.Combine(work, i.ToString("D6") + ".json");
                    string owner = request + ".owner";
                    try {
                        Files.Write(new ConvertRequest { Source = item.Path, Destination = pdf, OwnerFile = owner }, request);
                        RunProcess(Path.Combine(Files.BaseDirectory, "OfficePdfWorker.exe"), "--convert " + Files.Quote(request), job.TimeoutSeconds, owner);
                        item.Pages = Validate(pdf); item.FirstPage = status.Pages + 1; status.Pages += item.Pages;
                        converted.Add(pdf); string group = string.IsNullOrWhiteSpace(item.Group) ? "散文件" : item.Group;
                        if (!groups.ContainsKey(group)) groups[group] = new List<string>(); groups[group].Add(pdf);
                        item.State = "完成"; Log("已完成：" + Path.GetFileName(item.Path) + "（" + item.Pages + " 页）");
                    } catch (OperationCanceledException) { item.State = "已取消"; throw; }
                    catch (Exception ex) {
                        item.State = "失败"; item.Error = ex.Message; status.Failures++; Log("失败：" + item.Path + " — " + ex.Message);
                        if (File.Exists(pdf + ".trace")) File.Copy(pdf + ".trace", Path.Combine(dir, "failed-" + i + ".trace"), true);
                        if (!job.ContinueOnError) throw;
                    }
                    status.Completed++; Save(status.Message);
                }
                CheckCancel();
                if (converted.Count == 0) throw new InvalidOperationException("所有文件均失败，未生成 PDF。");
                string staged = Path.Combine(work, "merged.pdf"); Log("正在合并并校验 " + status.Pages + " 页"); Merge(converted, staged);
                if (Validate(staged) != status.Pages) throw new InvalidDataException("合并页数与输入页数不一致，结果未发布。");
                var groupResults = new List<KeyValuePair<string, string>>();
                if (job.PerGroup) {
                    int n = 0;
                    foreach (var group in groups) {
                        CheckCancel(); string output = Path.Combine(work, "group-" + (++n) + ".pdf"); Merge(group.Value, output); Validate(output);
                        groupResults.Add(new KeyValuePair<string, string>(Files.SafeName(group.Key) + ".pdf", output));
                    }
                }
                CheckCancel();
                string rawName = Path.GetFileName((job.OutputName ?? "").Trim());
                if (rawName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) {
                    rawName = rawName.Substring(0, rawName.Length - 4).TrimEnd();
                }
                string name = Files.SafeName(rawName);
                if (string.IsNullOrWhiteSpace(name)) name = "资料";
                if (status.Failures > 0) name += "_不完整";
                string destination = Files.Available(job.OutputDirectory, name + ".pdf");
                string groupDir = null;
                if (job.PerGroup) {
                    groupDir = Files.Available(job.OutputDirectory, name + "_分组PDF"); Directory.CreateDirectory(groupDir);
                    foreach (var output in groupResults) Publish(output.Value, Files.Available(groupDir, output.Key));
                }
                Publish(staged, destination); status.Output = destination;
                Directory.CreateDirectory(Files.Home);
                string destName = Path.GetFileName(destination);
                string baseName = destName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) ? destName.Substring(0, destName.Length - 4) : destName;
                string manifestContent = "文件,分组,状态,起始页,页数,错误\r\n" + string.Join("\r\n", job.Items.Select(x => string.Join(",", new[] { x.Path, x.Group, x.State, x.FirstPage.ToString(), x.Pages.ToString(), x.Error }.Select(Csv))));
                string manifestPath = Files.Available(Files.Home, baseName + ".清单.csv");
                File.WriteAllText(manifestPath, manifestContent, new UTF8Encoding(true));
                File.WriteAllText(Path.Combine(dir, "manifest.csv"), manifestContent, new UTF8Encoding(true));
                status.State = status.Failures > 0 ? "partial" : "completed";
                Log(status.Failures > 0 ? "已生成不完整 PDF；失败 " + status.Failures + " 个文件，请查看清单。" : "全部完成，共 " + status.Pages + " 页。" + (status.Warnings > 0 ? " 有 PDF 修复警告，请查看日志并检查页面。" : ""));
                string logPath = Files.Available(Files.Home, baseName + ".日志.txt");
                File.Copy(Path.Combine(dir, "progress.log"), logPath, true);
                return status.Failures > 0 ? 3 : 0;
            } catch (OperationCanceledException) { status.State = "cancelled"; Log("已取消；未发布本次总合并 PDF。"); return 2; }
            catch (Exception ex) { status.State = "failed"; Log("处理失败：" + ex.Message); return 1; }
            finally {
                // Work is created solely inside this job, never user inputs or output folders.
                try { if (Directory.Exists(work)) Directory.Delete(work, true); } catch { }
            }
        }
        static string Csv(string s) { return "\"" + (s ?? "").Replace("\"", "\"\"") + "\""; }
        static void Publish(string source, string destination) {
            string temp = destination + "." + Guid.NewGuid().ToString("N") + ".partial";
            try { File.Copy(source, temp, false); File.Move(temp, destination); }
            finally { if (File.Exists(temp)) File.Delete(temp); }
        }
    }
}
