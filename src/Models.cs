using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Web.Script.Serialization;

namespace OfficePdf {
    public class Item {
        public string Path { get; set; }
        public string Group { get; set; }
        public string State { get; set; }
        public string Error { get; set; }
        public int Pages { get; set; }
        public int FirstPage { get; set; }
    }
    public class Job {
        public List<Item> Items { get; set; }
        public string OutputDirectory { get; set; }
        public string OutputName { get; set; }
        public bool PerGroup { get; set; }
        public bool ContinueOnError { get; set; }
        public int TimeoutSeconds { get; set; }
        public string TaskName { get; set; }
        public Job() { Items = new List<Item>(); OutputName = "总合并.pdf"; TimeoutSeconds = 180; }
    }
    public class Status {
        public string State { get; set; }
        public string Message { get; set; }
        public int Completed { get; set; }
        public int Total { get; set; }
        public int Failures { get; set; }
        public int Warnings { get; set; }
        public int Pages { get; set; }
        public string Output { get; set; }
        public string Updated { get; set; }
        public double ElapsedSeconds { get; set; }
        public double? EtaSeconds { get; set; }
        public int WorkerPid { get; set; }
        public long WorkerStartTicks { get; set; }
        public List<Item> Items { get; set; }
    }
    public class ConvertRequest { public string Source; public string Destination; public string OwnerFile; }
    public class Owner { public int Pid; public long StartTicks; }
    public static class Files {
        public static readonly Encoding Utf8 = new UTF8Encoding(false);
        public static readonly string BaseDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
        public static readonly string Home = System.IO.Path.Combine(BaseDirectory, "OfficePdfData");
        public static readonly HashSet<string> Supported = new HashSet<string>(new[] { ".doc", ".docx", ".docm", ".rtf", ".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".pdf", ".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".gif" }, StringComparer.OrdinalIgnoreCase);
        public static T Read<T>(string path) {
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var reader = new StreamReader(stream, Utf8)) return new JavaScriptSerializer { MaxJsonLength = int.MaxValue }.Deserialize<T>(reader.ReadToEnd());
        }
        public static void Write(object value, string path) {
            string temp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            File.WriteAllText(temp, new JavaScriptSerializer { MaxJsonLength = int.MaxValue }.Serialize(value), Utf8);
            try {
                for (int attempt = 0; ; attempt++) {
                    try { if (File.Exists(path)) File.Replace(temp, path, null); else File.Move(temp, path); break; }
                    catch (IOException) { if (attempt >= 10) throw; System.Threading.Thread.Sleep(30); }
                }
            }
            finally { if (File.Exists(temp)) File.Delete(temp); }
        }
        public static bool Terminal(string state) { return state == "completed" || state == "partial" || state == "failed" || state == "cancelled"; }
        public static string Quote(string s) {
            var result = new StringBuilder("\""); int slashes = 0;
            foreach (char c in s) {
                if (c == '\\') { slashes++; continue; }
                result.Append('\\', c == '"' ? slashes * 2 + 1 : slashes); result.Append(c); slashes = 0;
            }
            result.Append('\\', slashes * 2); result.Append('"'); return result.ToString();
        }
        public static ProcessStartInfo StartInfo(string exe, string args) { return new ProcessStartInfo(exe, args) { UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden }; }
        public static string SafeName(string s) {
            foreach (char c in System.IO.Path.GetInvalidFileNameChars()) s = s.Replace(c, '_');
            s = s.Trim().TrimEnd('.');
            if (s.Length == 0) s = "资料";
            if (new[] { "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "LPT1", "LPT2", "LPT3" }.Contains(s.ToUpperInvariant())) s = "_" + s;
            return s;
        }
        public static string Available(string directory, string name) {
            string path = System.IO.Path.Combine(directory, name);
            if (!File.Exists(path) && !Directory.Exists(path)) return path;
            string ext = "", baseName = name;
            int lastDot = name.LastIndexOf('.');
            if (lastDot > 0 && (name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) ||
                                name.EndsWith(".csv", StringComparison.OrdinalIgnoreCase) ||
                                name.EndsWith(".txt", StringComparison.OrdinalIgnoreCase))) {
                ext = name.Substring(lastDot);
                baseName = name.Substring(0, lastDot);
            }
            for (int i = 2; ; i++) {
                string candidate = System.IO.Path.Combine(directory, baseName + " (" + i + ")" + ext);
                if (!File.Exists(candidate) && !Directory.Exists(candidate)) return candidate;
            }
        }
        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)] static extern int StrCmpLogicalW(string a, string b);
        public static List<Item> Collect(IEnumerable<string> paths, string excludedOutput) {
            var result = new List<Item>();
            foreach (string raw in paths) {
                string path = System.IO.Path.GetFullPath(raw);
                if (Directory.Exists(path)) Scan(path, path, excludedOutput, result);
                else Add(path, "散文件", result);
            }
            return result.GroupBy(x => x.Path, StringComparer.OrdinalIgnoreCase).Select(g => g.First()).ToList();
        }
        static void Add(string path, string group, List<Item> result) {
            if (!System.IO.Path.GetFileName(path).StartsWith("~$") && Supported.Contains(System.IO.Path.GetExtension(path))) result.Add(new Item { Path = path, Group = group, State = "待处理" });
        }
        static void Scan(string root, string dir, string exclude, List<Item> result) {
            if ((File.GetAttributes(dir) & FileAttributes.ReparsePoint) != 0) return;
            string relative = dir.Substring(root.Length).TrimStart('\\');
            string group = relative.Length == 0 ? System.IO.Path.GetFileName(root) : relative.Split('\\')[0];
            foreach (string f in Directory.GetFiles(dir).OrderBy(x => x, Comparer<string>.Create(StrCmpLogicalW))) Add(f, group, result);
            foreach (string sub in Directory.GetDirectories(dir).OrderBy(x => x, Comparer<string>.Create(StrCmpLogicalW))) {
                if (System.IO.Path.GetFileName(sub).Equals("_PDF输出", StringComparison.OrdinalIgnoreCase) || (!string.IsNullOrWhiteSpace(exclude) && sub.TrimEnd('\\').Equals(exclude.TrimEnd('\\'), StringComparison.OrdinalIgnoreCase))) continue;
                Scan(root, sub, exclude, result);
            }
        }
    }
}
