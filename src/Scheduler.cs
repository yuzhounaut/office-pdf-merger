using System;
using System.Diagnostics;
using System.IO;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;

namespace OfficePdf {
    // Existing UI API name; jobs now use an ordinary fixed worker process.
    // No scheduled tasks, copied executables, startup entries or services.
    public static class Scheduler {
        public static string Submit(Job job) {
            string worker = Path.Combine(Files.BaseDirectory, "OfficePdfWorker.exe");
            if (!File.Exists(worker)) throw new FileNotFoundException("缺少工作程序。请完整解压整个软件文件夹；若文件被安全软件隔离，请先处理检测问题，不要添加排除项。", worker);
            RuntimeFiles.Verify();
            string dir = Path.Combine(Files.Home, "jobs", DateTime.Now.ToString("yyyyMMdd_HHmmss") + "_" + Guid.NewGuid().ToString("N").Substring(0, 8));
            Directory.CreateDirectory(dir);
            Files.Write(job, Path.Combine(dir, "job.json"));
            var initial = new Status { State = "queued", Message = "正在启动工作进程", Total = job.Items.Count, Items = job.Items, Updated = DateTime.UtcNow.ToString("o") };
            Files.Write(initial, Path.Combine(dir, "status.json"));
            File.WriteAllText(Path.Combine(Files.Home, "last-job.txt"), dir, Files.Utf8);
            try {
                using (var process = LaunchWorker(worker, dir)) {
                    var deadline = DateTime.UtcNow.AddSeconds(20);
                    while (DateTime.UtcNow < deadline) {
                        Status state = Files.Read<Status>(Path.Combine(dir, "status.json"));
                        if (state.WorkerPid == process.Id) return dir;
                        if (process.HasExited) throw new IOException("工作进程已退出（退出码 " + process.ExitCode + "）。请检查任务日志及 Windows 安全中心的保护历史记录。");
                        System.Threading.Thread.Sleep(100);
                    }
                    File.WriteAllText(Path.Combine(dir, "cancel"), "startup-timeout", Files.Utf8);
                    if (!process.HasExited) { process.Kill(); process.WaitForExit(5000); }
                    throw new IOException("工作进程在 20 秒内未报告启动状态。请检查 Windows 安全中心的保护历史记录。");
                }
            } catch (Exception ex) {
                initial.State = "failed"; initial.Message = "工作进程启动失败：" + ex.Message; initial.Updated = DateTime.UtcNow.ToString("o");
                Files.Write(initial, Path.Combine(dir, "status.json"));
                File.AppendAllText(Path.Combine(dir, "progress.log"), initial.Message + Environment.NewLine, Files.Utf8);
                throw new InvalidOperationException(initial.Message, ex);
            }
        }
        // .NET Framework Process.Start can pass all inheritable handles, including
        // the CLI caller's capture pipe, even when standard streams are redirected.
        // Use the OS API with inheritance explicitly disabled. No shell is involved.
        static Process LaunchWorker(string worker, string dir) {
            var startup = new StartupInfo { Size = Marshal.SizeOf(typeof(StartupInfo)), Flags = 1, ShowWindow = 0 };
            ProcessInfo created;
            if (!CreateProcess(worker, new StringBuilder(Files.Quote(worker) + " --worker " + Files.Quote(dir)), IntPtr.Zero, IntPtr.Zero, false, 0x08000000, IntPtr.Zero, Files.BaseDirectory, ref startup, out created))
                throw new Win32Exception(Marshal.GetLastWin32Error());
            try {
                var process = Process.GetProcessById(created.ProcessId);
                try { IntPtr keepIdentity = process.Handle; return process; }
                catch { process.Dispose(); throw; }
            } finally { CloseHandle(created.Thread); CloseHandle(created.Process); }
        }
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        struct StartupInfo {
            public int Size; public string Reserved, Desktop, Title;
            public int X, Y, XSize, YSize, XCountChars, YCountChars, FillAttribute, Flags;
            public short ShowWindow, ReservedSize; public IntPtr ReservedData, Input, Output, Error;
        }
        [StructLayout(LayoutKind.Sequential)] struct ProcessInfo { public IntPtr Process, Thread; public int ProcessId, ThreadId; }
        [DllImport("kernel32.dll", EntryPoint = "CreateProcessW", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool CreateProcess(string application, StringBuilder command, IntPtr processSecurity, IntPtr threadSecurity, [MarshalAs(UnmanagedType.Bool)] bool inheritHandles, uint flags, IntPtr environment, string directory, ref StartupInfo startup, out ProcessInfo process);
        [DllImport("kernel32.dll")] [return: MarshalAs(UnmanagedType.Bool)] static extern bool CloseHandle(IntPtr handle);
    }
}
