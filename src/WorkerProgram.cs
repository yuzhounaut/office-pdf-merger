using System;
using System.IO;
using System.Text;

namespace OfficePdf {
    static class WorkerProgram {
        [STAThread]
        static int Main(string[] args) {
            try {
                if (args.Length == 2 && args[0] == "--worker") return new Engine(args[1]).Run();
                if (args.Length == 2 && args[0] == "--convert") { Conversion.Run(Files.Read<ConvertRequest>(args[1])); return 0; }
                throw new ArgumentException("工作程序由主界面启动，不需要单独打开。");
            } catch (Exception ex) {
                if (args.Length == 2 && args[0] == "--worker") {
                    try { File.AppendAllText(Path.Combine(args[1], "worker-error.log"), ex.ToString() + Environment.NewLine, Files.Utf8); } catch { }
                } else using (var stderr = new StreamWriter(Console.OpenStandardError(), new UTF8Encoding(false))) stderr.Write(ex.ToString());
                return 1;
            }
        }
    }
}
