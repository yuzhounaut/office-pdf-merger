using System;
using System.IO;
using System.Text;
using System.Windows;

namespace OfficePdf {
    static class Program {
        [STAThread]
        static int Main(string[] args) {
            try {
                if (args.Length > 0) {
                    if (args[0] == "--submit") { string dir = Scheduler.Submit(Files.Read<Job>(args[1])); File.WriteAllText(args[1] + ".submitted", dir, Files.Utf8); return 0; }
                }
                var app = new Application { ShutdownMode = ShutdownMode.OnMainWindowClose };
                var ui = new ModernWindow(); app.MainWindow = ui.Window;
                app.DispatcherUnhandledException += (s, e) => { e.Handled = true; ui.Info("程序错误", e.Exception.Message); };
                if (args.Length >= 2 && args[0] == "--preview") {
                    ui.TestMode = true;
                    ui.Window.ContentRendered += async (s, e) => { await System.Threading.Tasks.Task.Delay(250); ui.SavePreview(args[1], 96); ui.Window.Close(); };
                }
                app.Run(ui.Window); return 0;
            } catch (Exception ex) {
                if (args.Length > 0) { using (var stderr = new StreamWriter(Console.OpenStandardError(), new UTF8Encoding(false))) stderr.Write(ex.ToString()); }
                else MessageBox.Show(ex.Message, "程序启动失败", MessageBoxButton.OK, MessageBoxImage.Error);
                return 1;
            }
        }
    }
}
