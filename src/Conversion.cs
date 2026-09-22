using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;

namespace OfficePdf {
    public static class Conversion {
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
        static bool Track(dynamic app, string processName, HashSet<int> existing, string ownerFile) {
            try {
                int handle;
                try { handle = (int)app.HWND; } catch { handle = (int)app.HWND(); }
                uint pid; GetWindowThreadProcessId(new IntPtr(handle), out pid);
                Process process = Process.GetProcessById((int)pid);
                if (!existing.Contains((int)pid) && process.ProcessName.Equals(processName, StringComparison.OrdinalIgnoreCase)) {
                    Files.Write(new Owner { Pid = (int)pid, StartTicks = process.StartTime.ToUniversalTime().Ticks }, ownerFile);
                    return true;
                }
            } catch { }
            // Some PowerPoint installations do not expose Application.HWND. Accept only
            // one newly-created COM automation process, never an existing user process.
            var candidates = new List<Process>();
            foreach (Process process in Process.GetProcessesByName(processName)) {
                if (existing.Contains(process.Id)) continue;
                try {
                    using (var query = new ManagementObject("Win32_Process.Handle='" + process.Id + "'")) {
                        string command = Convert.ToString(query["CommandLine"]).ToLowerInvariant();
                        if (command.Contains("/automation") || command.Contains("-embedding") || command.Contains("/embedding")) candidates.Add(process);
                    }
                } catch { }
            }
            if (candidates.Count == 1) {
                Process process = candidates[0];
                Files.Write(new Owner { Pid = process.Id, StartTicks = process.StartTime.ToUniversalTime().Ticks }, ownerFile); return true;
            }
            return false;
        }
        public static void Run(ConvertRequest request) {
            if (!File.Exists(request.Source)) throw new FileNotFoundException("文件不存在或尚未下载到本机。", request.Source);
            Action<string> trace = message => File.AppendAllText(request.Destination + ".trace", DateTime.Now.ToString("o") + " " + message + Environment.NewLine, Files.Utf8);
            string ext = Path.GetExtension(request.Source).ToLowerInvariant();
            if (ext == ".pdf") { File.Copy(request.Source, request.Destination); return; }
            if (new[] { ".doc", ".docx", ".docm", ".rtf", ".ppt", ".pptx", ".pptm", ".pps", ".ppsx" }.Contains(ext)) {
                bool word = new[] { ".doc", ".docx", ".docm", ".rtf" }.Contains(ext);
                string name = word ? "WINWORD" : "POWERPNT";
                var existing = new HashSet<int>(Process.GetProcessesByName(name).Select(p => p.Id));
                Type type = Type.GetTypeFromProgID(word ? "Word.Application" : "PowerPoint.Application");
                if (type == null) throw new InvalidOperationException("此文件需要本机安装 Microsoft " + (word ? "Word" : "PowerPoint") + "。图片和 PDF 不需要 Office。");
                dynamic app = null, document = null;
                bool owned = false;
                try {
                    trace("Create Office instance"); app = Activator.CreateInstance(type);
                    owned = Track(app, name, existing, request.OwnerFile); trace("Ownership: " + owned);
                    // Never share an Office automation instance with an open user session.
                    if (!owned) throw new InvalidOperationException("无法获得独立的 " + name + " 实例；请保存并关闭该 Office 程序后重试。");
                    trace("Set Office options"); app.AutomationSecurity = 3;
                    if (word) {
                        app.Visible = false; app.DisplayAlerts = 0;
                        app.Options.UpdateLinksAtOpen = false;
                        trace("Open Word document"); document = app.Documents.Open(FileName: request.Source, ConfirmConversions: false, ReadOnly: true, AddToRecentFiles: false,
                            PasswordDocument: "OfficePdf-NoPassword", WritePasswordDocument: "OfficePdf-NoPassword", Revert: false, Visible: false, OpenAndRepair: false, NoEncodingDialog: true);
                        trace("Export Word PDF"); document.ExportAsFixedFormat(request.Destination, 17);
                    } else {
                        app.DisplayAlerts = 1;
                        trace("Open presentation"); document = app.Presentations.Open(request.Source, -1, 0, 0);
                        // Explicitly include hidden slides; every slide supplied by the user is a page.
                        // Office versions differ in honoring PrintHiddenSlides during PDF export.
                        // Change only the in-memory read-only presentation; never save the source.
                        for (int index = 1; index <= (int)document.Slides.Count; index++) {
                            dynamic slide = document.Slides.Item(index);
                            try { slide.SlideShowTransition.Hidden = 0; } finally { Release((object)slide); }
                        }
                        document.PrintOptions.PrintHiddenSlides = -1;
                        dynamic range = document.PrintOptions.Ranges.Add(1, document.Slides.Count);
                        try { document.ExportAsFixedFormat(request.Destination, 2, 2, 0, 1, 1, -1, range, 1, "", true, true, true, true, false); }
                        finally { Release((object)range); }
                        document.Saved = -1;
                    }
                } catch (Exception ex) { trace(ex.ToString()); throw; } finally {
                    trace("Close Office objects");
                    if ((object)document != null) { try { if (word) document.Close(0); else document.Close(); } catch { } Release((object)document); }
                    if ((object)app != null) { try { if (owned) app.Quit(); } catch { } Release((object)app); }
                }
            } else ImagePdf.Write(request.Source, request.Destination);
            if (!File.Exists(request.Destination) || new FileInfo(request.Destination).Length < 20) throw new IOException("转换后未生成有效 PDF。");
        }
        static void Release(object obj) { try { if (Marshal.IsComObject(obj)) Marshal.FinalReleaseComObject(obj); } catch { } }
    }
    // Image pages are emitted directly; there is no Word layout, trailing paragraph or blank page.
    public static class ImagePdf {
        static string N(double d) { return d.ToString("0.###", CultureInfo.InvariantCulture); }
        public static void Write(string source, string destination) {
            using (Image image = Image.FromFile(source))
            using (var stream = new FileStream(destination, FileMode.CreateNew, FileAccess.Write)) {
                int count = image.FrameDimensionsList.Contains(FrameDimension.Page.Guid) ? image.GetFrameCount(FrameDimension.Page) : 1;
                var offsets = new List<long> { 0 };
                Action<string> text = s => { byte[] b = Encoding.ASCII.GetBytes(s); stream.Write(b, 0, b.Length); };
                Action<int, string> obj = (id, data) => { offsets.Add(stream.Position); text(id + " 0 obj\n" + data + "\nendobj\n"); };
                text("%PDF-1.4\n%ImagePDF\n");
                obj(1, "<< /Type /Catalog /Pages 2 0 R >>");
                obj(2, "<< /Type /Pages /Count " + count + " /Kids [" + string.Join(" ", Enumerable.Range(0, count).Select(i => (3 + i * 3) + " 0 R")) + "] >>");
                for (int i = 0; i < count; i++) {
                    if (count > 1) image.SelectActiveFrame(FrameDimension.Page, i);
                    using (var frame = new Bitmap(image)) {
                        int orientation = 1;
                        try { orientation = BitConverter.ToUInt16(image.GetPropertyItem(0x0112).Value, 0); } catch (ArgumentException) { }
                        RotateFlipType[] transforms = { RotateFlipType.RotateNoneFlipNone, RotateFlipType.RotateNoneFlipNone, RotateFlipType.RotateNoneFlipX, RotateFlipType.Rotate180FlipNone, RotateFlipType.Rotate180FlipX, RotateFlipType.Rotate90FlipX, RotateFlipType.Rotate90FlipNone, RotateFlipType.Rotate270FlipX, RotateFlipType.Rotate270FlipNone };
                        if (orientation > 1 && orientation <= 8) frame.RotateFlip(transforms[orientation]);
                        double pw = frame.Width > frame.Height ? 841.89 : 595.276, ph = frame.Width > frame.Height ? 595.276 : 841.89;
                        double fit = Math.Min((pw - 36) / frame.Width, (ph - 36) / frame.Height);
                        double dw = frame.Width * fit, dh = frame.Height * fit;
                        double scale = Math.Min(1, Math.Min(dw * 300 / 72 / frame.Width, dh * 300 / 72 / frame.Height));
                        using (var rgb = new Bitmap(Math.Max(1, (int)Math.Round(frame.Width * scale)), Math.Max(1, (int)Math.Round(frame.Height * scale)), PixelFormat.Format24bppRgb)) {
                            using (Graphics g = Graphics.FromImage(rgb)) { g.Clear(Color.White); g.InterpolationMode = InterpolationMode.HighQualityBicubic; g.DrawImage(frame, new Rectangle(0, 0, rgb.Width, rgb.Height)); }
                            byte[] jpeg;
                            using (var bytes = new MemoryStream()) {
                                using (var parameters = new EncoderParameters(1)) {
                                    parameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 95L);
                                    rgb.Save(bytes, ImageCodecInfo.GetImageEncoders().First(c => c.MimeType == "image/jpeg"), parameters);
                                }
                                jpeg = bytes.ToArray();
                            }
                            int page = 3 + 3 * i;
                            obj(page, "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 " + N(pw) + " " + N(ph) + "] /Resources << /XObject << /Im0 " + (page + 2) + " 0 R >> >> /Contents " + (page + 1) + " 0 R >>");
                            string content = "q " + N(dw) + " 0 0 " + N(dh) + " " + N((pw - dw) / 2) + " " + N((ph - dh) / 2) + " cm /Im0 Do Q\n";
                            obj(page + 1, "<< /Length " + Encoding.ASCII.GetByteCount(content) + " >>\nstream\n" + content + "endstream");
                            offsets.Add(stream.Position);
                            text((page + 2) + " 0 obj\n<< /Type /XObject /Subtype /Image /Width " + rgb.Width + " /Height " + rgb.Height + " /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length " + jpeg.Length + " >>\nstream\n");
                            stream.Write(jpeg, 0, jpeg.Length); text("\nendstream\nendobj\n");
                        }
                    }
                }
                long xref = stream.Position;
                text("xref\n0 " + offsets.Count + "\n0000000000 65535 f \n");
                foreach (long offset in offsets.Skip(1)) text(offset.ToString("D10") + " 00000 n \n");
                text("trailer\n<< /Size " + offsets.Count + " /Root 1 0 R >>\nstartxref\n" + xref + "\n%%EOF\n");
            }
        }
    }
}
