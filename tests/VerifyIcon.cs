using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

class VerifyIcon {
    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)] struct FileInfo {
        public IntPtr Icon; public int IconIndex; public uint Attributes;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst=260)] public string DisplayName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst=80)] public string TypeName;
    }
    [DllImport("shell32.dll", CharSet=CharSet.Unicode)] static extern IntPtr SHGetFileInfo(string path, uint attributes, out FileInfo info, uint size, uint flags);
    [DllImport("user32.dll")] static extern bool DestroyIcon(IntPtr icon);
    [DllImport("user32.dll", SetLastError=true)] static extern IntPtr SendMessageTimeout(IntPtr window, uint message, IntPtr wParam, IntPtr lParam, uint flags, uint timeout, out UIntPtr result);
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
    static void SaveAndCheck(IntPtr handle, string path) {
        if (handle == IntPtr.Zero) throw new Exception("Windows returned no icon: " + path);
        using (var icon = Icon.FromHandle(handle)) using (var bitmap = icon.ToBitmap()) {
            int blue=0, red=0, transparent=0;
            for (int y=0;y<bitmap.Height;y++) for(int x=0;x<bitmap.Width;x++) {
                Color c=bitmap.GetPixel(x,y); if(c.A<30) transparent++;
                if(c.A>128 && c.B>c.R+40 && c.B>100) blue++;
                if(c.A>128 && c.R>c.G+50 && c.R>c.B+40) red++;
            }
            if(blue<8 || red<3 || transparent==0) throw new Exception("Expected blue/red transparent icon not found: " + path);
            bitmap.Save(path, ImageFormat.Png);
            Console.WriteLine("PASS " + Path.GetFileName(path) + " " + bitmap.Width + "x" + bitmap.Height + " blue=" + blue + " red=" + red);
        }
    }
    static int Main(string[] args) {
        string exe=Path.GetFullPath(args[0]), output=Path.GetFullPath(args[1]); Directory.CreateDirectory(output);
        foreach(string file in new[]{exe,Path.Combine(Path.GetDirectoryName(exe),"OfficePdfWorker.exe")}) {
            foreach(uint small in new uint[]{0,1}) {
                FileInfo info;
                if(SHGetFileInfo(file,0,out info,(uint)Marshal.SizeOf(typeof(FileInfo)),0x100|small)==IntPtr.Zero) throw new Exception("SHGetFileInfo failed");
                try { SaveAndCheck(info.Icon,Path.Combine(output,Path.GetFileNameWithoutExtension(file)+(small==0?"-shell-large.png":"-shell-small.png"))); }
                finally { if(info.Icon!=IntPtr.Zero) DestroyIcon(info.Icon); }
            }
        }
        using(var process=Process.Start(new ProcessStartInfo(exe){UseShellExecute=false,WindowStyle=ProcessWindowStyle.Hidden})) {
            try {
                var deadline=DateTime.UtcNow.AddSeconds(12); IntPtr window=IntPtr.Zero;
                while(DateTime.UtcNow<deadline) { process.Refresh(); if(process.HasExited) throw new Exception("GUI exited before verification"); window=process.MainWindowHandle; if(window!=IntPtr.Zero) break; Thread.Sleep(100); }
                if(window==IntPtr.Zero) throw new Exception("No main window");
                foreach(int kind in new[]{0,1}) {
                    UIntPtr result;
                    if(SendMessageTimeout(window,0x7F,new IntPtr(kind),IntPtr.Zero,2,2000,out result)==IntPtr.Zero) throw new Exception("Window icon query timed out");
                    SaveAndCheck(new IntPtr(unchecked((long)result.ToUInt64())),Path.Combine(output,kind==0?"window-taskbar-small.png":"window-taskbar-large.png"));
                }
                PostMessage(window,0x10,IntPtr.Zero,IntPtr.Zero);
                if(!process.WaitForExit(5000)) throw new Exception("GUI did not close");
                Console.WriteLine("PASS real GUI startup / native icon handles / close");
            } finally { if(!process.HasExited) { process.Kill(); process.WaitForExit(5000); } }
        }
        return 0;
    }
}
