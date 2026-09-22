using System;
using System.Runtime.InteropServices;

namespace OfficePdf {
    // The system Common Item Dialog uses the current Windows shell appearance.
    public static class FolderPicker {
        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = true)] static extern int SHCreateItemFromParsingName(string path, IntPtr context, ref Guid iid, out IShellItem item);
        public static string Show(IntPtr owner, string title, string initialDirectory = null) {
            IFileDialog dialog = (IFileDialog)new FileOpenDialog(); IShellItem item = null;
            try {
                uint options; dialog.GetOptions(out options); dialog.SetOptions(options | 0x20 | 0x40 | 0x800 | 0x8);
                dialog.SetTitle(title); dialog.SetOkButtonLabel("选择此文件夹");
                if (!string.IsNullOrWhiteSpace(initialDirectory)) {
                    IShellItem folder; Guid iid = typeof(IShellItem).GUID;
                    if (SHCreateItemFromParsingName(initialDirectory, IntPtr.Zero, ref iid, out folder) >= 0) { try { dialog.SetFolder(folder); } finally { Marshal.ReleaseComObject(folder); } }
                }
                int result = dialog.Show(owner); if (result == unchecked((int)0x800704C7)) return null;
                Marshal.ThrowExceptionForHR(result); dialog.GetResult(out item);
                IntPtr name; item.GetDisplayName(0x80058000, out name);
                try { return Marshal.PtrToStringUni(name); } finally { Marshal.FreeCoTaskMem(name); }
            } finally { if (item != null) Marshal.ReleaseComObject(item); Marshal.ReleaseComObject(dialog); }
        }
        [ComImport, Guid("DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7")] class FileOpenDialog { }
        [ComImport, Guid("42F85136-DB7E-439C-85F1-E4075D135FC8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IFileDialog {
            [PreserveSig] int Show(IntPtr owner);
            void SetFileTypes(uint count, IntPtr filters); void SetFileTypeIndex(uint index); void GetFileTypeIndex(out uint index);
            void Advise(IntPtr events, out uint cookie); void Unadvise(uint cookie); void SetOptions(uint options); void GetOptions(out uint options);
            void SetDefaultFolder(IShellItem folder); void SetFolder(IShellItem folder); void GetFolder(out IShellItem folder); void GetCurrentSelection(out IShellItem item);
            void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string name); void GetFileName(out IntPtr name);
            void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string title); void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string text); void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string text);
            void GetResult(out IShellItem result); void AddPlace(IShellItem item, int alignment); void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string extension);
            void Close(int result); void SetClientGuid(ref Guid guid); void ClearClientData(); void SetFilter(IntPtr filter);
        }
        [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IShellItem {
            void BindToHandler(IntPtr bindContext, ref Guid handler, ref Guid riid, out IntPtr result); void GetParent(out IShellItem parent);
            void GetDisplayName(uint format, out IntPtr name); void GetAttributes(uint mask, out uint attributes); void Compare(IShellItem item, uint hint, out int order);
        }
    }
}
