using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace OfficePdf {
    public sealed class FileRow : INotifyPropertyChanged {
        public Item Item;
        public int Index { get; set; }
        public string Name { get { return Path.GetFileName(Item.Path); } }
        public string FullPath { get { return Item.Path; } }
        public string Folder { get { return Path.GetDirectoryName(Item.Path); } }
        public string Group { get { return Item.Group; } }
        public string Pages { get { return Item.Pages == 0 ? "—" : Item.Pages.ToString(); } }
        public string State { get { return Item.State ?? "待处理"; } }
        public string Error { get { return Item.Error ?? Item.Path; } }
        public string Kind {
            get { string ext = Path.GetExtension(Item.Path).ToLowerInvariant(); if (new[] { ".doc", ".docx", ".docm", ".rtf" }.Contains(ext)) return "W"; if (new[] { ".ppt", ".pptx", ".pptm", ".pps", ".ppsx" }.Contains(ext)) return "P"; return ext == ".pdf" ? "PDF" : "IMG"; }
        }
        public string KindBackground { get { return Kind == "W" ? "#EAF1FD" : Kind == "P" ? "#FFF0E9" : Kind == "PDF" ? "#FDECEF" : "#E9F4EF"; } }
        public string KindForeground { get { return Kind == "W" ? "#2463AF" : Kind == "P" ? "#BC5A2F" : Kind == "PDF" ? "#B84B61" : "#35835E"; } }
        public string StateBackground { get { return State == "失败" ? "#FDEDED" : State == "完成" ? "#EAF5EE" : State == "转换中" ? "#E8F1FF" : "#F1F3F6"; } }
        public string StateForeground { get { return State == "失败" ? "#B33C38" : State == "完成" ? "#32724D" : State == "转换中" ? "#0067C0" : "#7B8795"; } }
        public event PropertyChangedEventHandler PropertyChanged;
        public void Notify() { var h = PropertyChanged; if (h != null) h(this, new PropertyChangedEventArgs(null)); }
    }
    public sealed class HistoryRow {
        public string Directory { get; set; }
        public string Title { get; set; }
        public string Detail { get; set; }
        public string State { get; set; }
        public string When { get; set; }
    }
    public sealed class ModernWindow {
        public readonly Window Window;
        readonly ObservableCollection<FileRow> rows = new ObservableCollection<FileRow>();
        readonly DispatcherTimer timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(750) };
        readonly List<Control> editors = new List<Control>();
        readonly DataGrid grid;
        readonly TextBox output, name, timeout;
        readonly CheckBox groups, skip;
        readonly TextBlock title, detail, metrics;
        readonly ProgressBar progress;
        string jobDir, lastViewedJob, resultFile, logText = "尚未开始任务。", previousUpdate;
        bool busy, scanning, polling, submitting, closed, used;
        int scanGeneration;
        Window logWindow;
        TextBox logView;
        public bool TestMode;
        public T Find<T>(string controlName) where T : class { return Window.FindName(controlName) as T; }
        [DllImport("dwmapi.dll")] static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int size);

        public ModernWindow() {
            using (Stream xaml = Assembly.GetExecutingAssembly().GetManifestResourceStream("MainWindow.xaml")) Window = (Window)XamlReader.Load(xaml);
            using (Stream icon = Assembly.GetExecutingAssembly().GetManifestResourceStream("AppIcon.png")) {
                var image = BitmapFrame.Create(icon, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
                image.Freeze(); Window.Icon = image; Find<Image>("AppLogo").Source = image;
            }
            grid = Find<DataGrid>("FilesGrid"); output = Find<TextBox>("OutputBox"); name = Find<TextBox>("NameBox"); timeout = Find<TextBox>("TimeoutBox");
            groups = Find<CheckBox>("GroupCheck"); skip = Find<CheckBox>("SkipCheck"); title = Find<TextBlock>("StatusTitle"); detail = Find<TextBlock>("StatusDetail"); metrics = Find<TextBlock>("Metrics"); progress = Find<ProgressBar>("JobProgress");
            output.Text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "_PDF输出"); output.ToolTip = output.Text; grid.ItemsSource = rows;
            foreach (string id in new[] { "AddFilesButton", "AddFolderButton", "ClearButton", "BrowseOutputButton", "OutputBox", "NameBox", "TimeoutBox", "GroupCheck", "SkipCheck" }) editors.Add(Find<Control>(id));
            Click("AddFilesButton", async () => await PickFiles()); Click("AddFolderButton", async () => await PickFolder());
            Click("MoveUpButton", () => Move(-1)); Click("MoveDownButton", () => Move(1)); Click("RemoveButton", Remove);
            Click("ClearButton", () => { if (!CanEdit) return; used = true; rows.Clear(); ResetResult(); RefreshRows(); });
            Click("BrowseOutputButton", PickOutput); Click("StartButton", async () => await StartJob());
            Click("CancelButton", async () => await Cancel()); Click("OpenOutputButton", async () => await OpenOutput());
            Click("LogsButton", ShowLogs);
            Click("MinimizeButton", () => Window.WindowState = WindowState.Minimized);
            Click("MaximizeButton", () => Window.WindowState = Window.WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized);
            Click("CloseButton", () => Window.Close());
            Click("RefreshHistoryButton", async () => await LoadHistory()); Click("LoadHistoryButton", async () => await OpenHistory());
            Find<RadioButton>("MergeNav").Checked += (s, e) => Page("MergePage");
            Find<RadioButton>("HistoryNav").Checked += async (s, e) => { Page("HistoryPage"); await LoadHistory(); };
            Find<RadioButton>("HelpNav").Checked += (s, e) => Page("HelpPage");
            Find<ListBox>("HistoryList").MouseDoubleClick += async (s, e) => await OpenHistory();
            grid.SelectionChanged += (s, e) => UpdateEnabled();
            grid.SizeChanged += (s, e) => ReflowColumns();
            grid.MouseDoubleClick += (s, e) => { FileRow row = grid.SelectedItem as FileRow; if (row != null) Info(row.Name, row.FullPath + (string.IsNullOrWhiteSpace(row.Item.Error) ? "" : "\n\n" + row.Item.Error)); };
            var card = Find<Border>("FilesCard");
            card.PreviewDragOver += (s, e) => { bool accept = CanEdit && e.Data.GetDataPresent(DataFormats.FileDrop); e.Effects = accept ? DragDropEffects.Copy : DragDropEffects.None; Find<Border>("DropOverlay").Visibility = accept ? Visibility.Visible : Visibility.Collapsed; e.Handled = true; };
            card.DragLeave += (s, e) => Find<Border>("DropOverlay").Visibility = Visibility.Collapsed;
            card.PreviewDrop += async (s, e) => { Find<Border>("DropOverlay").Visibility = Visibility.Collapsed; e.Handled = true; if (CanEdit && e.Data.GetDataPresent(DataFormats.FileDrop)) await AddPaths((string[])e.Data.GetData(DataFormats.FileDrop)); };
            Window.PreviewKeyDown += async (s, e) => {
                if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.O) { e.Handled = true; await PickFiles(); }
                else if (Keyboard.Modifiers == (ModifierKeys.Control | ModifierKeys.Shift) && e.Key == Key.O) { e.Handled = true; await PickFolder(); }
                else if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.Enter) { e.Handled = true; await StartJob(); }
                else if (grid.IsKeyboardFocusWithin) {
                    Key key = e.Key == Key.System ? e.SystemKey : e.Key;
                    if (e.Key == Key.Delete) { e.Handled = true; Remove(); }
                    else if (Keyboard.Modifiers == ModifierKeys.Alt && (key == Key.Up || key == Key.Down)) { e.Handled = true; Move(key == Key.Up ? -1 : 1); }
                }
            };
            Window.SourceInitialized += (s, e) => {
                try { int corner = 2; DwmSetWindowAttribute(new WindowInteropHelper(Window).Handle, 33, ref corner, sizeof(int)); } catch (DllNotFoundException) { }
                Rect area = SystemParameters.WorkArea; Window.Width = Math.Min(1200, area.Width - 24); Window.Height = Math.Min(860, area.Height - 24);
            };
            Window.SizeChanged += (s, e) => {
                bool narrow = Window.ActualWidth < 1080;
                Find<ColumnDefinition>("RailColumn").Width = new GridLength(narrow ? 160 : 186);
                Find<ColumnDefinition>("SettingsColumn").Width = new GridLength(narrow ? 260 : 292);
                grid.Columns[2].Visibility = narrow ? Visibility.Collapsed : Visibility.Visible;
                ReflowColumns();
            };
            Window.StateChanged += (s, e) => {
                Find<Button>("MaximizeButton").Content = Window.WindowState == WindowState.Maximized ? "\uE923" : "\uE922";
                Find<Grid>("RootGrid").Margin = Window.WindowState == WindowState.Maximized ? new Thickness(7) : new Thickness(0);
            };
            Window.Closed += (s, e) => { closed = true; scanGeneration++; timer.Stop(); if (logWindow != null) logWindow.Close(); };
            Window.ContentRendered += async (s, e) => { if (!TestMode) await Resume(); };
            output.TextChanged += (s, e) => output.ToolTip = output.Text;
            timer.Tick += async (s, e) => await Poll(); timer.Start(); RefreshRows();
        }
        bool CanEdit { get { return !busy && !scanning && !submitting; } }
        void ReflowColumns() {
            if (grid.ActualWidth <= 0) return;
            double fixedWidth = 38 + 47 + 74 + (grid.Columns[2].Visibility == Visibility.Visible ? 80 : 0) + 20;
            grid.Columns[1].Width = new DataGridLength(Math.Max(135, grid.ActualWidth - fixedWidth));
            ScrollViewer scroller = FindVisual<ScrollViewer>(grid); if (scroller != null) scroller.ScrollToHorizontalOffset(0);
        }
        static T FindVisual<T>(DependencyObject root) where T : DependencyObject { for (int i = 0; i < VisualTreeHelper.GetChildrenCount(root); i++) { DependencyObject child = VisualTreeHelper.GetChild(root, i); T value = child as T ?? FindVisual<T>(child); if (value != null) return value; } return null; }
        void Click(string id, Action action) { Find<Button>(id).Click += (s, e) => action(); }
        void Page(string id) { foreach (string page in new[] { "MergePage", "HistoryPage", "HelpPage" }) Find<Grid>(page).Visibility = page == id ? Visibility.Visible : Visibility.Collapsed; }
        void StatusText(string heading, string description) { title.Text = heading; detail.Text = description; detail.ToolTip = description; }
        void Error(Exception ex) { StatusText("需要处理", ex.Message); Info("无法完成操作", ex.Message); }
        public void Info(string heading, string text) {
            var dialog = new Window { Title = heading, Owner = Window, Width = 560, Height = 300, MinWidth = 420, MinHeight = 240, WindowStartupLocation = WindowStartupLocation.CenterOwner, Background = (Brush)new BrushConverter().ConvertFromString("#F6F8FB"), FontFamily = Window.FontFamily, FontSize = 13, ShowInTaskbar = false };
            dialog.Resources.MergedDictionaries.Add(Window.Resources);
            var panel = new DockPanel { Margin = new Thickness(24) }; var headingText = new TextBlock { Text = heading, FontSize = 19, FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 16) }; DockPanel.SetDock(headingText, Dock.Top); panel.Children.Add(headingText);
            var button = new Button { Content = "知道了", Width = 96, Margin = new Thickness(0, 16, 0, 0), HorizontalAlignment = HorizontalAlignment.Right, IsDefault = true, IsCancel = true, Style = (Style)Window.Resources["Primary"] }; DockPanel.SetDock(button, Dock.Bottom); button.Click += (s, e) => dialog.Close(); panel.Children.Add(button);
            panel.Children.Add(new TextBox { Text = text, IsReadOnly = true, TextWrapping = TextWrapping.Wrap, VerticalScrollBarVisibility = ScrollBarVisibility.Auto }); dialog.Content = panel; dialog.ShowDialog();
        }
        void UpdateEnabled() {
            if (closed) return;
            foreach (Control c in editors) c.IsEnabled = CanEdit;
            Find<Button>("StartButton").IsEnabled = CanEdit && rows.Count > 0;
            foreach (string id in new[] { "MoveUpButton", "MoveDownButton", "RemoveButton" }) Find<Button>(id).IsEnabled = CanEdit && grid.SelectedItems.Count > 0;
            Find<Button>("ClearButton").IsEnabled = CanEdit && rows.Count > 0;
            Find<Button>("CancelButton").Visibility = busy ? Visibility.Visible : Visibility.Collapsed;
            Find<Button>("CancelButton").IsEnabled = busy && !submitting && jobDir != null;
            Find<Button>("OpenOutputButton").Visibility = busy ? Visibility.Collapsed : Visibility.Visible;
            Find<Button>("LogsButton").IsEnabled = lastViewedJob != null;
        }
        void ResetResult() { resultFile = null; previousUpdate = null; jobDir = null; progress.Value = 0; title.Foreground = (Brush)Window.Resources["Ink"]; StatusText("准备就绪", "确认文件顺序和输出设置后，点击“生成 PDF”。"); metrics.Text = "文件在本机完成转换和合并"; }
        void RefreshRows() { for (int i = 0; i < rows.Count; i++) { rows[i].Index = i + 1; rows[i].Notify(); } Find<TextBlock>("FileCount").Text = rows.Count + " 个文件"; Find<Border>("EmptyPanel").Visibility = rows.Count == 0 ? Visibility.Visible : Visibility.Collapsed; UpdateEnabled(); }
        async Task PickFiles() {
            if (!CanEdit) return;
            var picker = new Microsoft.Win32.OpenFileDialog { Multiselect = true, Title = "选择要合并的资料", Filter = "支持的资料|*.doc;*.docx;*.docm;*.rtf;*.ppt;*.pptx;*.pptm;*.pps;*.ppsx;*.pdf;*.jpg;*.jpeg;*.png;*.bmp;*.tif;*.tiff;*.gif|所有文件|*.*" };
            if (picker.ShowDialog(Window) == true) await AddPaths(picker.FileNames);
        }
        async Task PickFolder() {
            if (!CanEdit) return;
            try { string selected = FolderPicker.Show(new WindowInteropHelper(Window).Handle, "选择资料文件夹（包含子文件夹）"); if (selected != null) await AddPaths(new[] { selected }); } catch (Exception ex) { Error(ex); }
        }
        void PickOutput() { if (!CanEdit) return; try { string selected = FolderPicker.Show(new WindowInteropHelper(Window).Handle, "选择 PDF 保存目录"); if (selected != null) output.Text = selected; } catch (Exception ex) { Error(ex); } }
        public async Task AddPaths(string[] paths) {
            if (!CanEdit) return; used = true; scanning = true; int generation = ++scanGeneration; UpdateEnabled(); StatusText("正在读取资料", "文件夹较大时需要一点时间，窗口仍可正常操作。"); string excluded = output.Text;
            try {
                List<Item> items = await Task.Run(() => Files.Collect(paths, excluded));
                if (closed || generation != scanGeneration) return;
                var known = new HashSet<string>(rows.Select(x => x.Item.Path), StringComparer.OrdinalIgnoreCase); int added = 0;
                foreach (Item item in items) if (known.Add(item.Path)) { rows.Add(new FileRow { Item = item }); added++; }
                ResetResult(); RefreshRows(); StatusText(added == 0 ? "没有添加新文件" : "已添加 " + added + " 个文件", added == 0 ? "重复文件会自动忽略；请确认文件格式受支持。" : "可选中文件上移 / 下移，清单顺序就是最终页序。");
            } catch (Exception ex) { if (!closed) Error(ex); }
            finally { scanning = false; if (!closed) UpdateEnabled(); }
        }
        void Move(int direction) {
            if (!CanEdit) return; used = true; var selected = new HashSet<FileRow>(grid.SelectedItems.Cast<FileRow>());
            IEnumerable<FileRow> sequence = direction < 0 ? selected.OrderBy(r => rows.IndexOf(r)) : selected.OrderByDescending(r => rows.IndexOf(r));
            foreach (FileRow row in sequence.ToArray()) { int i = rows.IndexOf(row), next = i + direction; if (next < 0 || next >= rows.Count || selected.Contains(rows[next])) continue; rows.Move(i, next); }
            ResetResult(); RefreshRows(); grid.SelectedItems.Clear(); foreach (FileRow row in selected) grid.SelectedItems.Add(row);
        }
        void Remove() { if (!CanEdit) return; used = true; foreach (FileRow row in grid.SelectedItems.Cast<FileRow>().ToArray()) rows.Remove(row); ResetResult(); RefreshRows(); }
        public async Task StartJob() {
            if (!CanEdit || rows.Count == 0) return;
            int seconds;
            if (!int.TryParse(timeout.Text, out seconds) || seconds < 10 || seconds > 3600) { Error(new Exception("单文件超时请填写 10–3600 之间的整数秒数。")); timeout.Focus(); return; }
            if (string.IsNullOrWhiteSpace(output.Text) || string.IsNullOrWhiteSpace(name.Text)) { Error(new Exception("请填写输出目录和文件名称。")); return; }
            used = true;
            var job = new Job { OutputDirectory = output.Text.Trim(), OutputName = name.Text.Trim(), PerGroup = groups.IsChecked == true, ContinueOnError = skip.IsChecked == true, TimeoutSeconds = seconds, Items = rows.Select(x => new Item { Path = x.Item.Path, Group = x.Item.Group, State = "待处理" }).ToList() };
            busy = true; submitting = true; jobDir = null; previousUpdate = null; resultFile = null; logText = "正在启动任务…"; progress.Value = 0; UpdateEnabled(); StatusText("正在启动", "后台任务启动后，关闭窗口也会继续处理。");
            try { string directory = await Task.Run(() => Scheduler.Submit(job)); jobDir = directory; lastViewedJob = directory; if (!closed) { for (int i = 0; i < rows.Count; i++) { rows[i].Item = job.Items[i]; rows[i].Notify(); } } }
            catch (Exception ex) { busy = false; if (!closed) Error(ex); }
            finally { submitting = false; if (!closed) UpdateEnabled(); }
            if (!closed) await Poll();
        }
        public async Task Cancel() {
            string directory = jobDir; if (!busy || submitting || directory == null) return; Find<Button>("CancelButton").IsEnabled = false;
            try { await Task.Run(() => File.WriteAllText(Path.Combine(directory, "cancel"), "cancel", Files.Utf8)); if (!closed) StatusText("正在取消", "当前转换结束后会清理临时文件。"); }
            catch (Exception ex) { if (!closed) { UpdateEnabled(); Error(ex); } }
        }
        async Task OpenOutput() {
            string directory = resultFile == null ? output.Text : Path.GetDirectoryName(resultFile);
            try { await Task.Run(() => { Directory.CreateDirectory(directory); Process.Start(new ProcessStartInfo("explorer.exe", Files.Quote(directory)) { UseShellExecute = true }); }); } catch (Exception ex) { if (!closed) Error(ex); }
        }
        async Task Resume() {
            try {
                string last = await Task.Run(() => { string path = Path.Combine(Files.Home, "last-job.txt"); return File.Exists(path) ? File.ReadAllText(path, Files.Utf8) : null; });
                if (last != null && !closed && !used) await Attach(last);
            } catch (Exception ex) { if (!closed) StatusText("准备就绪", "上次任务记录暂时无法读取，可重新添加资料。" + ex.Message); }
        }
        async Task Attach(string directory) {
            var data = await Task.Run(() => Tuple.Create(Files.Read<Job>(Path.Combine(directory, "job.json")), Files.Read<Status>(Path.Combine(directory, "status.json"))));
            if (closed) return; jobDir = directory; lastViewedJob = directory; previousUpdate = null; busy = !Files.Terminal(data.Item2.State); resultFile = data.Item2.Output;
            output.Text = data.Item1.OutputDirectory; name.Text = data.Item1.OutputName; timeout.Text = data.Item1.TimeoutSeconds.ToString(); groups.IsChecked = data.Item1.PerGroup; skip.IsChecked = data.Item1.ContinueOnError;
            rows.Clear(); foreach (Item item in data.Item2.Items ?? data.Item1.Items) rows.Add(new FileRow { Item = item }); RefreshRows(); await Poll();
        }
        public async Task Poll() {
            if (jobDir == null || polling || closed || submitting) return; polling = true; string current = jobDir;
            try {
                var data = await Task.Run(() => {
                    Status snapshot = Files.Read<Status>(Path.Combine(current, "status.json"));
                    DateTime updated;
                    if (!Files.Terminal(snapshot.State) && snapshot.WorkerPid > 0 && DateTime.TryParse(snapshot.Updated, out updated) && (DateTime.UtcNow - updated.ToUniversalTime()).TotalSeconds > 15) {
                        bool alive; try { using (var worker = Process.GetProcessById(snapshot.WorkerPid)) alive = !worker.HasExited && worker.ProcessName == "OfficePdfWorker" && worker.StartTime.ToUniversalTime().Ticks == snapshot.WorkerStartTicks; } catch (ArgumentException) { alive = false; }
                        if (!alive) { snapshot.State = "failed"; snapshot.Message = "工作进程已中断。请查看日志及 Windows 安全中心的保护历史记录；若被拦截，请先处理检测问题，不要添加排除项。"; snapshot.Updated = DateTime.UtcNow.ToString("o"); Files.Write(snapshot, Path.Combine(current, "status.json")); }
                    }
                    string path = Path.Combine(current, "progress.log"); return Tuple.Create(snapshot, File.Exists(path) ? ReadTail(path) : "正在准备任务…");
                });
                if (closed || current != jobDir) return; Status state = data.Item1;
                if (previousUpdate == state.Updated) return; previousUpdate = state.Updated; busy = !Files.Terminal(state.State); resultFile = state.Output; logText = data.Item2;
                if (state.Items != null && rows.Count == state.Items.Count) for (int i = 0; i < rows.Count; i++) { rows[i].Item = state.Items[i]; rows[i].Notify(); }
                string heading = StateTitle(state.State);
                if (state.State == "running") heading = state.Completed >= state.Total ? "正在合并并校验" : "正在处理  " + Math.Min(state.Completed + 1, state.Total) + " / " + state.Total;
                StatusText(heading, state.Message);
                title.Foreground = new BrushConverter().ConvertFromString(state.State == "failed" ? "#B33C38" : state.State == "partial" ? "#916112" : state.State == "completed" ? "#267448" : "#202733") as Brush;
                progress.Value = state.State == "completed" || state.State == "partial" ? 100 : state.Total == 0 ? 0 : Math.Min(98, state.Completed * 95 / state.Total);
                metrics.Text = state.Completed + "/" + state.Total + " 个文件  ·  " + state.Pages + " 页  ·  用时 " + Duration(state.ElapsedSeconds) + (state.EtaSeconds.HasValue ? "  ·  剩余约 " + Duration(state.EtaSeconds.Value) : "") + (state.Failures > 0 ? "  ·  失败 " + state.Failures : "") + (state.Warnings > 0 ? "  ·  警告 " + state.Warnings : "");
                if (logView != null && logView.Text != logText) { logView.Text = logText; logView.ScrollToEnd(); }
                UpdateEnabled(); if (!busy) jobDir = null;
            } catch (IOException) { } catch (Exception ex) { if (!closed) StatusText("暂时无法读取进度", ex.Message); }
            finally { polling = false; }
        }
        static string Duration(double seconds) { return TimeSpan.FromSeconds(Math.Max(0, Math.Min(seconds, 9999999))).ToString(@"hh\:mm\:ss"); }
        public static string StateTitle(string state) { switch (state) { case "completed": return "合并完成"; case "partial": return "已生成不完整 PDF"; case "failed": return "处理失败"; case "cancelled": return "任务已取消"; case "queued": return "正在启动"; default: return "处理中"; } }
        static string ReadTail(string path) {
            using (var f = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete)) { if (f.Length > 64000) f.Seek(-64000, SeekOrigin.End); using (var reader = new StreamReader(f, Files.Utf8)) { if (f.Position > 0) reader.ReadLine(); return reader.ReadToEnd(); } }
        }
        void ShowLogs() {
            if (logWindow != null) { logWindow.Activate(); return; }
            logWindow = new Window { Title = "处理日志", Owner = Window, Width = 780, Height = 480, MinWidth = 520, MinHeight = 320, WindowStartupLocation = WindowStartupLocation.CenterOwner, FontFamily = Window.FontFamily, FontSize = 12, Background = Brushes.White, ShowInTaskbar = false };
            logWindow.Resources.MergedDictionaries.Add(Window.Resources);
            var panel = new DockPanel { Margin = new Thickness(20) }; var heading = new TextBlock { Text = "处理日志", FontSize = 20, FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 16) }; DockPanel.SetDock(heading, Dock.Top); panel.Children.Add(heading);
            logView = new TextBox { IsReadOnly = true, Text = logText, TextWrapping = TextWrapping.Wrap, VerticalScrollBarVisibility = ScrollBarVisibility.Auto }; panel.Children.Add(logView); logWindow.Content = panel; logWindow.Closed += (s, e) => { logWindow = null; logView = null; }; logWindow.Show();
        }
        async Task LoadHistory() {
            Find<Button>("RefreshHistoryButton").IsEnabled = false;
            try {
                var history = await Task.Run(() => {
                    string directory = Path.Combine(Files.Home, "jobs"); var result = new List<HistoryRow>(); if (!Directory.Exists(directory)) return result;
                    foreach (string dir in Directory.GetDirectories(directory).OrderByDescending(x => x).Take(50)) {
                        try { Status state = Files.Read<Status>(Path.Combine(dir, "status.json")); Job job = Files.Read<Job>(Path.Combine(dir, "job.json")); DateTime updated; DateTime.TryParse(state.Updated, out updated); result.Add(new HistoryRow { Directory = dir, Title = job.OutputName, Detail = state.Total + " 个文件 · " + state.Pages + " 页 · " + job.OutputDirectory, State = StateTitle(state.State), When = updated.ToLocalTime().ToString("yyyy-MM-dd  HH:mm") }); } catch (IOException) { } catch (ArgumentException) { }
                    }
                    return result;
                });
                if (!closed) { Find<ListBox>("HistoryList").ItemsSource = history; Find<TextBlock>("HistoryEmpty").Visibility = history.Count == 0 ? Visibility.Visible : Visibility.Collapsed; Find<Button>("LoadHistoryButton").IsEnabled = history.Count > 0; }
            } catch (Exception ex) { if (!closed) Error(ex); }
            finally { if (!closed) Find<Button>("RefreshHistoryButton").IsEnabled = true; }
        }
        async Task OpenHistory() {
            HistoryRow selected = Find<ListBox>("HistoryList").SelectedItem as HistoryRow; if (selected == null) return;
            if ((busy || submitting) && selected.Directory != jobDir) { Info("有任务正在运行", "请等待当前任务完成，或先取消，再切换到其他任务。"); return; }
            try { await Attach(selected.Directory); Find<RadioButton>("MergeNav").IsChecked = true; } catch (Exception ex) { if (!closed) Error(ex); }
        }
        public void SavePreview(string path, double dpi) {
            Window.UpdateLayout(); var view = Find<Grid>("RootGrid"); int width = (int)Math.Ceiling(view.ActualWidth * dpi / 96), height = (int)Math.Ceiling(view.ActualHeight * dpi / 96);
            var bitmap = new RenderTargetBitmap(width, height, dpi, dpi, PixelFormats.Pbgra32); bitmap.Render(view); var encoder = new PngBitmapEncoder(); encoder.Frames.Add(BitmapFrame.Create(bitmap)); using (var file = File.Create(path)) encoder.Save(file);
        }
    }
}
