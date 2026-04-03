using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

using System.Windows.Interop;

namespace LinHPMonitor
{
    public partial class MainWindow : Window
    {
        [DllImport("user32.dll")]
        public static extern bool SetWindowDisplayAffinity(IntPtr hWnd, uint dwAffinity);
        private const uint WDA_EXCLUDEFROMCAPTURE = 0x11;

        private HPMonitor? _monitor;
        private CancellationTokenSource? _cts;
        private bool _isRunning = false;
        private MonitorConfig _config;
        // Keypress queues and workers for HP/MP
        private System.Collections.Concurrent.BlockingCollection<ushort>? _hpKeyQueue;
        private System.Collections.Concurrent.BlockingCollection<ushort>? _mpKeyQueue;
        private Task? _hpKeyWorkerTask;
        private Task? _mpKeyWorkerTask;
        // 上次顯示的 current 值，用來抑制短期抖動
        private int _lastDisplayedHpCur = -1;
        private int _lastDisplayedMpCur = -1;
        // 現在使用設定檔中的開關來控制是否跳過校準（UI 可切換）

        public MainWindow()
        {
            InitializeComponent();
            _config = ConfigManager.Load();

            // Set Capture Invisibility
            try {
                var hwnd = new System.Windows.Interop.WindowInteropHelper(this).EnsureHandle();
                SetWindowDisplayAffinity(hwnd, WDA_EXCLUDEFROMCAPTURE);
            } catch { /* Silent if platform doesn't support */ }

            try {
                ApplyConfig();
                RefreshWindowList();
                StartMouseTracking(); // V4.1: 啟動滑鼠座標追蹤
                Log("系統啟動成功 V3.9.7 (強力解析版)。");
            } catch (Exception ex) {
                File.AppendAllText("debug_log.txt", $"[INIT ERROR] {ex}\n");
                MessageBox.Show($"初始化異常: {ex.Message}");
            }
        }

        // V4.1: 新增滑鼠位置監聽定時器
        private void StartMouseTracking()
        {
            var timer = new DispatcherTimer();
            timer.Interval = TimeSpan.FromMilliseconds(50);
            timer.Tick += (s, e) =>
            {
                if (Win32Input.GetCursorPos(out var point))
                {
                    try
                    {
                        // 優先使用設定檔中的 WindowTitle 找到目標視窗
                        var hwnd = Win32Input.FindWindowPartial(_config.WindowTitle);
                        if (hwnd != IntPtr.Zero)
                        {
                            var rect = Win32Input.GetActualWindowRect(hwnd);
                            int relX = point.X - rect.Left;
                            int relY = point.Y - rect.Top;
                            int w = rect.Right - rect.Left;
                            int h = rect.Bottom - rect.Top;
                            if (relX >= 0 && relY >= 0 && relX <= w && relY <= h)
                            {
                                MousePosTextBlock.Text = $"{relX}, {relY} (rel)";
                            }
                            else
                            {
                                // 游標不在目標視窗內，顯示相對座標並標註
                                MousePosTextBlock.Text = $"{relX}, {relY} (outside)";
                            }
                        }
                        else
                        {
                            // 若找不到目標視窗，回退到全螢幕座標
                            MousePosTextBlock.Text = $"{point.X}, {point.Y}";
                        }
                    }
                    catch
                    {
                        MousePosTextBlock.Text = $"{point.X}, {point.Y}";
                    }
                }
            };
            timer.Start();
        }

        private void ApplyConfig()
        {
            HPThresholdTextBox.Text = _config.HPThreshold.ToString();
            MPThresholdTextBox.Text = _config.MPThreshold.ToString();
            MaxHPTextBox.Text = _config.UserMaxHp.ToString();
            MaxMPTextBox.Text = _config.UserMaxMp.ToString();
            PressCountTextBox.Text = _config.PressCount.ToString();
            SetComboBoxItem(HPHotkeyComboBox, _config.HPHotkey);
            SetComboBoxItem(MPHotkeyComboBox, _config.MPHotkey);
            // 同步預覽開關
            SkipCalibrationCheckBox.IsChecked = _config.SkipCalibrationForPreview;
            Log("已載入設定檔值。");
        }

        private void SkipCalibrationCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            bool isChecked = SkipCalibrationCheckBox.IsChecked == true;
            _config.SkipCalibrationForPreview = isChecked;
            ConfigManager.Save(_config);
            Log($"SkipCalibrationForPreview 設定為: {isChecked}");
        }

        private void SetComboBoxItem(System.Windows.Controls.ComboBox comboBox, string value)
        {
            foreach (System.Windows.Controls.ComboBoxItem item in comboBox.Items)
            {
                if (item.Content.ToString() == value) { comboBox.SelectedItem = item; break; }
            }
        }

        private void Log(string message)
        {
            Dispatcher.Invoke(() => {
                LogTextBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
                LogTextBox.ScrollToEnd();
                File.AppendAllText("debug_log.txt", $"[{DateTime.Now:HH:mm:ss}] {message}\n");
            });
        }

        private void RefreshWindowList()
        {
            WindowComboBox.Items.Clear();
            var keywords = new[] { "Lineage", "PURPLE" };
            var processes = Process.GetProcesses()
                .Where(p => !string.IsNullOrEmpty(p.MainWindowTitle) && 
                            keywords.Any(k => p.MainWindowTitle.Contains(k, StringComparison.OrdinalIgnoreCase)))
                .OrderBy(p => p.MainWindowTitle);

            foreach (var p in processes) WindowComboBox.Items.Add(p.MainWindowTitle);
            
            if (WindowComboBox.Items.Contains(_config.WindowTitle)) WindowComboBox.Text = _config.WindowTitle;
            else if (WindowComboBox.Items.Count > 0) WindowComboBox.SelectedIndex = 0;
        }

        private void RefreshButton_Click(object sender, RoutedEventArgs e) => RefreshWindowList();

        private async void MainActionButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isRunning) { 
                StopMonitoring(); 
                return; 
            }
            
            try {
                if (string.IsNullOrEmpty(WindowComboBox.Text)) {
                    MessageBox.Show("請先選擇遊戲視窗！");
                    return;
                }
                
                 Log("正在載入 OCR 引擎...");
                if (_monitor == null) {
                    try {
                        _monitor = new HPMonitor();
                    } catch (Exception ocrEx) {
                        Log($"[OCR 錯誤] {ocrEx.Message}");
                        return; // 終止開始
                    }
                }
                // 若設定檔中有儲存區域，嘗試將其套用到 Monitor。
                // 支援絕對座標（HPRegionX/HPRegionY）或正規化比例值（HPRegionNormX/HPRegionNormY + RegionWidthRatio）
                try {
                    IntPtr hwnd = Win32Input.FindWindowPartial(_config.WindowTitle);
                    if (hwnd != IntPtr.Zero)
                    {
                        var rect = Win32Input.GetActualWindowRect(hwnd);
                        int winW = rect.Right - rect.Left;
                        int winH = rect.Bottom - rect.Top;

                        int w = _config.RegionWidth > 0 ? _config.RegionWidth : Math.Max(140, (int)(winW * 0.135));
                        int h = _config.RegionHeight > 0 ? _config.RegionHeight : Math.Max(22, (int)(winH * 0.037));

                        System.Drawing.Rectangle hpRect, mpRect;

                        if (_config.RegionWidthRatio > 0 && _config.HPRegionNormX > 0)
                        {
                            int centerX = rect.Left + (int)(_config.HPRegionNormX * winW);
                            int centerY = rect.Top + (int)(_config.HPRegionNormY * winH);
                            int hpX = centerX - w / 2;
                            int hpY = centerY - h / 2;
                            hpX = Math.Max(0, Math.Min(hpX, winW - w));
                            hpY = Math.Max(0, Math.Min(hpY, winH - h));
                            hpRect = new System.Drawing.Rectangle(hpX, hpY, w, h);
                        }
                        else if (_config.HPRegionX != 0 || _config.HPRegionY != 0)
                        {
                            int hpX = _config.HPRegionX - w / 2;
                            int hpY = _config.HPRegionY - h / 2;
                            if (hpX < 0) hpX = _config.HPRegionX;
                            if (hpY < 0) hpY = _config.HPRegionY;
                            hpRect = new System.Drawing.Rectangle(hpX, hpY, w, h);
                        }
                        else
                        {
                            // fallback to internal defaults
                            hpRect = new System.Drawing.Rectangle(_monitor.HPRegion.X, _monitor.HPRegion.Y, _monitor.HPRegion.Width, _monitor.HPRegion.Height);
                        }

                        if (_config.RegionWidthRatio > 0 && _config.MPRegionNormX > 0)
                        {
                            int centerX = rect.Left + (int)(_config.MPRegionNormX * winW);
                            int centerY = rect.Top + (int)(_config.MPRegionNormY * winH);
                            int mpX = centerX - w / 2;
                            int mpY = centerY - h / 2;
                            mpX = Math.Max(0, Math.Min(mpX, winW - w));
                            mpY = Math.Max(0, Math.Min(mpY, winH - h));
                            mpRect = new System.Drawing.Rectangle(mpX, mpY, w, h);
                        }
                        else if (_config.MPRegionX != 0 || _config.MPRegionY != 0)
                        {
                            int mpX = _config.MPRegionX - w / 2;
                            int mpY = _config.MPRegionY - h / 2;
                            if (mpX < 0) mpX = _config.MPRegionX;
                            if (mpY < 0) mpY = _config.MPRegionY;
                            mpRect = new System.Drawing.Rectangle(mpX, mpY, w, h);
                        }
                        else
                        {
                            mpRect = new System.Drawing.Rectangle(_monitor.MPRegion.X, _monitor.MPRegion.Y, _monitor.MPRegion.Width, _monitor.MPRegion.Height);
                        }

                        _monitor.SetRegions(hpRect, mpRect);
                        Log($"已套用儲存區域：HP({hpRect.X},{hpRect.Y}) MP({mpRect.X},{mpRect.Y})");
                    }
                }
                catch (Exception ex) { Log($"套用儲存區域失敗: {ex.Message}"); }
                
                _config.WindowTitle = WindowComboBox.Text;
                if (double.TryParse(HPThresholdTextBox.Text, out double hpT)) _config.HPThreshold = hpT;
                if (double.TryParse(MPThresholdTextBox.Text, out double mpT)) _config.MPThreshold = mpT;
                if (int.TryParse(MaxHPTextBox.Text, out int maxHp)) _config.UserMaxHp = maxHp;
                if (int.TryParse(MaxMPTextBox.Text, out int maxMp)) _config.UserMaxMp = maxMp;
                if (int.TryParse(PressCountTextBox.Text, out int pc) && pc >= 1 && pc <= 5) _config.PressCount = pc;
                _config.HPHotkey = (HPHotkeyComboBox.SelectedItem as System.Windows.Controls.ComboBoxItem)?.Content.ToString() ?? "F5";
                _config.MPHotkey = (MPHotkeyComboBox.SelectedItem as System.Windows.Controls.ComboBoxItem)?.Content.ToString() ?? "F6";
                ConfigManager.Save(_config);

                _isRunning = true;
                _cts = new CancellationTokenSource();
                MainActionButton.Content = "🛑 停止監測";
                MainActionButton.Background = System.Windows.Media.Brushes.Tomato;
                // 建立按鍵佇列與工作者（獨立執行緒處理按鍵），使偵測與按鍵互不干擾
                _hpKeyQueue = new System.Collections.Concurrent.BlockingCollection<ushort>(new System.Collections.Concurrent.ConcurrentQueue<ushort>());
                _mpKeyQueue = new System.Collections.Concurrent.BlockingCollection<ushort>(new System.Collections.Concurrent.ConcurrentQueue<ushort>());
                _hpKeyWorkerTask = Task.Run(() => KeyWorkerLoop(_hpKeyQueue, _cts.Token));
                _mpKeyWorkerTask = Task.Run(() => KeyWorkerLoop(_mpKeyQueue, _cts.Token));

                if (_config.SkipCalibrationForPreview)
                {
                    Log("[預覽模式] 已跳過自動位置偵測，使用程式內預設或先前設定的 HP/MP 區域。");
                    // 將目前 Monitor 的區域同步回設定檔，並儲存
                    try
                    {
                        if (_monitor != null)
                        {
                            _config.HPRegionX = _monitor.HPRegion.X;
                            _config.HPRegionY = _monitor.HPRegion.Y;
                            _config.MPRegionX = _monitor.MPRegion.X;
                            _config.MPRegionY = _monitor.MPRegion.Y;
                            _config.RegionWidth = _monitor.HPRegion.Width;
                            _config.RegionHeight = _monitor.HPRegion.Height;
                            ConfigManager.Save(_config);
                            Log("已將預覽區域同步至設定檔。");
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"同步預覽區域失敗: {ex.Message}");
                    }
                }
                else
                {
                    Log("執行自動位置偵測...");
                    CalibrationProgressBar.Visibility = System.Windows.Visibility.Visible;
                    CalibrationProgressBar.Value = 0;

                    // V3.9.1: 使用 Task.Run 將耗時的 OCR 定位放在背景，解放 UI 緒以利進度條更新
                    string calib = await Task.Run(() => 
                        _monitor.AutoCalibrate(_config.WindowTitle, (val, status) => {
                            Dispatcher.Invoke(() => {
                                CalibrationProgressBar.Value = val;
                            });
                        })
                    );

                    Log(calib);
                    CalibrationProgressBar.Visibility = System.Windows.Visibility.Collapsed;

                    // 校準完成後立即擷取 HP/MP 區域預覽，讓使用者確認校準位置是否正確
                    try
                    {
                        var (hpRatio, hpBmp) = await Task.Run(() => _monitor.GetQuickRatio(_config.WindowTitle, false));
                        var (mpRatio, mpBmp) = await Task.Run(() => _monitor.GetQuickRatio(_config.WindowTitle, true));
                        if (hpBmp != null)
                        {
                            HPPreviewImage.Source = BitmapToImageSource(hpBmp);
                            HPPreviewTextBlock.Text = $"[校準預覽] 填充率 {hpRatio:P0}";
                        }
                        if (mpBmp != null)
                        {
                            MPPreviewImage.Source = BitmapToImageSource(mpBmp);
                            MPPreviewTextBlock.Text = $"[校準預覽] 填充率 {mpRatio:P0}";
                        }
                    }
                    catch (Exception prevEx) { Log($"[預覽擷取失敗] {prevEx.Message}"); }

                    // 將校準結果轉為正規化座標並儲存到設定檔，方便跨視窗大小自動套用
                    try
                    {
                        IntPtr hwnd = Win32Input.FindWindowPartial(_config.WindowTitle);
                        if (hwnd != IntPtr.Zero)
                        {
                            var rect = Win32Input.GetActualWindowRect(hwnd);
                            int winW = rect.Right - rect.Left;
                            int winH = rect.Bottom - rect.Top;
                            if (winW > 0 && winH > 0)
                            {
                                var hp = _monitor.HPRegion;
                                var mp = _monitor.MPRegion;
                                _config.RegionWidth = hp.Width;
                                _config.RegionHeight = hp.Height;
                                _config.RegionWidthRatio = Math.Round((double)hp.Width / winW, 4);
                                _config.RegionHeightRatio = Math.Round((double)hp.Height / winH, 4);
                                _config.HPRegionNormX = Math.Round((double)(hp.X + hp.Width / 2) / winW, 4);
                                _config.HPRegionNormY = Math.Round((double)(hp.Y + hp.Height / 2) / winH, 4);
                                _config.MPRegionNormX = Math.Round((double)(mp.X + mp.Width / 2) / winW, 4);
                                _config.MPRegionNormY = Math.Round((double)(mp.Y + mp.Height / 2) / winH, 4);
                                ConfigManager.Save(_config);
                                Log($"已儲存正規化校準資料（相對視窗）：HP({_config.HPRegionNormX:F3},{_config.HPRegionNormY:F3}) MP({_config.MPRegionNormX:F3},{_config.MPRegionNormY:F3}) 尺寸比: {_config.RegionWidthRatio:F3}x{_config.RegionHeightRatio:F3}");
                            }
                        }
                    }
                    catch (Exception normEx) { Log($"[儲存校準比例失敗] {normEx.Message}"); }
                }

                Log("監測器啟動成功。");
                // 啟動兩個獨立任務分別處理 HP 與 MP，互不干擾
                var hpTask = Task.Run(() => RunHpLoop(_cts.Token));
                var mpTask = Task.Run(() => RunMpLoop(_cts.Token));
                await Task.WhenAll(hpTask, mpTask);
            } catch (Exception ex) {
                Log($"啟動失敗：{ex.Message}");
                File.AppendAllText("debug_log.txt", $"[START ERROR] {ex}\n");
                StopMonitoring();
            }
        }

        // 獨立處理 HP 的背景迴圈
        private void RunHpLoop(CancellationToken token)
        {
            string title = _config.WindowTitle;
            ushort hpVk = GetVkFromHotkey(_config.HPHotkey);
            while (!token.IsCancellationRequested)
            {
                try
                {
                    double displayHpRatio;
                    int displayHpCur, displayHpMax;
                    string hpOcrText = "";
                    Bitmap? hpPreview = null;

                    if (_config.UserMaxHp > 0)
                    {
                        // 已設定最大値：把血條切成 UserMaxHp 格，數有色格子 = 當前血量
                        var seg = _monitor?.GetCurrentBySegments(title, false, _config.UserMaxHp) ?? (0, 0.0, null);
                        displayHpMax = _config.UserMaxHp;
                        displayHpCur = seg.current;
                        displayHpRatio = seg.fillRatio;
                        hpPreview = seg.preview;
                    }
                    else
                    {
                        var hp = _monitor?.DetectHP(title) ?? (0, 0, 1.0, 0.0, null, "");
                        displayHpCur = hp.current; displayHpMax = hp.max; displayHpRatio = hp.ratio;
                        hpPreview = hp.processed;
                        hpOcrText = hp.rawText;
                        if (hp.max > 0)
                        {
                            displayHpCur = (hp.current > 0 && hp.current <= hp.max)
                                ? hp.current
                                : (hp.fillRatio > 0.01 ? (int)Math.Round(hp.fillRatio * hp.max) : displayHpCur);
                            displayHpRatio = displayHpMax > 0 ? (double)displayHpCur / displayHpMax : hp.ratio;
                        }
                        else if (hp.fillRatio > 0.01 && _lastDisplayedHpCur > 0)
                        {
                            displayHpCur = (int)Math.Round(hp.fillRatio * _lastDisplayedHpCur);
                            displayHpRatio = hp.fillRatio;
                        }
                    }

                    // 初始化快取
                    if (_lastDisplayedHpCur < 0) _lastDisplayedHpCur = displayHpCur;

                    // 黏性滿值：上次滿血且本次跌幅 ≤ 3 格，視為雜訊維持滿值
                    if (_lastDisplayedHpCur == displayHpMax && displayHpCur >= displayHpMax - 3)
                        displayHpCur = displayHpMax;
                    // 一般抖動壓制：變化量 ≤ 3 格時保留舊值（邊緣偵測精度約 ±1~2 格）
                    else if (Math.Abs(displayHpCur - _lastDisplayedHpCur) <= 3)
                        displayHpCur = _lastDisplayedHpCur;
                    else
                        _lastDisplayedHpCur = displayHpCur;

                    Bitmap? capturedHpPreview = hpPreview;
                    Dispatcher.Invoke(() => {
                        UpdateStatusDisplay(true, displayHpCur, displayHpMax, displayHpRatio);
                        if (capturedHpPreview != null) HPPreviewImage.Source = BitmapToImageSource(capturedHpPreview);
                        if (!string.IsNullOrEmpty(hpOcrText))
                            HPPreviewTextBlock.Text = $"HP OCR: {hpOcrText}";
                    });
                    if (!string.IsNullOrEmpty(hpOcrText)) Log($"[HP OCR]: {hpOcrText}");

                    if (displayHpMax > 0 && displayHpRatio < _config.HPThreshold / 100.0)
                    {
                        Log($"[HP] {displayHpRatio:P0} 低於閥值，排程鍵盤碼 {hpVk:X}");
                        try { _hpKeyQueue?.Add(hpVk); } catch { }
                        Thread.Sleep(200);
                    }
                }
                catch { }
                Thread.Sleep(_config.CheckInterval);
            }
        }

        // 獨立處理 MP 的背景迴圈
        private void RunMpLoop(CancellationToken token)
        {
            string title = _config.WindowTitle;
            ushort mpVk = GetVkFromHotkey(_config.MPHotkey);
            while (!token.IsCancellationRequested)
            {
                try
                {
                    double displayMpRatio;
                    int displayMpCur, displayMpMax;
                    string mpOcrText = "";
                    Bitmap? mpPreview = null;

                    if (_config.UserMaxMp > 0)
                    {
                        // 已設定最大値：把魔條切成 UserMaxMp 格，數有色格子 = 當前魔量
                        var seg = _monitor?.GetCurrentBySegments(title, true, _config.UserMaxMp) ?? (0, 0.0, null);
                        displayMpMax = _config.UserMaxMp;
                        displayMpCur = seg.current;
                        displayMpRatio = seg.fillRatio;
                        mpPreview = seg.preview;
                    }
                    else
                    {
                        var mp = _monitor?.DetectMP(title) ?? (0, 0, 1.0, 0.0, null, "");
                        displayMpCur = mp.current; displayMpMax = mp.max; displayMpRatio = mp.ratio;
                        mpPreview = mp.processed;
                        mpOcrText = mp.rawText;
                        if (mp.max > 0)
                        {
                            displayMpCur = (mp.current > 0 && mp.current <= mp.max)
                                ? mp.current
                                : (mp.fillRatio > 0.01 ? (int)Math.Round(mp.fillRatio * mp.max) : displayMpCur);
                            displayMpRatio = displayMpMax > 0 ? (double)displayMpCur / displayMpMax : mp.ratio;
                        }
                        else if (mp.fillRatio > 0.01 && _lastDisplayedMpCur > 0)
                        {
                            displayMpCur = (int)Math.Round(mp.fillRatio * _lastDisplayedMpCur);
                            displayMpRatio = mp.fillRatio;
                        }
                    }

                    // 抑制過度抖動：差異 ≤ 2 時保留原值
                    // 初始化快取
                    if (_lastDisplayedMpCur < 0) _lastDisplayedMpCur = displayMpCur;

                    // 黏性滿值：上次滿魔且本次跌幅 ≤ 3 格，視為雜訊維持滿值
                    if (_lastDisplayedMpCur == displayMpMax && displayMpCur >= displayMpMax - 3)
                        displayMpCur = displayMpMax;
                    // 一般抖動壓制：變化量 ≤ 3 格時保留舊值
                    else if (Math.Abs(displayMpCur - _lastDisplayedMpCur) <= 3)
                        displayMpCur = _lastDisplayedMpCur;
                    else
                        _lastDisplayedMpCur = displayMpCur;

                    Bitmap? capturedMpPreview = mpPreview;
                    Dispatcher.Invoke(() => {
                        UpdateStatusDisplay(false, displayMpCur, displayMpMax, displayMpRatio);
                        if (capturedMpPreview != null) MPPreviewImage.Source = BitmapToImageSource(capturedMpPreview);
                        if (!string.IsNullOrEmpty(mpOcrText))
                            MPPreviewTextBlock.Text = $"MP OCR: {mpOcrText}";
                    });
                    if (!string.IsNullOrEmpty(mpOcrText)) Log($"[MP OCR]: {mpOcrText}");

                    if (displayMpMax > 0 && displayMpRatio < _config.MPThreshold / 100.0)
                    {
                        Log($"[MP] {displayMpRatio:P0} 低於閥值，排程鍵盤碼 {mpVk:X}");
                        try { _mpKeyQueue?.Add(mpVk); } catch { }
                        Thread.Sleep(200);
                    }
                }
                catch { }
                Thread.Sleep(_config.CheckInterval);
            }
        }

        // Key worker loop: 消費佇列並實際發送鍵盤按鍵
        private void KeyWorkerLoop(System.Collections.Concurrent.BlockingCollection<ushort>? queue, CancellationToken token)
        {
            if (queue == null) return;
            try
            {
                foreach (var vk in queue.GetConsumingEnumerable(token))
                {
                    try
                    {
                        int pressCount = Math.Max(1, _config.PressCount);
                        for (int i = 0; i < pressCount; i++)
                        {
                            Win32Input.SendKeyPress(vk, _config.KeyPressDuration);
                            // 連按間隔 150ms（讓遊戲有時間辨識下一次鍵擊）
                            if (i < pressCount - 1) Thread.Sleep(150);
                        }
                    }
                    catch { }
                    Thread.Sleep(50);
                }
            }
            catch (OperationCanceledException) { }
            catch { }
        }

        private void StopMonitoring()
        {
            _cts?.Cancel();
            // 停止佇列輸入並讓工作者結束
            try { _hpKeyQueue?.CompleteAdding(); } catch { }
            try { _mpKeyQueue?.CompleteAdding(); } catch { }
            _isRunning = false;
            MainActionButton.Content = "🚀 開始監測";
            MainActionButton.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(187, 134, 252));
            Log("監測已手動停止。");
        }

        private void RunLoop(CancellationToken token)
        {
            string title = _config.WindowTitle;
            ushort hpVk = GetVkFromHotkey(_config.HPHotkey);
            ushort mpVk = GetVkFromHotkey(_config.MPHotkey);

            while (!token.IsCancellationRequested)
            {
                try {
                    var hp = _monitor?.DetectHP(title) ?? (0, 0, 1.0, 0.0, null, "");
                    // 優先以顏色填充比例計算當前值（以手動最大值或解析到的 max 為基準）
                    int displayHpCur = hp.current;
                    int displayHpMax = hp.max;
                    double displayHpRatio = hp.ratio;
                    if (_config.UserMaxHp > 0)
                    {
                        displayHpMax = _config.UserMaxHp;
                        displayHpCur = (int)Math.Round(hp.fillRatio * displayHpMax);
                        displayHpRatio = displayHpMax > 0 ? (double)displayHpCur / displayHpMax : hp.ratio;
                    }
                    else if (hp.max > 0 && hp.fillRatio > 0.01)
                    {
                        displayHpCur = (int)Math.Round(hp.fillRatio * hp.max);
                        displayHpRatio = hp.max > 0 ? (double)displayHpCur / hp.max : hp.ratio;
                    }

                    Dispatcher.Invoke(() => {
                        UpdateStatusDisplay(true, displayHpCur, displayHpMax, displayHpRatio);
                        if (hp.processed != null) HPPreviewImage.Source = BitmapToImageSource(hp.processed);
                    });

                    // V3.5: 即時輸出全量的 OCR 字串到 Log 區供使用者觀測
                    if (!string.IsNullOrEmpty(hp.rawText)) {
                        Log($"[HP OCR]: {hp.rawText}");
                    }

                    if (hp.max > 0 && hp.ratio < _config.HPThreshold / 100.0) {
                        Log($"[HP] {hp.ratio:P0} 低於閥值，送出鍵盤碼 {hpVk:X}");
                        Win32Input.SendKeyPress(hpVk, 50);
                        Thread.Sleep(200);
                        continue;
                    }

                    var mp = _monitor?.DetectMP(title) ?? (0, 0, 1.0, 0.0, null, "");
                    int displayMpCur = mp.current;
                    int displayMpMax = mp.max;
                    double displayMpRatio = mp.ratio;
                    if (_config.UserMaxMp > 0)
                    {
                        displayMpMax = _config.UserMaxMp;
                        displayMpCur = (int)Math.Round(mp.fillRatio * displayMpMax);
                        displayMpRatio = displayMpMax > 0 ? (double)displayMpCur / displayMpMax : mp.ratio;
                    }
                    else if (mp.max > 0 && mp.fillRatio > 0.01)
                    {
                        displayMpCur = (int)Math.Round(mp.fillRatio * mp.max);
                        displayMpRatio = mp.max > 0 ? (double)displayMpCur / mp.max : mp.ratio;
                    }

                    Dispatcher.Invoke(() => {
                        UpdateStatusDisplay(false, displayMpCur, displayMpMax, displayMpRatio);
                        if (mp.processed != null) MPPreviewImage.Source = BitmapToImageSource(mp.processed);
                        // 顯示 OCR 字串於預覽下方，並寫入 log
                        HPPreviewTextBlock.Text = string.IsNullOrEmpty(hp.rawText) ? "HP OCR:" : $"HP OCR: {hp.rawText}";
                        MPPreviewTextBlock.Text = string.IsNullOrEmpty(mp.rawText) ? "MP OCR:" : $"MP OCR: {mp.rawText}";
                    });

                    if (!string.IsNullOrEmpty(mp.rawText)) {
                        Log($"[MP OCR]: {mp.rawText}");
                    }

                    if (mp.max > 0 && mp.ratio < _config.MPThreshold / 100.0) {
                        Log($"[MP] {mp.ratio:P0} 低於閥值，送出鍵盤碼 {mpVk:X}");
                        Win32Input.SendKeyPress(mpVk, 50);
                        Thread.Sleep(200);
                    }
                } catch { /* Suppress runtime capture errors to keep loop alive */ }
                Thread.Sleep(300);
            }
        }

        private void UpdateStatusDisplay(bool isHp, int current, int max, double ratio)
        {
            if (isHp) {
                HPValueTextBlock.Text = $"HP: {current} / {max}";
                HPProgressBar.Value = Math.Max(0, Math.Min(100, ratio * 100));
            } else {
                MPValueTextBlock.Text = $"MP: {current} / {max}";
                MPProgressBar.Value = Math.Max(0, Math.Min(100, ratio * 100));
            }
        }

        private ImageSource BitmapToImageSource(Bitmap bitmap)
        {
            using (MemoryStream memory = new MemoryStream()) {
                bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Bmp);
                memory.Position = 0;
                BitmapImage bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = memory;
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();
                bitmapImage.Freeze();
                return bitmapImage;
            }
        }

        private ushort GetVkFromHotkey(string hotkey)
        {
            return hotkey switch {
                "F5" => 0x74, "F6" => 0x75, "F7" => 0x76, "F8" => 0x77,
                "F9" => 0x78, "F10" => 0x79, "F11" => 0x7A, "F12" => 0x7B,
                _ => 0x74
            };
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            _cts?.Cancel();
            base.OnClosing(e);
        }
    }
}