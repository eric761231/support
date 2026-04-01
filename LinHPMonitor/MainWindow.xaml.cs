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
                Log("系統啟動成功 V3.9.7 (強力解析版)。");
            } catch (Exception ex) {
                File.AppendAllText("debug_log.txt", $"[INIT ERROR] {ex}\n");
                MessageBox.Show($"初始化異常: {ex.Message}");
            }
        }

        private void ApplyConfig()
        {
            HPThresholdTextBox.Text = _config.HPThreshold.ToString();
            MPThresholdTextBox.Text = _config.MPThreshold.ToString();
            SetComboBoxItem(HPHotkeyComboBox, _config.HPHotkey);
            SetComboBoxItem(MPHotkeyComboBox, _config.MPHotkey);
            Log("已載入設定檔值。");
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
            var processes = Process.GetProcesses().Where(p => !string.IsNullOrEmpty(p.MainWindowTitle)).OrderBy(p => p.MainWindowTitle);
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
                
                _config.WindowTitle = WindowComboBox.Text;
                if (double.TryParse(HPThresholdTextBox.Text, out double hpT)) _config.HPThreshold = hpT;
                if (double.TryParse(MPThresholdTextBox.Text, out double mpT)) _config.MPThreshold = mpT;
                _config.HPHotkey = (HPHotkeyComboBox.SelectedItem as System.Windows.Controls.ComboBoxItem)?.Content.ToString() ?? "F5";
                _config.MPHotkey = (MPHotkeyComboBox.SelectedItem as System.Windows.Controls.ComboBoxItem)?.Content.ToString() ?? "F6";
                ConfigManager.Save(_config);

                _isRunning = true;
                _cts = new CancellationTokenSource();
                MainActionButton.Content = "🛑 停止監測";
                MainActionButton.Background = System.Windows.Media.Brushes.Tomato;
                
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

                Log("監測器啟動成功。");
                await Task.Run(() => RunLoop(_cts.Token));
            } catch (Exception ex) {
                Log($"啟動失敗：{ex.Message}");
                File.AppendAllText("debug_log.txt", $"[START ERROR] {ex}\n");
                StopMonitoring();
            }
        }

        private void StopMonitoring()
        {
            _cts?.Cancel();
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
                    var hp = _monitor?.DetectHP(title) ?? (0, 0, 1.0, null, "");
                    Dispatcher.Invoke(() => {
                        UpdateStatusDisplay(true, hp.current, hp.max, hp.ratio);
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

                    var mp = _monitor?.DetectMP(title) ?? (0, 0, 1.0, null, "");
                    Dispatcher.Invoke(() => {
                        UpdateStatusDisplay(false, mp.current, mp.max, mp.ratio);
                        if (mp.processed != null) MPPreviewImage.Source = BitmapToImageSource(mp.processed);
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