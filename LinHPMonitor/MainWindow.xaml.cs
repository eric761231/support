using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Runtime.InteropServices;

namespace LinHPMonitor
{
    public partial class MainWindow : Window
    {
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string? lpClassName, string lpWindowName);

        private HPMonitor? _monitor;
        private CancellationTokenSource? _cts;
        private bool _isRunning = false;

        public MainWindow()
        {
            InitializeComponent();
            _monitor = new HPMonitor();
            Log("應用程式初始化完成。");
        }

        private void Log(string message)
        {
            Dispatcher.Invoke(() =>
            {
                LogTextBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
                LogTextBox.ScrollToEnd();
            });
        }

        private async void StartButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isRunning) return;

            string windowTitle = WindowTitleTextBox?.Text ?? string.Empty;
            
            if (!double.TryParse(ThresholdTextBox?.Text, out double threshold))
            {
                MessageBox.Show("請輸入有效的補血百分比。");
                return;
            }

            int duration = int.Parse(DurationTextBox.Text);
            int interval = int.Parse(IntervalTextBox.Text);
            ushort vkCode = GetVkFromHotkey(HotkeyComboBox.Text);

            _isRunning = true;
            _cts = new CancellationTokenSource();
            StartButton.IsEnabled = false;
            StopButton.IsEnabled = true;

            Log($"開始監測視窗：'{windowTitle}'...");

            try
            {
                await Task.Run(() => RunLoop(windowTitle, threshold / 100.0, vkCode, duration, interval, _cts.Token));
            }
            catch (OperationCanceledException)
            {
                Log("監測已停止。");
            }
            finally
            {
                _isRunning = false;
                StartButton.IsEnabled = true;
                StopButton.IsEnabled = false;
            }
        }

        private void StopButton_Click(object sender, RoutedEventArgs e)
        {
            _cts?.Cancel();
        }

        private void RunLoop(string windowTitle, double threshold, ushort vkCode, int duration, int interval, CancellationToken token)
        {
            while (!token.IsCancellationRequested)
            {
                var hp = _monitor?.DetectHP(windowTitle) ?? (0, 0, 1.0);
                
                Dispatcher.Invoke(() => UpdateHPDisplay(hp.current, hp.max, hp.ratio));

                if (hp.max > 0 && hp.ratio < threshold)
                {
                    Log($"偵測到低體力 ({hp.ratio:P1})，正在發送補血指令...");
                    Win32Input.SendKeyPress(vkCode, duration);
                    Thread.Sleep(interval);
                }

                Thread.Sleep(500); // 2 samples per second
            }
        }

        private void UpdateHPDisplay(int current, int max, double ratio)
        {
            HPValueTextBlock.Text = $"{current} / {max} ({ratio:P1})";
            HPProgressBar.Value = ratio * 100;

            if (ratio < 0.3) HPProgressBar.Foreground = System.Windows.Media.Brushes.Red;
            else if (ratio < 0.7) HPProgressBar.Foreground = System.Windows.Media.Brushes.Orange;
            else HPProgressBar.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(3, 218, 198));
        }

        private ushort GetVkFromHotkey(string hotkey)
        {
            return hotkey switch
            {
                "F5" => 0x74,
                "F6" => 0x75,
                "F7" => 0x76,
                "F8" => 0x77,
                "F9" => 0x78,
                "F10" => 0x79,
                "F11" => 0x7A,
                "F12" => 0x7B,
                _ => 0x74
            };
        }
    }
}