using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using Tesseract;

namespace LinHPMonitor
{
    public class HPMonitor
    {
        private TesseractEngine? _ocr;
        private int _lastMaxHp = 0;

        public HPMonitor(string? tessDataPath = null)
        {
            if (string.IsNullOrEmpty(tessDataPath))
            {
                tessDataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");
            }
            
            if (!Directory.Exists(tessDataPath))
            {
                Directory.CreateDirectory(tessDataPath);
            }

            try
            {
                _ocr = new TesseractEngine(tessDataPath, "eng", EngineMode.Default);
                _ocr.SetVariable("tessedit_char_whitelist", "0123456789/HP: ");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"OCR Initialization Error: {ex.Message}");
            }
        }

        public (int current, int max, double ratio) DetectHP(string windowTitle, Rectangle? region = null)
        {
            IntPtr hwnd = FindWindow(null, windowTitle);
            if (hwnd == IntPtr.Zero) return (0, 0, 1.0);

            Win32Input.GetWindowRect(hwnd, out Win32Input.RECT rect);
            int x = rect.Left + (region?.X ?? 0);
            int y = rect.Top + (region?.Y ?? 0);
            int width = region?.Width ?? (rect.Right - rect.Left);
            int height = region?.Height ?? (rect.Bottom - rect.Top);

            using (Bitmap screenshot = CaptureRegion(x, y, width, height))
            using (Mat mat = screenshot.ToMat())
            {
                using (Mat processed = PreprocessHPImage(mat))
                using (Bitmap processedBitmap = processed.ToBitmap())
                using (var img = PixConverter.ToPix(processedBitmap))
                {
                    if (_ocr == null) return (0, 0, 1.0);
                    using (var page = _ocr.Process(img, PageSegMode.SparseText))
                    {
                        string text = page.GetText().Trim();
                        text = PostProcessOCRText(text);

                        return ParseHPText(text);
                    }
                }
            }
        }

        private Mat PreprocessHPImage(Mat src)
        {
            Mat gray = new Mat();
            Cv2.CvtColor(src, gray, ColorConversionCodes.BGR2GRAY);

            // Rescale 3x (like the Python version)
            Mat resized = new Mat();
            Cv2.Resize(gray, resized, new OpenCvSharp.Size(gray.Width * 3, gray.Height * 3), 0, 0, InterpolationFlags.Cubic);

            // Otsu Binarization
            Mat thresh = new Mat();
            Cv2.Threshold(resized, thresh, 0, 255, ThresholdTypes.BinaryInv | ThresholdTypes.Otsu);

            return thresh;
        }

        private string PostProcessOCRText(string text)
        {
            // Translation table for common OCR errors in Lineage font
            var map = new (string from, string to)[]
            {
                ("I", "1"), ("l", "1"), ("|", "1"),
                ("G", "6"), ("b", "6"),
                ("T", "7"), ("t", "7"),
                ("S", "5"), ("s", "5"),
                ("B", "8"),
                ("O", "0"), ("o", "0"),
                ("A", "4")
            };

            foreach (var pair in map)
            {
                text = text.Replace(pair.from, pair.to);
            }

            return text;
        }

        private (int current, int max, double ratio) ParseHPText(string text)
        {
            var match = Regex.Match(text, @"(\d+)\s*/\s*(\d+)");
            if (match.Success)
            {
                int curr = int.Parse(match.Groups[1].Value);
                int max = int.Parse(match.Groups[2].Value);
                if (max > 0)
                {
                    _lastMaxHp = max;
                    return Validate(curr, max);
                }
            }

            // Fallback for concatenated strings (e.g., 1007167)
            if (_lastMaxHp > 0)
            {
                string maxStr = _lastMaxHp.ToString();
                int idx = text.LastIndexOf(maxStr);
                if (idx > 0)
                {
                    string left = text.Substring(0, idx);
                    var numbers = Regex.Matches(left, @"\d+").Cast<Match>().Select(m => m.Value).ToList();
                    if (numbers.Any())
                    {
                        string leftStr = string.Join("", numbers);
                        if (int.TryParse(leftStr, out int curr))
                        {
                            return Validate(curr, _lastMaxHp);
                        }
                    }
                }
            }

            return (0, 0, 1.0);
        }

        private (int current, int max, double ratio) Validate(int curr, int max)
        {
            if (max > 999) return (0, 0, 1.0);
            if (curr > max) return (curr % 100, max, (curr % 100) / (double)max); // Naive correction
            return (curr, max, curr / (double)max);
        }

        private Bitmap CaptureRegion(int x, int y, int width, int height)
        {
            Bitmap bmp = new Bitmap(width, height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.CopyFromScreen(x, y, 0, 0, new System.Drawing.Size(width, height));
            }
            return bmp;
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    }
}
