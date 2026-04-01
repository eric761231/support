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
        private int _lastMaxMp = 0;

        public Rectangle HPRegion { get; private set; } = new Rectangle(50, 20, 100, 20);
        public Rectangle MPRegion { get; private set; } = new Rectangle(50, 40, 100, 20);

        public string AutoCalibrate(string windowTitle, Action<int, string>? onProgress = null)
        {
            onProgress?.Invoke(5, "正在尋找遊戲視窗...");
            IntPtr hwnd = FindWindow(null, windowTitle);
            if (hwnd == IntPtr.Zero) return "找不到遊戲視窗";

            onProgress?.Invoke(10, "正在解析畫面尺寸...");
            // 使用 DWM 獲取實際無視窗邊框的座標
            Win32Input.RECT rect = Win32Input.GetActualWindowRect(hwnd);
            int width = rect.Right - rect.Left;
            int height = rect.Bottom - rect.Top;

            if (width <= 0 || height <= 0) return "視窗大小異常，請確認遊戲是否正常執行";

            onProgress?.Invoke(15, "正在擷取工作區域...");
            using (Bitmap screenshot = CaptureRegion(rect.Left, rect.Top, width, height))
            using (Mat mat = screenshot.ToMat())
            using (Mat hsv = new Mat())
            {
                Cv2.CvtColor(mat, hsv, ColorConversionCodes.BGR2HSV);

                // 1. 尋找 HP/MP 文字錨點
                string log = "日誌: ";
                onProgress?.Invoke(30, "正在搜尋 HP 文字錨點...");
                Rectangle? hpAnchor = FindTextAnchor(mat, "HP", false);
                
                onProgress?.Invoke(50, "正在搜尋 MP 文字錨點...");
                Rectangle? mpAnchor = FindTextAnchor(mat, "MP", true);

                log += hpAnchor.HasValue ? $"[找到 HP 錨點 {hpAnchor.Value.X},{hpAnchor.Value.Y}] " : "[未找到 HP 錨點] ";
                log += mpAnchor.HasValue ? $"[找到 MP 錨點 {mpAnchor.Value.X},{mpAnchor.Value.Y}] " : "[未找到 MP 錨點] ";

                // 2. 備用方案：如果找到 HP 但沒找到 MP，則在 HP 附近範圍再搜尋一次
                if (hpAnchor.HasValue && !mpAnchor.HasValue)
                {
                    mpAnchor = FindTextAnchorInBand(mat, "MP", hpAnchor.Value.Y, hpAnchor.Value.Height);
                    if (mpAnchor.HasValue) log += "[使用掃描帶找到 MP] ";
                }

                // 3. 顏色遮罩輔助與條狀物搜尋
                Mat maskRed1 = new Mat(); Mat maskRed2 = new Mat();
                Cv2.InRange(hsv, new OpenCvSharp.Scalar(0, 100, 50), new OpenCvSharp.Scalar(10, 255, 255), maskRed1);
                Cv2.InRange(hsv, new OpenCvSharp.Scalar(160, 100, 50), new OpenCvSharp.Scalar(180, 255, 255), maskRed2);
                Mat maskHP = maskRed1 | maskRed2;
                Mat maskMP = new Mat();
                Cv2.InRange(hsv, new OpenCvSharp.Scalar(100, 30, 50), new OpenCvSharp.Scalar(140, 255, 255), maskMP);

                onProgress?.Invoke(80, "正在精確定位血量條邊界...");
                // 4. 定位邏輯：優先以 OCR 錨點位移
                // 使用 140x20 的高度，確保數字不會被切到（特別是上下震動時）
                var hpFound = hpAnchor != null ? (Rectangle?)new Rectangle(hpAnchor.Value.X + 22, hpAnchor.Value.Y, 140, 20) : FindLargestHorizontalBar(maskHP);
                
                // --- V3.1 支援非垂直對齊佈局 ---
                Rectangle? mpFound = null;
                if (mpAnchor != null)
                {
                    mpFound = new Rectangle(mpAnchor.Value.X + 6, mpAnchor.Value.Y, 140, 20);
                }
                else if (hpFound.HasValue)
                {
                    // 若 OCR 沒找到 MP，則在 HP 平行的高度範圍內尋找最可能的藍色條
                    log += "[啟動平行搜尋] ";
                    mpFound = FindBarNear(maskMP, hpFound.Value.X, hpFound.Value.Y, 20);
                }
                else
                {
                    // 若連 HP 都沒找到，才進行全畫面搜尋（最後手段）
                    mpFound = FindLargestHorizontalBar(maskMP);
                }

                // 5. 資源驗證：不再強制 X 對齊，且支援橫向並排 (Same Y)
                if (hpFound.HasValue && mpFound.HasValue)
                {
                    // 只有在 X 與 Y 同時都非常接近時（代表找到同一個物件），才進行垂直偏移修正
                    // 若 X 不同（如並排佈局），則允許相同高度 (Same Y)
                    if (Math.Abs(hpFound.Value.X - mpFound.Value.X) < 10 && Math.Abs(hpFound.Value.Y - mpFound.Value.Y) < 5) {
                        log += "[重疊修正啟動] ";
                        mpFound = new Rectangle(mpFound.Value.X, hpFound.Value.Y + 18, mpFound.Value.Width, mpFound.Value.Height);
                    }
                }


                string resultTxt = log + "\n";
                if (hpFound.HasValue) { HPRegion = hpFound.Value; resultTxt += $"座標 HP(X,Y): {HPRegion.X},{HPRegion.Y} "; }
                if (mpFound.HasValue) { MPRegion = mpFound.Value; resultTxt += $"MP(X,Y): {MPRegion.X},{MPRegion.Y} "; }

                onProgress?.Invoke(100, "校準完成！");
                return string.IsNullOrEmpty(resultTxt) ? "校準失敗：找不到狀態條。請確認遊戲介面未被遮擋。" : "校準報表: " + resultTxt;
            }
        }

        private Rectangle? FindBarNear(Mat mask, int x, int y, int heightRange)
        {
            // V3.3: 寬域搜尋，搜尋整個視窗寬度，但限定 Y 軸在 HP 高度附近 (上下各偏移一些)
            int scanY = Math.Max(0, y - 10);
            int scanH = Math.Min(mask.Height - scanY, heightRange + 20);

            using (Mat roi = new Mat(mask, new OpenCvSharp.Rect(0, scanY, mask.Width, scanH)))
            {
                var bar = FindLargestHorizontalBar(roi);
                if (bar.HasValue)
                    return new Rectangle(bar.Value.X, bar.Value.Y + scanY, bar.Value.Width, bar.Value.Height);
            }
            return null;
        }

        private Rectangle? FindTextAnchorInBand(Mat src, string keyword, int y, int height)
        {
            int scanY = Math.Max(0, y - 30);
            int scanHeight = Math.Min(src.Height - scanY, height + 60);
            using (Mat band = new Mat(src, new OpenCvSharp.Rect(0, scanY, src.Width, scanHeight)))
            {
                // 對於掃描帶搜尋，如果是 MP 關鍵字，同樣套用 MP 專屬辨識優化
                var anchor = FindTextAnchor(band, keyword, keyword.Contains("MP", StringComparison.OrdinalIgnoreCase));
                if (anchor.HasValue) 
                    return new Rectangle(anchor.Value.X, anchor.Value.Y + scanY, anchor.Value.Width, anchor.Value.Height);
            }
            return null;
        }

        private Rectangle? FindTextAnchor(Mat src, string keyword, bool forMp)
        {
            if (_ocr == null) return null;
            
            using (Mat gray = new Mat())
            using (Mat resized = new Mat())
            {
                if (forMp)
                {
                    // 對於 MP 標籤：使用紅色版提取以提高藍色背景對比度
                    Cv2.ExtractChannel(src, gray, 2);
                }
                else
                {
                    Cv2.CvtColor(src, gray, ColorConversionCodes.BGR2GRAY);
                }

                // V3.0 對比度拉伸優化
                Cv2.Normalize(gray, gray, 0, 255, NormTypes.MinMax);

                // 錨點辨識率提升至 5 倍，更加靈敏
                Cv2.Resize(gray, resized, new OpenCvSharp.Size(gray.Width * 5, gray.Height * 5), 0, 0, InterpolationFlags.Cubic);
                
                using (Bitmap bmp = resized.ToBitmap())
                using (var img = PixConverter.ToPix(bmp))
                using (var page = _ocr.Process(img, PageSegMode.SparseText))
                {
                    var iter = page.GetIterator();
                    if (iter != null)
                    {
                        do {
                            string word = iter.GetText(PageIteratorLevel.Word);
                            if (!string.IsNullOrEmpty(word) && word.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                            {
                                if (iter.TryGetBoundingBox(PageIteratorLevel.Word, out var box))
                                    return new Rectangle(box.X1 / 5, box.Y1 / 5, (box.X2 - box.X1) / 5, (box.Y2 - box.Y1) / 5);
                            }
                        } while (iter.Next(PageIteratorLevel.Word));
                    }
                }
            }
            return null;
        }

        private Rectangle? FindLargestHorizontalBar(Mat mask)
        {
            Cv2.Dilate(mask, mask, null, iterations: 1);
            Cv2.Erode(mask, mask, null, iterations: 1);

            OpenCvSharp.Point[][] contours;
            OpenCvSharp.HierarchyIndex[] hierarchy;
            Cv2.FindContours(mask, out contours, out hierarchy, RetrievalModes.External, ContourApproximationModes.ApproxSimple);

            Rectangle? best = null;
            double maxArea = 0;

            foreach (var contour in contours)
            {
                var rect = Cv2.BoundingRect(contour);
                if (rect.Width > 50 && rect.Width > rect.Height * 3 && rect.Height > 5)
                {
                    double area = rect.Width * rect.Height;
                    if (area > maxArea)
                    {
                        maxArea = area;
                        best = new Rectangle(rect.X, rect.Y, rect.Width, rect.Height);
                    }
                }
            }
            return best;
        }

        public HPMonitor()
        {
            try {
                // V3.7: 強化部署檢查，確保 'eng.traineddata' 存在
                string tessPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");
                string engData = Path.Combine(tessPath, "eng.traineddata");
                
                if (!Directory.Exists(tessPath) || !File.Exists(engData)) {
                    throw new FileNotFoundException($"OCR 語言包缺失！請確保 {tessPath} 資料夾內含有 eng.traineddata 檔案。");
                }
                
                _ocr = new TesseractEngine(tessPath, "eng", EngineMode.Default);
                _ocr.SetVariable("tessedit_char_whitelist", "0123456789/|\\!iIl- ");
                _ocr.SetVariable("tessedit_pageseg_mode", "7"); // 單行文字模式
            } catch (Exception ex) {
                // 如果初始化失敗，將詳細原因回報
                throw new Exception($"OCR 引擎初始化失敗: {ex.Message}");
            }
        }

        public (int current, int max, double ratio, Bitmap? processed, string rawText) DetectHP(string windowTitle)
        {
            return DetectValue(windowTitle, HPRegion, ref _lastMaxHp, false);
        }

        public (int current, int max, double ratio, Bitmap? processed, string rawText) DetectMP(string windowTitle)
        {
            return DetectValue(windowTitle, MPRegion, ref _lastMaxMp, true);
        }

        private (int current, int max, double ratio, Bitmap? processed, string rawText) DetectValue(string windowTitle, Rectangle region, ref int lastMax, bool isMp)
        {
            IntPtr hwnd = FindWindow(null, windowTitle);
            if (hwnd == IntPtr.Zero) return (0, 0, 1.0, null, "");

            // 使用 GetActualWindowRect 避免隱形邊框偏差
            Win32Input.RECT rect = Win32Input.GetActualWindowRect(hwnd);
            int x = rect.Left + region.X;
            int y = rect.Top + region.Y;

            try
            {
                using (Bitmap screenshot = CaptureRegion(x, y, region.Width, region.Height))
                using (Mat mat = screenshot.ToMat())
                using (Mat processed = PreprocessImage(mat, isMp))
                {
                    Bitmap processedBitmap = processed.ToBitmap();
                    using (var img = PixConverter.ToPix(processedBitmap))
                    {
                        string text = "";
                        if (_ocr != null)
                        {
                            // V3.9.3: 改用 SingleLine (7) 模式，這對數字狀態條更精準
                            using (var page = _ocr.Process(img, PageSegMode.SingleLine))
                            {
                                text = page.GetText().Trim();
                                text = PostProcessOCRText(text);
                            }
                        }
                        var (cur, max, rat) = ParseStatusText(text, ref lastMax);

                        // V3.5: 儲存偵錯影像到目錄中，供使用者輔助檢查截圖品質
                        try {
                            processedBitmap.Save(isMp ? "mp_debug.bmp" : "hp_debug.bmp", System.Drawing.Imaging.ImageFormat.Bmp);
                        } catch { }

                        return (cur, max, rat, processedBitmap, text);
                    }
                }
            }
            catch
            {
                return (0, 0, 1.0, null, "");
            }
        }

        private Mat PreprocessImage(Mat src, bool isMp)
        {
            Mat work = new Mat();
            if (isMp)
            {
                // 針對藍色 MP 條：提取紅色版以控制背景噪聲並獲得最高文字對比度
                // 改用 ExtractChannel 避免 Cv2.Split 造成的生命週期管理錯誤與閃退
                Cv2.ExtractChannel(src, work, 2); 
            }
            else
            {
                Cv2.CvtColor(src, work, ColorConversionCodes.BGR2GRAY);
            }

            Mat resized = new Mat();
            // V3.8: 使用 4 倍縮放，這是 Tesseract 辨認的最佳尺寸
            Cv2.Resize(work, resized, new OpenCvSharp.Size(work.Width * 4, work.Height * 4), 0, 0, InterpolationFlags.Cubic);
            
            // V3.9.6: 修復筆劃削掉太多的問題
            // 移除 GaussianBlur，保持原始筆劃銳利度
            
            // 將邏輯改回 Binary：恢復白色數字為白色 (255)，背景變黑 (0)
            // 門檻再下修至 135，確保筆劃邊緣的灰色部分也能被捕捉到，使字體加粗
            Mat thresh = new Mat();
            Cv2.Threshold(resized, resized, 135, 255, ThresholdTypes.Binary);
            
            // V3.9.3: 增加 10 像素的黑色邊框 (Padding)
            // Tesseract 在文字緊貼邊緣時辨識率極差，增加邊框可顯著提升準確率
            Mat padded = new Mat();
            Cv2.CopyMakeBorder(resized, padded, 10, 10, 10, 10, BorderTypes.Constant, new OpenCvSharp.Scalar(0));
            
            // 資源釋放
            work.Dispose();
            resized.Dispose();

            return padded;
        }

        private string PostProcessOCRText(string text)
        {
            // V3.9.7: 強力移除頭尾非數字字元 (不論長度)
            // 這能徹底清除如 `i`, `il`, `|`, `!` 等各種邊界雜訊
            text = text.Trim();
            
            // 移除頭部的所有非數字非空格字元
            text = Regex.Replace(text, @"^[^0-9]+", "").Trim();
            // 移除尾部的所有非數字非空格字元
            text = Regex.Replace(text, @"[^0-9]+$", "").Trim();
            
            // V3.9.4: 強效映射，將誤認的英文字母與符號修正回數字
            var map = new (string from, string to)[]
            {
                ("G", "6"), ("b", "6"), ("T", "7"), ("t", "7"), ("S", "5"), ("s", "5"), ("B", "8"),
                ("O", "0"), ("o", "0"), ("A", "4"),
                ("i", "1"), ("l", "1"), ("I", "1"), ("!", "1"),
                ("|", "/"), ("\\", "/"), ("-", "/")
            };
            foreach (var pair in map) text = text.Replace(pair.from, pair.to);
            return text;
        }

        private (int current, int max, double ratio) ParseStatusText(string text, ref int lastMax)
        {
            if (!text.Any(char.IsDigit)) return (0, 0, 1.0);

            // V3.9.7: 優化 Regex
            var match = Regex.Match(text, @"(\d+)\s*[/\|\\!1\s-]+\s*(\d+)");
            if (match.Success)
            {
                int curr = int.Parse(match.Groups[1].Value);
                int max = int.Parse(match.Groups[2].Value);
                if (max > 0)
                {
                    lastMax = max;
                    return Validate(curr, max);
                }
            }

            var numbers = Regex.Matches(text, @"\d+").Cast<Match>().Select(m => int.Parse(m.Value)).ToList();
            if (numbers.Count >= 2)
            {
                int curr = numbers[0];
                int max = numbers[1];

                // 排除開頭是「1」但後接大數字的常見噪點（若我們已經有穩定的 lastMax）
                if (numbers.Count >= 3 && curr == 1 && lastMax > 0 && Math.Abs(max - lastMax) > 5)
                {
                    curr = numbers[1];
                    max = numbers[2];
                }

                if (curr > max) { int temp = curr; curr = max; max = temp; }
                if (max > 0) 
                {
                    lastMax = max;
                    return Validate(curr, max);
                }
            }

            if (lastMax > 0 && numbers.Count == 1)
            {
                return Validate(numbers[0], lastMax);
            }
            
            return (0, 0, 1.0);
        }

        private (int current, int max, double ratio) Validate(int curr, int max)
        {
            if (max > 9999) return (0, 0, 1.0);
            if (curr > max) return (max, max, 1.0);
            return (curr, max, (double)curr / max);
        }

        private Bitmap CaptureRegion(int x, int y, int width, int height)
        {
            x = Math.Max(0, x);
            y = Math.Max(0, y);
            Bitmap bmp = new Bitmap(width, height);
            using (Graphics g = Graphics.FromImage(bmp))
                g.CopyFromScreen(x, y, 0, 0, new System.Drawing.Size(width, height));
            return bmp;
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string? lpClassName, string lpWindowName);
    }
}
