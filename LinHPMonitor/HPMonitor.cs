using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using System.Diagnostics;
using System.Text.Json;
using Tesseract;

namespace LinHPMonitor
{
    public class HPMonitor
    {
        public bool UsePaddleOcr { get; set; } = true;
        private TesseractEngine? _ocr;
        private int _lastMaxHp = 0;
        private int _lastMaxMp = 0;

        // 顏色填充比快取：fillRatio 未達變動門檻時跳過 OCR，直接回傳快取
        // 5% 門檻：遊戲 UI 渲染噪點通常 < 1%，HP/MP 真實改變才會超過 5%
        private const double OcrChangeTrigger = 0.05;
        // 即使 fillRatio 有小變動，最短也要等 2 秒才重跑一次 OCR
        private const int MinOcrIntervalMs = 2000;
        private double _lastHpFillRatio = -1.0;
        private double _lastMpFillRatio = -1.0;
        private DateTime _lastHpOcrTime = DateTime.MinValue;
        private DateTime _lastMpOcrTime = DateTime.MinValue;
        // 快取上一次成功 OCR 的解析結果（不含 Bitmap，避免記憶體問題）
        private (int cur, int max, double ratio, string text) _cachedHpOcr = (0, 0, 1.0, "");
        private (int cur, int max, double ratio, string text) _cachedMpOcr = (0, 0, 1.0, "");

        // V4.0: 預設座標調整計畫 - 寬度 140, 高度增加為 22 以確保不切到文字
        // 使用者提供：血條與魔條為寬 260 高 40，座標為中心點
        // 轉換為左上角座標 (centerX - width/2, centerY - height/2)
        public Rectangle HPRegion { get; private set; } = new Rectangle(770 - 260/2, 775 - 40/2, 260, 40);
        public Rectangle MPRegion { get; private set; } = new Rectangle(1110 - 260/2, 780 - 40/2, 260, 40);

        public string AutoCalibrate(string windowTitle, Action<int, string>? onProgress = null)
        {
            onProgress?.Invoke(5, "正在尋找遊戲視窗...");
            IntPtr hwnd = Win32Input.FindWindowPartial(windowTitle);
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
                // 優先嘗試含冒號的標籤（例如 "HP:"），若找不到再降級到不含冒號的 "HP"
                Rectangle? hpAnchor = FindTextAnchor(mat, "HP:", false) ?? FindTextAnchor(mat, "HP", false);

                onProgress?.Invoke(50, "正在搜尋 MP 文字錨點...");
                Rectangle? mpAnchor = FindTextAnchor(mat, "MP:", true) ?? FindTextAnchor(mat, "MP", true);

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
                // 4. 定位邏輯：根據視窗尺寸自動決定擷取區域大小與位置
                // 以視窗寬高的比例計算建議的區域尺寸，並以 OCR 錨點中心為基準
                int regionWidth = Math.Max(140, (int)(width * 0.135));
                int regionHeight = Math.Max(22, (int)(height * 0.037));

                Rectangle? hpFound = null;
                if (hpAnchor != null)
                {
                    int cx = hpAnchor.Value.X + hpAnchor.Value.Width / 2;
                    int cy = hpAnchor.Value.Y + hpAnchor.Value.Height / 2;
                    int x = cx - regionWidth / 2 + (int)(regionWidth * 0.08); // 小幅右移以避開標籤
                    int y = cy - regionHeight / 2 + (int)(regionHeight * 0.0) + 10; // 向下偏移 10 像素
                    x = Math.Max(0, Math.Min(x, width - regionWidth));
                    y = Math.Max(0, Math.Min(y, height - regionHeight));
                    hpFound = new Rectangle(x, y, regionWidth, regionHeight);
                }
                else
                {
                    var bar = FindLargestHorizontalBar(maskHP);
                    if (bar.HasValue)
                    {
                        int cx = bar.Value.X + bar.Value.Width / 2;
                        int cy = bar.Value.Y + bar.Value.Height / 2;
                        int x = cx - regionWidth / 2;
                        int y = cy - regionHeight / 2 + 10;
                        x = Math.Max(0, Math.Min(x, width - regionWidth));
                        y = Math.Max(0, Math.Min(y, height - regionHeight));
                        hpFound = new Rectangle(x, y, regionWidth, regionHeight);
                    }
                }

                // MP 同理：優先使用 MP 錨點，否則在 HP 附近或全畫面搜尋
                Rectangle? mpFound = null;
                if (mpAnchor != null)
                {
                    int cx = mpAnchor.Value.X + mpAnchor.Value.Width / 2;
                    int cy = mpAnchor.Value.Y + mpAnchor.Value.Height / 2;
                    int x = cx - regionWidth / 2 + (int)(regionWidth * 0.08);
                    int y = cy - regionHeight / 2 + (int)(regionHeight * 0.0) + 10;
                    x = Math.Max(0, Math.Min(x, width - regionWidth));
                    y = Math.Max(0, Math.Min(y, height - regionHeight));
                    mpFound = new Rectangle(x, y, regionWidth, regionHeight);
                }
                else if (hpFound.HasValue)
                {
                    log += "[啟動平行搜尋] ";
                    mpFound = FindBarNear(maskMP, hpFound.Value.X, hpFound.Value.Y, 60);
                    if (mpFound.HasValue)
                    {
                        // 若找到的 bar 大小和預期差很多，重新套用建議尺寸並置中
                        int cx = mpFound.Value.X + mpFound.Value.Width / 2;
                        int cy = mpFound.Value.Y + mpFound.Value.Height / 2;
                        int x = Math.Max(0, Math.Min(cx - regionWidth / 2, width - regionWidth));
                        int y = Math.Max(0, Math.Min(cy - regionHeight / 2 + 10, height - regionHeight));
                        mpFound = new Rectangle(x, y, regionWidth, regionHeight);
                    }
                }
                else
                {
                    var bar = FindLargestHorizontalBar(maskMP);
                    if (bar.HasValue)
                    {
                        int cx = bar.Value.X + bar.Value.Width / 2;
                        int cy = bar.Value.Y + bar.Value.Height / 2;
                        int x = Math.Max(0, Math.Min(cx - regionWidth / 2, width - regionWidth));
                        int y = Math.Max(0, Math.Min(cy - regionHeight / 2, height - regionHeight));
                        mpFound = new Rectangle(x, y, regionWidth, regionHeight);
                    }
                }

                // 5. 資源驗證：不再強制 X 對齊，且支援橫向並排 (Same Y)
                if (hpFound.HasValue && mpFound.HasValue)
                {
                    // 只有在 X 與 Y 同時都非常接近時（代表找到同一個物件），才進行垂直偏移修正
                    // 若 X 不同（如並排佈局），則允許相同高度 (Same Y)
                    if (Math.Abs(hpFound.Value.X - mpFound.Value.X) < 10 && Math.Abs(hpFound.Value.Y - mpFound.Value.Y) < 5) {
                        log += "[重疊修正啟動] ";
                        // V4.0: 位置往下位移 20 像素以避開 HP 條
                        mpFound = new Rectangle(mpFound.Value.X, hpFound.Value.Y + 20, mpFound.Value.Width, mpFound.Value.Height);
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
            int scanHeight = Math.Min(src.Height - scanY, Math.Max(height + 60, 100));
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

        // 平滑與去抖：保留最近 N 次的解析結果，減少錯誤觸發
        private readonly int _smoothWindow = 5;
        private readonly Queue<(int cur, int max)> _hpHistory = new Queue<(int cur, int max)>();
        private readonly Queue<(int cur, int max)> _mpHistory = new Queue<(int cur, int max)>();

        // 允許在運行時設定/覆寫 HP/MP 的擷取區域
        public void SetRegions(Rectangle hp, Rectangle mp)
        {
            HPRegion = hp;
            MPRegion = mp;
        }

        // 針對擷取時可微調的 Y 偏移（像素）以修正預覽位置
        public int HPCaptureYOffset { get; set; } = 35;
        public int MPCaptureYOffset { get; set; } = 35;

        // ─────────────────────────────────────────────────────────────
        // 欄位投影法 fillRatio（方向感應版）
        //   HP 紅色由左至右遞減：找最右端有紅色的欄位 → rightEdge / totalCols
        //   MP 藍色由右至左遞減：找最左端有藍色的欄位 → (totalCols - leftEdge) / totalCols
        // ─────────────────────────────────────────────────────────────
        private double ComputeFillRatio(Mat mat, bool isMp)
        {
            try
            {
                using var hsv = new Mat();
                Cv2.CvtColor(mat, hsv, ColorConversionCodes.BGR2HSV);

                Mat mask;
                if (isMp)
                {
                    mask = new Mat();
                    Cv2.InRange(hsv, new OpenCvSharp.Scalar(100, 30, 50), new OpenCvSharp.Scalar(140, 255, 255), mask);
                }
                else
                {
                    Mat m1 = new Mat(), m2 = new Mat();
                    Cv2.InRange(hsv, new OpenCvSharp.Scalar(0, 100, 50), new OpenCvSharp.Scalar(10, 255, 255), m1);
                    Cv2.InRange(hsv, new OpenCvSharp.Scalar(160, 100, 50), new OpenCvSharp.Scalar(180, 255, 255), m2);
                    mask = m1 | m2;
                    m1.Dispose(); m2.Dispose();
                }

                // 沿 row 方向加總，得到每欄的顏色強度（32 位元整數）
                using var proj = new Mat();
                Cv2.Reduce(mask, proj, 0, ReduceTypes.Sum, MatType.CV_32S);
                mask.Dispose();

                // 門檻：欄內至少 25% 的行有顏色才算有效填充欄
                int threshold = (int)(mat.Height * 0.25 * 255);
                int totalCols = proj.Cols;

                double ratio;
                if (!isMp)
                {
                    // HP：由左至右遞減，找最右端有色欄位
                    int lastColoredCol = -1;
                    for (int col = 0; col < totalCols; col++)
                        if (proj.At<int>(0, col) >= threshold)
                            lastColoredCol = col;
                    ratio = lastColoredCol < 0 ? 0.0 : (double)(lastColoredCol + 1) / totalCols;
                }
                else
                {
                    // MP：由右至左遞減，找最左端有色欄位
                    int firstColoredCol = -1;
                    for (int col = 0; col < totalCols; col++)
                    {
                        if (proj.At<int>(0, col) >= threshold)
                        {
                            firstColoredCol = col;
                            break;
                        }
                    }
                    ratio = firstColoredCol < 0 ? 0.0 : (double)(totalCols - firstColoredCol) / totalCols;
                }

                return Math.Clamp(ratio, 0.0, 1.0);
            }
            catch { return 0.0; }
        }

        // ─────────────────────────────────────────────────────────────
        // 輕量偵測：只算 fillRatio + 預覽圖，完全不跑 OCR
        //   適用於已設定 UserMaxHp/Mp > 0 的情況
        // ─────────────────────────────────────────────────────────────
        public (double fillRatio, Bitmap? preview) GetQuickRatio(string windowTitle, bool isMp)
        {
            IntPtr hwnd = Win32Input.FindWindowPartial(windowTitle);
            if (hwnd == IntPtr.Zero) return (0.0, null);

            Win32Input.RECT rect = Win32Input.GetActualWindowRect(hwnd);
            Rectangle region = isMp ? MPRegion : HPRegion;
            int x = rect.Left + region.X;
            int y = rect.Top + region.Y + (isMp ? MPCaptureYOffset : HPCaptureYOffset);

            try
            {
                using var screenshot = CaptureRegion(x, y, region.Width, region.Height);
                using var mat = screenshot.ToMat();

                // 彩色預覽（4x 放大）
                using var colorPreview = new Mat();
                Cv2.Resize(mat, colorPreview, new OpenCvSharp.Size(mat.Width * 4, mat.Height * 4),
                           0, 0, InterpolationFlags.Cubic);
                var previewBitmap = colorPreview.ToBitmap();

                double fillRatio = ComputeFillRatio(mat, isMp);
                return (fillRatio, previewBitmap);
            }
            catch { return (0.0, null); }
        }

        // ─────────────────────────────────────────────────────────────
        // 精確偵測：用邊緣偵測法找血/魔條填充邊界，再映射到 maxValue
        //   HP（左→右遞減）：找最右端紅色欄 → (rightEdge+1)/totalCols × maxValue
        //   MP（右→左遞減）：找最左端藍色欄 → (totalCols-leftEdge)/totalCols × maxValue
        //
        //   改用邊緣而非分段計數，是因為條上疊加的「174/174」白字會遮擋部分紅/藍像素，
        //   讓中間的分段計數偏低；而邊緣位置不受文字遮擋影響，遠端邊界仍在正確位置。
        // ─────────────────────────────────────────────────────────────
        public (int current, double fillRatio, Bitmap? preview) GetCurrentBySegments(
            string windowTitle, bool isMp, int maxValue)
        {
            if (maxValue <= 0) return (0, 0.0, null);

            IntPtr hwnd = Win32Input.FindWindowPartial(windowTitle);
            if (hwnd == IntPtr.Zero) return (0, 0.0, null);

            Win32Input.RECT rect = Win32Input.GetActualWindowRect(hwnd);
            Rectangle region = isMp ? MPRegion : HPRegion;
            int x = rect.Left + region.X;
            int y = rect.Top + region.Y + (isMp ? MPCaptureYOffset : HPCaptureYOffset);

            try
            {
                using var screenshot = CaptureRegion(x, y, region.Width, region.Height);
                using var mat = screenshot.ToMat();

                // 彩色預覽（4x 放大）
                using var colorPreview = new Mat();
                Cv2.Resize(mat, colorPreview, new OpenCvSharp.Size(mat.Width * 4, mat.Height * 4),
                           0, 0, InterpolationFlags.Cubic);
                var previewBitmap = colorPreview.ToBitmap();

                // 用邊緣偵測法計算填充比，映射到 maxValue 格
                double ratio  = ComputeFillRatio(mat, isMp);
                int current   = (int)Math.Round(ratio * maxValue);
                current       = Math.Clamp(current, 0, maxValue);
                double fillRatio = (double)current / maxValue;

                return (current, fillRatio, previewBitmap);
            }
            catch { return (0, 0.0, null); }
        }

        // 現在回傳包含填充比例 fillRatio，供 UI 依顏色區塊計算當前數值
        public (int current, int max, double ratio, double fillRatio, Bitmap? processed, string rawText) DetectHP(string windowTitle)
        {
            return DetectValue(windowTitle, HPRegion, ref _lastMaxHp, false);
        }

        public (int current, int max, double ratio, double fillRatio, Bitmap? processed, string rawText) DetectMP(string windowTitle)
        {
            return DetectValue(windowTitle, MPRegion, ref _lastMaxMp, true);
        }

        private (int current, int max, double ratio, double fillRatio, Bitmap? processed, string rawText) DetectValue(string windowTitle, Rectangle region, ref int lastMax, bool isMp)
        {
            IntPtr hwnd = Win32Input.FindWindowPartial(windowTitle);
            if (hwnd == IntPtr.Zero) return (0, 0, 1.0, 0.0, null, "");

            Win32Input.RECT rect = Win32Input.GetActualWindowRect(hwnd);
            int x = rect.Left + region.X;
            int y = rect.Top + region.Y + (isMp ? MPCaptureYOffset : HPCaptureYOffset);

            try
            {
                using (Bitmap screenshot = CaptureRegion(x, y, region.Width, region.Height))
                using (Mat mat = screenshot.ToMat())
                {
                    // ✅ 預覽用：保留彩色原圖，只做縮放
                    Mat colorPreview = new Mat();
                    Cv2.Resize(mat, colorPreview, new OpenCvSharp.Size(mat.Width * 4, mat.Height * 4), 0, 0, InterpolationFlags.Cubic);
                    Bitmap previewBitmap = colorPreview.ToBitmap();
                    colorPreview.Dispose();

                    // ── 第一步：計算顏色填充比（快速，採欄位投影法）──
                    double fillRatio = ComputeFillRatio(mat, isMp);

                    // ── 第二步：雙重門檻決定是否重跑 OCR ──
                    // 條件 A：fillRatio 變化超過 5%（真實 HP/MP 改變）
                    // 條件 B：距上次 OCR 已超過 2 秒（防止噪點觸發）
                    double lastFill = isMp ? _lastMpFillRatio : _lastHpFillRatio;
                    DateTime lastOcrTime = isMp ? _lastMpOcrTime : _lastHpOcrTime;
                    var cachedOcr = isMp ? _cachedMpOcr : _cachedHpOcr;

                    bool fillChanged = lastFill < 0 || Math.Abs(fillRatio - lastFill) >= OcrChangeTrigger;
                    bool tooSoon    = (DateTime.Now - lastOcrTime).TotalMilliseconds < MinOcrIntervalMs;

                    // 快取命中：fillRatio 無顯著變化，或距上次 OCR 不到 2 秒
                    if (!fillChanged || tooSoon)
                    {
                        // rawText 回傳空字串，表示本次無 OCR，不觸發 log
                        return (cachedOcr.cur, cachedOcr.max, cachedOcr.ratio, fillRatio, previewBitmap, "");
                    }

                    // 更新 fillRatio 快取與 OCR 時間戳記
                    if (isMp) { _lastMpFillRatio = fillRatio; _lastMpOcrTime = DateTime.Now; }
                    else      { _lastHpFillRatio = fillRatio; _lastHpOcrTime = DateTime.Now; }

                    // ── 第三步：fillRatio 有顯著變化 → 執行完整 OCR 流程 ──
                    using (var processed = PreprocessImage(mat, isMp))
                    {
                        string text = "";
                        if (_ocr != null)
                        {
                            text = TryOcrWithVariants(processed, _ocr);
                            text = PostProcessOCRText(text);
                        }
                        var (cur, max, rat) = ParseStatusText(text, ref lastMax);

                        // 偵錯圖：彩色預覽 + 二值化結果
                        try { previewBitmap.Save(isMp ? "mp_debug.bmp" : "hp_debug.bmp", System.Drawing.Imaging.ImageFormat.Bmp); } catch { }
                        try {
                            using (var bin = processed.Clone())
                            using (var binBmp = bin.ToBitmap())
                                binBmp.Save(isMp ? "mp_bin_debug.bmp" : "hp_bin_debug.bmp", System.Drawing.Imaging.ImageFormat.Bmp);
                        } catch { }

                        // 平滑緩衝（多數決）
                        var hist = isMp ? _mpHistory : _hpHistory;
                        hist.Enqueue((cur, max));
                        while (hist.Count > _smoothWindow) hist.Dequeue();
                        (int sCur, int sMax) smoothed = (cur, max);
                        try
                        {
                            smoothed = hist.GroupBy(h => h).OrderByDescending(g => g.Count()).First().Key;
                        }
                        catch { }
                        double smoothRatio = smoothed.sMax > 0 ? (double)smoothed.sCur / smoothed.sMax : 1.0;

                        // PaddleOCR 備援（結果可疑時）
                        bool suspicious = smoothed.sMax == 0 || smoothed.sCur > smoothed.sMax || smoothRatio <= 0.01
                            || text.Count(c => char.IsDigit(c)) < 2
                            || Regex.Matches(text, "[^0-9/]").Count > text.Length / 2;
                        if (UsePaddleOcr && suspicious)
                        {
                            try
                            {
                                string imgPath = isMp ? "mp_debug.bmp" : "hp_debug.bmp";
                                string paddleText = RunPaddleOcr(imgPath);
                                if (!string.IsNullOrEmpty(paddleText))
                                {
                                    var (pcur, pmax, prat) = ParseStatusText(paddleText, ref lastMax);
                                    if (pmax > 0)
                                    {
                                        if (isMp) _cachedMpOcr = (pcur, pmax, prat, paddleText);
                                        else      _cachedHpOcr = (pcur, pmax, prat, paddleText);
                                        return (pcur, pmax, prat, fillRatio, previewBitmap, paddleText);
                                    }
                                }
                            }
                            catch { }
                        }

                        // 更新 OCR 快取
                        if (isMp) _cachedMpOcr = (smoothed.sCur, smoothed.sMax, smoothRatio, text);
                        else      _cachedHpOcr = (smoothed.sCur, smoothed.sMax, smoothRatio, text);
                        return (smoothed.sCur, smoothed.sMax, smoothRatio, fillRatio, previewBitmap, text);
                    }
                }
            }
            catch
            {
                return (0, 0, 1.0, 0.0, null, "");
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
            Cv2.Threshold(resized, resized, 135, 255, ThresholdTypes.Binary);

            // 使用形態學開運算移除孤立噪點，接著膨脹以補回細小筆劃
            Mat kernel = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(3, 3));
            Mat opened = new Mat();
            Cv2.MorphologyEx(resized, opened, MorphTypes.Open, kernel);
            Cv2.Dilate(opened, opened, kernel);

            // V3.9.3: 增加 10 像素的黑色邊框 (Padding)
            // Tesseract 在文字緊貼邊緣時辨識率極差，增加邊框可顯著提升準確率
            Mat padded = new Mat();
            Cv2.CopyMakeBorder(opened, padded, 10, 10, 10, 10, BorderTypes.Constant, new OpenCvSharp.Scalar(0));

            // 釋放中間資源
            kernel.Dispose();
            opened.Dispose();
            
            // 資源釋放
            work.Dispose();
            resized.Dispose();

            return padded;
        }

        private string PostProcessOCRText(string text)
        {
            // V3.9.7: 先做字元映射，再清理頭尾雜訊，讓分隔符不會被誤刪
            text = text.Trim();

            // V3.9.4: 強效映射，將誤認的英文字母與符號修正回數字或分隔符
            var map = new (string from, string to)[]
            {
                ("G", "6"), ("b", "6"), ("T", "7"), ("t", "7"), ("S", "5"), ("s", "5"), ("B", "8"),
                ("O", "0"), ("o", "0"), ("A", "4"),
                ("i", "1"), ("l", "1"), ("I", "1"), ("!", "1"),
                ("|", "/"), ("\\", "/"), ("-", "/")
            };
            foreach (var pair in map) text = text.Replace(pair.from, pair.to);

            // 映射完成後，再移除頭尾雜訊。尾部允許保留 '/' 做為分隔符
            text = Regex.Replace(text, @"^[^0-9]+", "").Trim();
            text = Regex.Replace(text, @"[^0-9/]+$", "").Trim();
            return text;
        }

        private (int current, int max, double ratio) ParseStatusText(string text, ref int lastMax)
        {
            if (!text.Any(char.IsDigit)) return (0, 0, 1.0);

            // V3.9.7: 優化 Regex
            // 優先匹配含分隔符的 fraction（如 179/174）
            var fracMatch = Regex.Match(text, @"(\d+)\s*[/]\s*(\d+)");
            if (fracMatch.Success)
            {
                int curr = int.Parse(fracMatch.Groups[1].Value);
                int max = int.Parse(fracMatch.Groups[2].Value);
                if (max > 0)
                {
                    lastMax = max;
                    return Validate(curr, max);
                }
            }

            // 找出所有數字片段，然後嘗試選出最合理的一對 (curr,max)
            var numbers = Regex.Matches(text, @"\d+").Cast<Match>().Select(m => int.Parse(m.Value)).ToList();
            if (numbers.Count >= 2)
            {
                // 列出所有可能的 pair (i<j)
                List<(int curr, int max)> candidates = new List<(int, int)>();
                for (int i = 0; i < numbers.Count - 1; i++)
                {
                    for (int j = i + 1; j < numbers.Count; j++)
                    {
                        int a = numbers[i];
                        int b = numbers[j];
                        // 只考慮 max > 0
                        if (b > 0) candidates.Add((a, b));
                        // 也考慮反向（若誤斷詞順序）
                        if (a > 0) candidates.Add((b, a));
                    }
                }

                // 選擇最合理的 candidate：優先 curr <= max，且與 lastMax 差距最小
                (int currSel, int maxSel) = (0, 0);
                int bestScore = int.MaxValue;
                foreach (var c in candidates)
                {
                    int curr = c.curr, max = c.max;
                    if (max <= 0) continue;
                    int score = 0;
                    if (curr > max) score += 100000; // 不合理的懲罰
                    // 與 lastMax 差距越小分數越低
                    score += lastMax > 0 ? Math.Abs(max - lastMax) : 0;
                    // 偏好較大的 max（避免選到單個 1 當 max）
                    score -= Math.Min(max, 500);

                    if (score < bestScore)
                    {
                        bestScore = score;
                        currSel = curr; maxSel = max;
                    }
                }

                if (maxSel > 0)
                {
                    lastMax = maxSel;
                    return Validate(currSel, maxSel);
                }
            }

            // 若只有一個數字且有 lastMax，則視為 current
            if (lastMax > 0 && numbers.Count == 1)
            {
                return Validate(numbers[0], lastMax);
            }

            return (0, 0, 1.0);
        }

        // 嘗試多種影像變體以提升 OCR 命中率：原圖、反相、膨脹、閉運算
        private string TryOcrWithVariants(Mat bin, TesseractEngine ocr)
        {
            string best = "";
            int bestDigits = 0;

            Mat[] variants = null;
            try
            {
                var kern = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(3, 3));
                var inv = new Mat(); Cv2.BitwiseNot(bin, inv);
                var dil = new Mat(); Cv2.Dilate(bin, dil, kern);
                var closed = new Mat(); Cv2.MorphologyEx(bin, closed, MorphTypes.Close, kern);
                variants = new Mat[] { bin.Clone(), inv, dil, closed };

                foreach (var v in variants)
                {
                    try
                    {
                        using (var bmp = v.ToBitmap())
                        using (var pix = PixConverter.ToPix(bmp))
                        using (var page = ocr.Process(pix, PageSegMode.SingleLine))
                        {
                            string t = page.GetText() ?? "";
                            t = t.Trim();
                            // 計分：先看數字數量，若相同則偏好含 '/' 的
                            int digits = t.Count(c => char.IsDigit(c));
                            bool hasFrac = t.Contains("/");
                            if (digits > bestDigits || (digits == bestDigits && hasFrac && !best.Contains("/")))
                            {
                                best = t;
                                bestDigits = digits;
                            }
                        }
                    }
                    catch { }
                }
            }
            finally
            {
                if (variants != null) foreach (var m in variants) m?.Dispose();
            }

            return best;
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

        // 呼叫外部 Python PaddleOCR 腳本，回傳解析到的文字（若失敗回傳空字串）
        private string RunPaddleOcr(string imgPath)
        {
            try
            {
                string python = "python";
                string script = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "paddle_ocr.py");
                var psi = new ProcessStartInfo
                {
                    FileName = python,
                    Arguments = $"\"{script}\" \"{imgPath}\"",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WorkingDirectory = AppDomain.CurrentDomain.BaseDirectory
                };

                using (var proc = Process.Start(psi))
                {
                    if (proc == null) return "";
                    string stdout = proc.StandardOutput.ReadToEnd();
                    string stderr = proc.StandardError.ReadToEnd();
                    proc.WaitForExit(5000);

                    if (string.IsNullOrWhiteSpace(stdout)) return "";

                    try
                    {
                        using (var doc = JsonDocument.Parse(stdout))
                        {
                            if (doc.RootElement.TryGetProperty("text", out var t))
                                return t.GetString() ?? "";
                        }
                    }
                    catch
                    {
                        // 如果不是 JSON，就回傳原始 stdout 的一行內容
                        return stdout.Trim();
                    }
                }
            }
            catch
            {
                // 靜默失敗，回傳空字串讓呼叫端決定後續
            }
            return "";
        }

    }
}
