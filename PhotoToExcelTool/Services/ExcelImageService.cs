// ==============================================================================
// 檔案名稱：Services/ExcelImageService.cs
// 學習重點：
// 1. 純記憶體串流處理 (MemoryStream)：全程不在硬碟產生任何暫存檔（零落地），透過 MemoryStream 讀寫 Excel 與圖片，速度快且無殘留檔案權限問題。
// 2. 針對 NAS / SMB 網路磁碟機的多執行緒併發 I/O 優化 (Parallel I/O & SequentialScan)：
//    - 掃描階段：不遞迴掃描整個目錄樹，僅針對 config.yml 規則指定的子資料夾（如 PRE/01~15、FQC/01~15）以多執行緒平行查詢目錄項目。
//    - 讀取階段：透過 Parallel.ForEach 搭配 128KB 緩衝區與 FileOptions.SequentialScan 旗標，同時以多執行緒從 NAS 將命中的原始圖片平行預讀至記憶體 byte[]（100% 保留原始解析度與畫質、零壓縮），徹底消除單執行緒跨網等待 (SMB Latency)。
// 3. 雜湊表 O(1) 快速查找 (Dictionary / ConcurrentDictionary)：將照片清單轉為雜湊表，查找時間複雜度為 O(1) 常數時間。
// 4. DPI 物理單位換算：標準螢幕與 Excel 圖形座標為 96 DPI（每英吋 96 像素，1 英吋 = 2.54 公分）：
//    - 1 cm = 96 / 2.54 ≈ 37.79527559 px
//    - 1 mm = 3.779527559 px
// 5. ClosedXML 圖片錨點定位 (Anchor & Placement)：
//    - 必須先呼叫 `.MoveTo(cell, xOffset, yOffset)`（將圖片錨定為 OneCellAnchor / Move 模式），
//      再呼叫 `.WithSize(widthPx, heightPx)` 設定精確像素寬高，才不會被壓縮進單一儲存格或拋出 Placement 例外。
// ==============================================================================

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using PhotoToExcelTool.Models;

namespace PhotoToExcelTool.Services;

/// <summary>
/// 單筆掃描到的照片檔案中繼資料（供前端快速建立比對預覽清單，無須先傳輸整張圖片本體）。
/// </summary>
public record ScannedPhotoInfo(
    string RelativePath,
    string FullPath,
    string FileName,
    long SizeBytes
);

/// <summary>
/// 負責極速掃描本機/NAS 照片目錄，並在記憶體中解析 Excel 底稿、平行預讀原始圖片且批次精確插入指定儲存格的核心服務。
/// </summary>
public class ExcelImageService
{
    // NAS / SMB 平行讀取執行緒上限：12~16 條併發連線可有效填滿區網頻寬並掩蓋網路延遲，同時避免過度競爭
    private const int ScanParallelism = 16;
    private const int ReadParallelism = 12;
    private const int SmbReadBufferSize = 131072; // 128 KB 緩衝區，觸發 SMB2 大區塊預讀

    /// <summary>
    /// 僅針對規則有使用到的目標相對子資料夾（如 PRE/01、FQC/01），以多執行緒平行掃描 NAS / 本機目錄下的相容圖片檔。
    /// 避免全資料夾遞迴掃描無關目錄，在網路磁碟機環境下可於毫秒級完成。
    /// </summary>
    public List<ScannedPhotoInfo> ScanRuleFolders(
        string rootFolderPath,
        IEnumerable<string> targetRelativeFolders,
        IEnumerable<string> acceptExtensions)
    {
        string cleanRoot = Path.GetFullPath(rootFolderPath.Trim().Trim('"'));
        if (!Directory.Exists(cleanRoot))
        {
            throw new DirectoryNotFoundException($"找不到指定的照片資料夾：{cleanRoot}");
        }

        // 建立合法副檔名集合（統一含 '.' 且忽略大小寫，例如 ".jpg", ".png"）
        var validExts = new HashSet<string>(
            acceptExtensions
                .Where(e => !string.IsNullOrWhiteSpace(e))
                .Select(e =>
                {
                    string trimmed = e.Trim();
                    return trimmed.StartsWith('.') ? trimmed : $".{trimmed}";
                }),
            StringComparer.OrdinalIgnoreCase);

        if (validExts.Count == 0)
        {
            validExts.Add(".jpg");
            validExts.Add(".png");
        }

        // 去重並正規化目標子目錄（例如 "PRE/01"）
        var distinctFolders = targetRelativeFolders
            .Where(f => !string.IsNullOrWhiteSpace(f))
            .Select(f => f.Trim().Replace('\\', '/').Trim('/'))
            .Where(f => f.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        var results = new ConcurrentBag<ScannedPhotoInfo>();

        Parallel.ForEach(
            distinctFolders,
            new ParallelOptions { MaxDegreeOfParallelism = ScanParallelism },
            relFolder =>
            {
                try
                {
                    string subDirFullPath = Path.Combine(cleanRoot, relFolder.Replace('/', Path.DirectorySeparatorChar));
                    var dirInfo = new DirectoryInfo(subDirFullPath);
                    if (!dirInfo.Exists) return;

                    // DirectoryInfo.EnumerateFiles 在 SMB2 下單次網路封包即可取回檔名與檔案大小，無須逐檔發送 stat 請求
                    foreach (var fileInfo in dirInfo.EnumerateFiles("*", SearchOption.TopDirectoryOnly))
                    {
                        if (validExts.Contains(fileInfo.Extension))
                        {
                            string normalizedRelPath = $"{relFolder}/{fileInfo.Name}";
                            results.Add(new ScannedPhotoInfo(
                                RelativePath: normalizedRelPath,
                                FullPath: fileInfo.FullName,
                                FileName: fileInfo.Name,
                                SizeBytes: fileInfo.Length));
                        }
                    }
                }
                catch
                {
                    // 單一子資料夾若無權限或不存在則安全略過，後續比對會自動列為缺漏
                }
            });

        return results
            .OrderBy(r => r.RelativePath, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    /// <summary>
    /// 【極速直連模式】直接從本機或 NAS 根目錄多執行緒平行預讀 100% 原始照片至記憶體，並批次插入 Excel 底稿。
    /// 省去瀏覽器 webkitdirectory 掃描與 HTTP FormData 上傳的雙重跨網開銷。
    /// </summary>
    public (byte[] FileBytes, List<string> Warnings) ProcessExcelFromDirectory(
        IFormFile templateFile,
        string rootFolderPath,
        AppConfig config)
    {
        // 1. 收集 config.yml 中所有規則定義的目標子目錄
        var ruleFolders = new List<string>();
        foreach (var rule in config.Rules)
        {
            if (!string.IsNullOrWhiteSpace(rule.IQC?.Folder)) ruleFolders.Add(rule.IQC!.Folder);
            if (!string.IsNullOrWhiteSpace(rule.FQC?.Folder)) ruleFolders.Add(rule.FQC!.Folder);
        }

        // 2. 多執行緒平行掃描目標資料夾，建立 O(1) 檔案路徑索引表
        var scannedFiles = ScanRuleFolders(rootFolderPath, ruleFolders, config.AcceptExtension);
        var pathMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in scannedFiles)
        {
            string key = item.RelativePath.Trim().Replace('\\', '/').ToLowerInvariant();
            pathMap.TryAdd(key, item.FullPath);
        }

        // 3. 執行共用核心引擎（支援多執行緒平行將命中的實體檔案讀入記憶體 byte[]）
        return ExecuteWorkbookInsertion(
            templateFile,
            config,
            tryExists: targetKey => pathMap.ContainsKey(targetKey),
            loadBytesParallel: matchedKeys =>
            {
                var byteCache = new ConcurrentDictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
                var errors = new ConcurrentDictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                Parallel.ForEach(
                    matchedKeys,
                    new ParallelOptions { MaxDegreeOfParallelism = ReadParallelism },
                    targetKey =>
                    {
                        if (!pathMap.TryGetValue(targetKey, out var fullPath)) return;
                        try
                        {
                            byteCache[targetKey] = ReadFileBytesSequential(fullPath);
                        }
                        catch (Exception ex)
                        {
                            errors[targetKey] = ex.Message;
                        }
                    });

                return (byteCache, errors);
            });
    }

    /// <summary>
    /// 【相容上傳模式】接收前端透過 multipart/form-data 上傳的照片集合，多執行緒平行讀入記憶體並插入 Excel 底稿。
    /// </summary>
    public (byte[] FileBytes, List<string> Warnings) ProcessExcel(
        IFormFile templateFile,
        IReadOnlyList<IFormFile> photos,
        IReadOnlyList<string> relativePaths,
        AppConfig config)
    {
        var photoMap = new Dictionary<string, IFormFile>(StringComparer.OrdinalIgnoreCase);
        int count = Math.Min(photos.Count, relativePaths.Count);
        for (int i = 0; i < count; i++)
        {
            var rawPath = relativePaths[i];
            if (string.IsNullOrWhiteSpace(rawPath)) continue;

            string normalizedPath = rawPath.Trim().Replace('\\', '/').ToLowerInvariant();
            photoMap[normalizedPath] = photos[i];
        }

        return ExecuteWorkbookInsertion(
            templateFile,
            config,
            tryExists: targetKey => photoMap.ContainsKey(targetKey),
            loadBytesParallel: matchedKeys =>
            {
                var byteCache = new ConcurrentDictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
                var errors = new ConcurrentDictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                Parallel.ForEach(
                    matchedKeys,
                    new ParallelOptions { MaxDegreeOfParallelism = ReadParallelism },
                    targetKey =>
                    {
                        if (!photoMap.TryGetValue(targetKey, out var formFile)) return;
                        try
                        {
                            using var src = formFile.OpenReadStream();
                            if (formFile.Length > 0 && formFile.Length <= int.MaxValue)
                            {
                                var buf = GC.AllocateUninitializedArray<byte>((int)formFile.Length);
                                src.ReadExactly(buf);
                                byteCache[targetKey] = buf;
                            }
                            else
                            {
                                using var ms = new MemoryStream();
                                src.CopyTo(ms);
                                byteCache[targetKey] = ms.ToArray();
                            }
                        }
                        catch (Exception ex)
                        {
                            errors[targetKey] = ex.Message;
                        }
                    });

                return (byteCache, errors);
            });
    }

    /// <summary>
    /// 核心兩階段處理管線：
    /// 第一階段：解析 Excel 底稿並比對所有有效工作表真正需要的照片清單（缺失的工作表直接跳過不讀圖）。
    /// 第二階段：以多執行緒平行將所有命中照片 100% 原封不動預讀進 RAM，再於記憶體內極速寫入 ClosedXML。
    /// </summary>
    private (byte[] FileBytes, List<string> Warnings) ExecuteWorkbookInsertion(
        IFormFile templateFile,
        AppConfig config,
        Func<string, bool> tryExists,
        Func<HashSet<string>, (ConcurrentDictionary<string, byte[]> ByteCache, ConcurrentDictionary<string, string> ReadErrors)> loadBytesParallel)
    {
        var warnings = new List<string>();

        // 載入 Excel 底稿至純記憶體串流
        using var templateStream = new MemoryStream();
        templateFile.CopyTo(templateStream);
        templateStream.Position = 0;

        using var workbook = new XLWorkbook(templateStream);

        // 收集待執行的插圖任務清單，先完成所有路徑匹配以便一次性平行預讀圖片
        var pendingPlacements = new List<PendingPlacement>();
        var uniqueTargetKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var rule in config.Rules)
        {
            if (!workbook.TryGetWorksheet(rule.Sheet, out var ws))
            {
                warnings.Add($"[{rule.Sheet}]：Sheet缺失");
                continue;
            }

            int widthPx = (int)Math.Round(rule.Size.Width * 37.79527559);
            int heightPx = (int)Math.Round(rule.Size.Height * 37.79527559);
            int leftOffsetPx = Math.Max(0, (int)Math.Round(rule.Size.LeftPadding * 3.779527559));
            int topOffsetPx = Math.Max(0, (int)Math.Round(rule.Size.TopPadding * 3.779527559));

            if (rule.IQC != null)
            {
                CollectStagePlacements(ws, rule.Sheet, "IQC", rule.IQC, config.AcceptExtension, tryExists, widthPx, heightPx, leftOffsetPx, topOffsetPx, pendingPlacements, uniqueTargetKeys, warnings);
            }

            if (rule.FQC != null)
            {
                CollectStagePlacements(ws, rule.Sheet, "FQC", rule.FQC, config.AcceptExtension, tryExists, widthPx, heightPx, leftOffsetPx, topOffsetPx, pendingPlacements, uniqueTargetKeys, warnings);
            }
        }

        // 多執行緒平行將所有需要的原始照片位元組一次拉進記憶體（100% 原始畫質零壓縮）
        var (byteCache, readErrors) = loadBytesParallel(uniqueTargetKeys);

        // 在純記憶體中將已預讀的 byte[] 插入對應儲存格（零磁碟/網路等待）
        foreach (var item in pendingPlacements)
        {
            if (readErrors.TryGetValue(item.TargetKey, out var readErr))
            {
                warnings.Add($"[{item.SheetName}] {item.StageName} - 編號 {item.PhotoKey} ({item.CellAddress})：讀取照片失敗 ({readErr})");
                continue;
            }

            if (!byteCache.TryGetValue(item.TargetKey, out var photoBytes))
            {
                warnings.Add($"[{item.SheetName}] {item.StageName} - 編號 {item.PhotoKey} ({item.CellAddress})：該格照片缺失");
                continue;
            }

            try
            {
                using var imgStream = new MemoryStream(photoBytes, writable: false);

                // ClosedXML 規定同一張工作表內的圖片名稱 (Picture Name) 不可重複，且長度上限為 31 字元，
                // 因此組合「階段_編號_儲存格」（例如 IQC_14_B28）作為簡潔的唯一識別名稱
                string keyTrimmed = item.PhotoKey.Trim();
                string picName = $"{item.StageName}_{keyTrimmed}_{item.CellAddress.Trim()}".Replace(" ", "_");
                if (picName.Length > 31)
                {
                    picName = picName[..31];
                }

                item.Worksheet.AddPicture(imgStream, picName)
                    .MoveTo(item.Worksheet.Cell(item.CellAddress.Trim()), item.LeftOffsetPx, item.TopOffsetPx)
                    .WithSize(item.WidthPx, item.HeightPx);
            }
            catch (Exception ex)
            {
                warnings.Add($"[{item.SheetName}] {item.StageName} - 編號 {item.PhotoKey} ({item.CellAddress})：插入照片失敗 ({ex.Message})");
            }
        }

        using var outputStream = new MemoryStream();
        workbook.SaveAs(outputStream);
        return (outputStream.ToArray(), warnings);
    }

    /// <summary>
    /// 解析單一檢驗階段的儲存格與照片對應關係（保留嚴格目錄階層與單數位補零匹配設計）。
    /// </summary>
    private static void CollectStagePlacements(
        IXLWorksheet ws,
        string sheetName,
        string stageName,
        StageConfig stage,
        List<string> acceptExtensions,
        Func<string, bool> tryExists,
        int widthPx,
        int heightPx,
        int leftOffsetPx,
        int topOffsetPx,
        List<PendingPlacement> pendingPlacements,
        HashSet<string> uniqueTargetKeys,
        List<string> warnings)
    {
        if (stage.Cells == null || string.IsNullOrWhiteSpace(stage.Folder)) return;

        string folderNorm = stage.Folder.Trim().Replace('\\', '/').Trim('/');

        foreach (var (photoKey, cellAddress) in stage.Cells)
        {
            string keyTrimmed = photoKey.Trim();

            // 建立候選檔名清單：優先比對原始編號（如 "1"），若為單一數字則同時支援補零格式（如 "01"）
            var keyCandidates = new List<string> { keyTrimmed };
            if (int.TryParse(keyTrimmed, out _) && keyTrimmed.Length == 1)
            {
                keyCandidates.Add(keyTrimmed.PadLeft(2, '0'));
            }

            string? matchedTargetKey = null;

            foreach (var candidateKey in keyCandidates)
            {
                if (matchedTargetKey != null) break;

                foreach (var ext in acceptExtensions)
                {
                    string targetExt = ext.StartsWith('.') ? ext : $".{ext}";
                    string targetKey = $"{folderNorm}/{candidateKey}{targetExt}".ToLowerInvariant();

                    if (tryExists(targetKey))
                    {
                        matchedTargetKey = targetKey;
                        break;
                    }
                }
            }

            if (matchedTargetKey != null)
            {
                uniqueTargetKeys.Add(matchedTargetKey);
                pendingPlacements.Add(new PendingPlacement(
                    Worksheet: ws,
                    SheetName: sheetName,
                    StageName: stageName,
                    PhotoKey: photoKey,
                    CellAddress: cellAddress,
                    TargetKey: matchedTargetKey,
                    WidthPx: widthPx,
                    HeightPx: heightPx,
                    LeftOffsetPx: leftOffsetPx,
                    TopOffsetPx: topOffsetPx));
            }
            else
            {
                warnings.Add($"[{sheetName}] {stageName} - 編號 {photoKey} ({cellAddress})：該格照片缺失");
            }
        }
    }

    /// <summary>
    /// 使用 128KB 緩衝區與 SequentialScan 旗標最佳化 SMB / NAS 大檔循序讀取，並直接配置精確大小的記憶體陣列。
    /// </summary>
    private static byte[] ReadFileBytesSequential(string fullPath)
    {
        using var fs = new FileStream(
            fullPath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite,
            bufferSize: SmbReadBufferSize,
            options: FileOptions.SequentialScan);

        long len = fs.Length;
        if (len > 0 && len <= int.MaxValue)
        {
            var buffer = GC.AllocateUninitializedArray<byte>((int)len);
            fs.ReadExactly(buffer);
            return buffer;
        }

        using var ms = new MemoryStream();
        fs.CopyTo(ms);
        return ms.ToArray();
    }

    private readonly record struct PendingPlacement(
        IXLWorksheet Worksheet,
        string SheetName,
        string StageName,
        string PhotoKey,
        string CellAddress,
        string TargetKey,
        int WidthPx,
        int HeightPx,
        int LeftOffsetPx,
        int TopOffsetPx
    );
}
