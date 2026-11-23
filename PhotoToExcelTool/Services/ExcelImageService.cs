using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using PhotoToExcelTool.Models;

namespace PhotoToExcelTool.Services;

public class ExcelImageService
{
    public (byte[] FileBytes, List<string> Warnings) ProcessExcel(
        IFormFile templateFile,
        IReadOnlyList<IFormFile> photos,
        IReadOnlyList<string> relativePaths,
        AppConfig config)
    {
        var warnings = new List<string>();

        // 1. 建立 O(1) 檔案對照表 (Key 統一轉為小寫且斜線為 /)
        var photoMap = new Dictionary<string, IFormFile>(StringComparer.OrdinalIgnoreCase);
        int count = Math.Min(photos.Count, relativePaths.Count);
        for (int i = 0; i < count; i++)
        {
            var rawPath = relativePaths[i];
            if (string.IsNullOrWhiteSpace(rawPath)) continue;

            string normalizedPath = rawPath.Trim().Replace('\\', '/').ToLowerInvariant();
            photoMap[normalizedPath] = photos[i];
        }

        // 2. ClosedXML 純記憶體寫入邏輯
        using var templateStream = new MemoryStream();
        templateFile.CopyTo(templateStream);
        templateStream.Position = 0;

        using var workbook = new XLWorkbook(templateStream);

        // 3. 遍歷 config.yml 的每一條 rule
        foreach (var rule in config.Rules)
        {
            // Sheet 檢核
            if (!workbook.TryGetWorksheet(rule.Sheet, out var ws))
            {
                warnings.Add($"[{rule.Sheet}]：Sheet缺失");
                continue;
            }

            // 尺寸與留白換算 (96 DPI 標準：1 cm = 37.79527559 px, 1 mm = 3.779527559 px)
            int widthPx = (int)Math.Round(rule.Size.Width * 37.79527559);
            int heightPx = (int)Math.Round(rule.Size.Height * 37.79527559);
            int leftOffsetPx = Math.Max(0, (int)Math.Round(rule.Size.LeftPadding * 3.779527559));
            int topOffsetPx = Math.Max(0, (int)Math.Round(rule.Size.TopPadding * 3.779527559));

            // 處理 IQC
            if (rule.IQC != null)
            {
                ProcessStage(ws, rule.Sheet, "IQC", rule.IQC, config.AcceptExtension, photoMap, widthPx, heightPx, leftOffsetPx, topOffsetPx, warnings);
            }

            // 處理 FQC
            if (rule.FQC != null)
            {
                ProcessStage(ws, rule.Sheet, "FQC", rule.FQC, config.AcceptExtension, photoMap, widthPx, heightPx, leftOffsetPx, topOffsetPx, warnings);
            }
        }

        // 4. 將修改完成的 XLWorkbook 存入輸出 MemoryStream
        using var outputStream = new MemoryStream();
        workbook.SaveAs(outputStream);
        return (outputStream.ToArray(), warnings);
    }

    private void ProcessStage(
        IXLWorksheet ws,
        string sheetName,
        string stageName,
        StageConfig stage,
        List<string> acceptExtensions,
        Dictionary<string, IFormFile> photoMap,
        int widthPx,
        int heightPx,
        int leftOffsetPx,
        int topOffsetPx,
        List<string> warnings)
    {
        if (stage.Cells == null || string.IsNullOrWhiteSpace(stage.Folder)) return;

        string folderNorm = stage.Folder.Trim().Replace('\\', '/').Trim('/');

        foreach (var (photoKey, cellAddress) in stage.Cells)
        {
            bool matched = false;
            string keyTrimmed = photoKey.Trim();
            var keyCandidates = new List<string> { keyTrimmed };
            if (int.TryParse(keyTrimmed, out _) && keyTrimmed.Length == 1)
            {
                keyCandidates.Add(keyTrimmed.PadLeft(2, '0'));
            }

            foreach (var candidateKey in keyCandidates)
            {
                if (matched) break;

                foreach (var ext in acceptExtensions)
                {
                    string targetExt = ext.StartsWith('.') ? ext : $".{ext}";
                    string targetKey = $"{folderNorm}/{candidateKey}{targetExt}".ToLowerInvariant();

                    if (photoMap.TryGetValue(targetKey, out var photoFile))
                    {
                        try
                        {
                            using var imgStream = photoFile.OpenReadStream();
                            string picName = $"P_{sheetName}_{stageName}_{photoKey}_{cellAddress.Trim()}".Replace(" ", "_");
                            ws.AddPicture(imgStream, picName)
                              .MoveTo(ws.Cell(cellAddress.Trim()), leftOffsetPx, topOffsetPx)
                              .WithSize(widthPx, heightPx);
                            matched = true;
                            break;
                        }
                        catch (Exception ex)
                        {
                            warnings.Add($"[{sheetName}] {stageName} - 編號 {photoKey} ({cellAddress})：插入照片失敗 ({ex.Message})");
                            matched = true;
                            break;
                        }
                    }
                }
            }

            if (!matched)
            {
                warnings.Add($"[{sheetName}] {stageName} - 編號 {photoKey} ({cellAddress})：該格照片缺失");
            }
        }
    }
}
