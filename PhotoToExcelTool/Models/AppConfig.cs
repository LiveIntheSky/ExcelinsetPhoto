// ==============================================================================
// 檔案名稱：Models/AppConfig.cs
// 學習重點：
// 1. 資料模型 (POCO / DTO)：將外部設定檔 (config.yml) 映射為強型別 C# 物件，避免在程式中到處使用魔術字串 (Magic Strings)。
// 2. 雙重序列化標籤 (Attributes)：
//    - [YamlMember(Alias = "...")]：供 YamlDotNet 套件使用，將 YAML 的蛇形命名 (snake_case，如 accept_extension) 對應到 C# 的帕斯卡命名 (PascalCase，如 AcceptExtension)。
//    - [JsonPropertyName("...")]：供 System.Text.Json 套件使用，當後端透過 API 回傳給前端 JS 時，自動轉為駝峰命名 (camelCase，如 acceptExtension)。
// 3. 屬性預設值 (Property Initializers)：使用 `= new()` 或預設數值，防止反序列化時遇到空值導致 NullReferenceException。
// ==============================================================================

using System.Collections.Generic;
using System.Text.Json.Serialization;
using YamlDotNet.Serialization;

namespace PhotoToExcelTool.Models;

/// <summary>
/// 對應 config.yml 最外層根節點的設定類別。
/// </summary>
public class AppConfig
{
    /// <summary>
    /// 允許讀取的照片副檔名白名單（例如：.jpg, .png）。
    /// YamlDotNet 讀取 YAML 中的 "accept_extension"，JSON 輸出時轉為 "acceptExtension"。
    /// </summary>
    [YamlMember(Alias = "accept_extension")]
    [JsonPropertyName("acceptExtension")]
    public List<string> AcceptExtension { get; set; } = new() { ".jpg", ".png" };

    /// <summary>
    /// 所有工作表的插圖規則清單，對應 YAML 中的 "rules" 陣列。
    /// </summary>
    [YamlMember(Alias = "rules")]
    [JsonPropertyName("rules")]
    public List<SheetRule> Rules { get; set; } = new();
}

/// <summary>
/// 單一 Excel 工作表 (Worksheet) 的插圖規則。
/// 每個 SheetRule 定義了目標工作表名稱、該表照片尺寸，以及 IQC / FQC 兩大檢驗階段的資料夾與儲存格對照。
/// </summary>
public class SheetRule
{
    /// <summary>
    /// 目標 Excel 工作表名稱（必須與 Excel 底稿下方的分頁名稱完全一致，例如 "Ring #1" 或 "Susceptor Sidewalls"）。
    /// </summary>
    [YamlMember(Alias = "sheet")]
    [JsonPropertyName("sheet")]
    public string Sheet { get; set; } = string.Empty;

    /// <summary>
    /// 該工作表內插入照片的尺寸與邊距設定（寬、高、左留白、上留白）。
    /// 若 YAML 未特別填寫 size 區塊，會自動套用 ImageSize 類別中的預設值。
    /// </summary>
    [YamlMember(Alias = "size")]
    [JsonPropertyName("size")]
    public ImageSize Size { get; set; } = new();

    /// <summary>
    /// IQC（進料/清洗前檢驗）階段的資料夾與儲存格位置對照。
    /// 使用可為 Null 的參考型別 (StageConfig?)，代表若某張工作表沒有 IQC 區塊，程式可安全跳過不處理。
    /// </summary>
    [YamlMember(Alias = "IQC")]
    [JsonPropertyName("iqc")]
    public StageConfig? IQC { get; set; }

    /// <summary>
    /// FQC（最終/清洗後檢驗）階段的資料夾與儲存格位置對照。
    /// </summary>
    [YamlMember(Alias = "FQC")]
    [JsonPropertyName("fqc")]
    public StageConfig? FQC { get; set; }
}

/// <summary>
/// 照片尺寸與儲存格邊距設定模型。
/// 學習重點：Excel 中的圖片定位是以像素 (px) 或 EMU 為單位，但工程人員習慣使用公分 (cm) 與毫米 (mm)，
/// 因此在設定檔使用直覺的物理單位，再由服務層 (ExcelImageService) 依據 96 DPI 標準進行換算。
/// </summary>
public class ImageSize
{
    /// <summary>
    /// 照片寬度（單位：公分 cm）。預設 8.50 cm。
    /// </summary>
    [YamlMember(Alias = "width")]
    [JsonPropertyName("width")]
    public double Width { get; set; } = 8.50;

    /// <summary>
    /// 照片高度（單位：公分 cm）。預設 11.00 cm。
    /// </summary>
    [YamlMember(Alias = "height")]
    [JsonPropertyName("height")]
    public double Height { get; set; } = 11.00;

    /// <summary>
    /// 照片距離目標儲存格左邊框的留白距離（單位：毫米 mm）。預設 2.0 mm。
    /// 用途：避免插入的照片直接壓在儲存格左側格線上。
    /// </summary>
    [YamlMember(Alias = "left_padding")]
    [JsonPropertyName("leftPadding")]
    public double LeftPadding { get; set; } = 2.0;

    /// <summary>
    /// 照片距離目標儲存格上邊框的留白距離（單位：毫米 mm）。預設 1.0 mm（約 4 px）。
    /// 用途：避免插入的照片直接壓在儲存格上方格線上。
    /// </summary>
    [YamlMember(Alias = "top_padding")]
    [JsonPropertyName("topPadding")]
    public double TopPadding { get; set; } = 1.0;
}

/// <summary>
/// 單一檢驗階段（如 IQC 或 FQC）的目標子資料夾與儲存格映射設定。
/// </summary>
public class StageConfig
{
    /// <summary>
    /// 照片所在的相對子資料夾路徑（例如："PRE/01" 或 "FQC/01"）。
    /// </summary>
    [YamlMember(Alias = "folder")]
    [JsonPropertyName("folder")]
    public string Folder { get; set; } = string.Empty;

    /// <summary>
    /// 照片檔名編號（不含副檔名）與目標 Excel 儲存格座標的鍵值對 (Dictionary)。
    /// Key：照片編號字串（例如 "1", "2" 或 "Pre1", "Pre2"）。
    /// Value：Excel 儲存格 A1 座標（例如 "B7", "F28"）。
    /// </summary>
    [YamlMember(Alias = "cells")]
    [JsonPropertyName("cells")]
    public Dictionary<string, string> Cells { get; set; } = new();
}
