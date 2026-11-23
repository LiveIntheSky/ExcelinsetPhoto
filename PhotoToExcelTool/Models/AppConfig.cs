using System.Collections.Generic;
using System.Text.Json.Serialization;
using YamlDotNet.Serialization;

namespace PhotoToExcelTool.Models;

public class AppConfig
{
    [YamlMember(Alias = "accept_extension")]
    [JsonPropertyName("acceptExtension")]
    public List<string> AcceptExtension { get; set; } = new() { ".jpg", ".png" };

    [YamlMember(Alias = "rules")]
    [JsonPropertyName("rules")]
    public List<SheetRule> Rules { get; set; } = new();
}

public class SheetRule
{
    [YamlMember(Alias = "sheet")]
    [JsonPropertyName("sheet")]
    public string Sheet { get; set; } = string.Empty;

    [YamlMember(Alias = "size")]
    [JsonPropertyName("size")]
    public ImageSize Size { get; set; } = new();

    [YamlMember(Alias = "IQC")]
    [JsonPropertyName("iqc")]
    public StageConfig? IQC { get; set; }

    [YamlMember(Alias = "FQC")]
    [JsonPropertyName("fqc")]
    public StageConfig? FQC { get; set; }
}

public class ImageSize
{
    [YamlMember(Alias = "width")]
    [JsonPropertyName("width")]
    public double Width { get; set; } = 8.50;

    [YamlMember(Alias = "height")]
    [JsonPropertyName("height")]
    public double Height { get; set; } = 11.00;

    [YamlMember(Alias = "left_padding")]
    [JsonPropertyName("leftPadding")]
    public double LeftPadding { get; set; } = 2.0;

    [YamlMember(Alias = "top_padding")]
    [JsonPropertyName("topPadding")]
    public double TopPadding { get; set; } = 1.0;
}

public class StageConfig
{
    [YamlMember(Alias = "folder")]
    [JsonPropertyName("folder")]
    public string Folder { get; set; } = string.Empty;

    [YamlMember(Alias = "cells")]
    [JsonPropertyName("cells")]
    public Dictionary<string, string> Cells { get; set; } = new();
}
