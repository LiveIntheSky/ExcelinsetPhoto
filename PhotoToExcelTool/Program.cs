using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Hosting.Server.Features;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using PhotoToExcelTool.Models;
using PhotoToExcelTool.Services;
using YamlDotNet.Serialization;

var builder = WebApplication.CreateBuilder(new WebApplicationOptions
{
    Args = args,
    ContentRootPath = AppDomain.CurrentDomain.BaseDirectory
});

// 1. 配置本機隨機 Port (http://127.0.0.1:0 防衝突)
builder.WebHost.UseUrls("http://127.0.0.1:0");

// 2. 提升大檔案 (高解析照片批次) 上傳上限為 500MB
builder.Services.Configure<FormOptions>(options =>
{
    options.MultipartBodyLengthLimit = 524_288_000;
});
builder.WebHost.ConfigureKestrel(serverOptions =>
{
    serverOptions.Limits.MaxRequestBodySize = 524_288_000;
});

// 3. 註冊核心服務
builder.Services.AddSingleton<ExcelImageService>();

var app = builder.Build();

// 4. 啟用靜態檔案 (wwwroot)
app.UseDefaultFiles();
app.UseStaticFiles();

// 確保訪問首頁 "/" 時優先載入 index.html (相容根目錄或 wwwroot 目錄)
app.MapGet("/", async (context) =>
{
    string indexPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "index.html");
    if (File.Exists(indexPath))
    {
        context.Response.ContentType = "text/html; charset=utf-8";
        await context.Response.SendFileAsync(indexPath);
        return;
    }
    string wwwrootIndex = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot", "index.html");
    if (File.Exists(wwwrootIndex))
    {
        context.Response.ContentType = "text/html; charset=utf-8";
        await context.Response.SendFileAsync(wwwrootIndex);
        return;
    }
    context.Response.StatusCode = 404;
    await context.Response.WriteAsync("找不到 index.html");
});

// ==============================================================================
// API 路由契約
// ==============================================================================

// A. GET /api/config
app.MapGet("/api/config", () =>
{
    string configPath = GetConfigFilePath();
    if (!File.Exists(configPath))
    {
        return Results.Problem(statusCode: 500, detail: $"找不到設定檔: {configPath}");
    }

    try
    {
        var config = LoadConfig(configPath);
        return Results.Ok(config);
    }
    catch (Exception ex)
    {
        return Results.Problem(statusCode: 500, detail: $"解析 config.yml 失敗: {ex.Message}");
    }
});

// A2. GET /config.yml
app.MapGet("/config.yml", () =>
{
    string configPath = GetConfigFilePath();
    if (!File.Exists(configPath))
    {
        return Results.NotFound("config.yml not found");
    }
    return Results.File(configPath, "text/yaml; charset=utf-8");
});

// A3. POST /api/config (動態儲存前端編輯的 config.yml 全文內容)
app.MapPost("/api/config", async (HttpRequest request) =>
{
    using var reader = new StreamReader(request.Body, Encoding.UTF8);
    string newYaml = await reader.ReadToEndAsync();
    if (string.IsNullOrWhiteSpace(newYaml))
    {
        return Results.BadRequest(new { error = "設定檔內容不可為空。" });
    }

    try
    {
        var deserializer = new DeserializerBuilder()
            .IgnoreUnmatchedProperties()
            .Build();
        deserializer.Deserialize<AppConfig>(newYaml);
    }
    catch (Exception ex)
    {
        return Results.BadRequest(new { error = $"YAML 格式語法驗證失敗: {ex.Message}" });
    }

    string configPath = GetConfigFilePath();
    await File.WriteAllTextAsync(configPath, newYaml, new UTF8Encoding(false));
    return Results.Ok(new { message = "config.yml 儲存成功！", path = configPath });
}).DisableAntiforgery();

// B. POST /api/export-excel
app.MapPost("/api/export-excel", async (
    HttpRequest request,
    HttpResponse response,
    ExcelImageService imageService) =>
{
    if (!request.HasFormContentType)
    {
        return Results.BadRequest(new { error = "請求必須為 multipart/form-data 格式。" });
    }

    var form = await request.ReadFormAsync();
    var templateFile = form.Files.GetFile("templateFile");
    if (templateFile == null || templateFile.Length == 0)
    {
        return Results.BadRequest(new { error = "未選取 Excel 底稿檔案 (templateFile)。" });
    }

    var photos = form.Files.GetFiles("photos");
    var relativePaths = form["relativePaths"].Select(s => s ?? string.Empty).ToList();

    AppConfig config;
    if (form.ContainsKey("configYaml") && !string.IsNullOrWhiteSpace(form["configYaml"]))
    {
        try
        {
            var deserializer = new DeserializerBuilder()
                .IgnoreUnmatchedProperties()
                .Build();
            config = deserializer.Deserialize<AppConfig>(form["configYaml"]!) ?? new AppConfig();
        }
        catch
        {
            string configPath = GetConfigFilePath();
            config = File.Exists(configPath) ? LoadConfig(configPath) : new AppConfig();
        }
    }
    else
    {
        string configPath = GetConfigFilePath();
        if (!File.Exists(configPath))
        {
            return Results.Problem(statusCode: 500, detail: $"找不到設定檔: {configPath}");
        }

        try
        {
            config = LoadConfig(configPath);
        }
        catch (Exception ex)
        {
            return Results.Problem(statusCode: 500, detail: $"解析 config.yml 失敗: {ex.Message}");
        }
    }

    // 純記憶體 ClosedXML 批次處理
    var (fileBytes, warnings) = imageService.ProcessExcel(templateFile, photos, relativePaths, config);

    // 處理警告 Header (URI 編碼避免非 ASCII 字元錯誤)
    response.Headers["Access-Control-Expose-Headers"] = "X-Process-Warnings";
    if (warnings.Count > 0)
    {
        string warningsJson = JsonSerializer.Serialize(warnings);
        response.Headers["X-Process-Warnings"] = Uri.EscapeDataString(warningsJson);
    }

    string originalName = Path.GetFileNameWithoutExtension(templateFile.FileName);
    string downloadName = $"{originalName}_已插圖.xlsx";

    return Results.File(
        fileBytes,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        downloadName
    );
}).DisableAntiforgery();

// 5. 啟動完成後自動開啟系統預設瀏覽器
app.Lifetime.ApplicationStarted.Register(() =>
{
    var serverAddressesFeature = app.Services.GetRequiredService<IServer>().Features.Get<IServerAddressesFeature>();
    var address = serverAddressesFeature?.Addresses.FirstOrDefault();
    if (!string.IsNullOrEmpty(address))
    {
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("==========================================================================");
        Console.WriteLine("  🚀 SiC Excel 照片批次自動填入工具 (純記憶體執行 / 零落地微伺服器)");
        Console.WriteLine($"  🌐 服務已啟動: {address}");
        Console.WriteLine("  💡 已自動呼叫系統預設瀏覽器開啟網頁介面...");
        Console.WriteLine("==========================================================================");
        Console.ResetColor();

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = address,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[提示] 自動開啟瀏覽器失敗，請手動於瀏覽器開啟: {address} ({ex.Message})");
        }
    }
});

app.Run();

// ==============================================================================
// 輔助方法
// ==============================================================================

static string GetConfigFilePath()
{
    string baseDir = AppDomain.CurrentDomain.BaseDirectory;
    string path1 = Path.Combine(baseDir, "config.yml");
    if (File.Exists(path1)) return path1;

    string currentDir = Directory.GetCurrentDirectory();
    string path2 = Path.Combine(currentDir, "config.yml");
    if (File.Exists(path2)) return path2;

    return path1;
}

static AppConfig LoadConfig(string path)
{
    string yaml = File.ReadAllText(path, Encoding.UTF8);
    var deserializer = new DeserializerBuilder()
        .IgnoreUnmatchedProperties()
        .Build();
    return deserializer.Deserialize<AppConfig>(yaml) ?? new AppConfig();
}
