// ==============================================================================
// 檔案名稱：Program.cs
// 學習重點：
// 1. 單一無黑框桌面化入口 (`OutputType = WinExe` + `EmbeddedResource`)：
//    將前端網頁 `插圖軟體.html` 直接編譯內嵌於 `插圖軟體.exe` 中，雙擊執行檔即靜默啟動本機微伺服器並自動開啟瀏覽器，
//    不再需要額外的 `run.bat` 或外露 `.html` 檔案，整包複製到任何電腦皆可零設定使用。
// 2. 分頁生命週期自動回收 (`/api/session` Server-Sent Events)：
//    由於隱藏了傳統 Console 視窗，透過瀏覽器常駐 SSE 連線追蹤活躍分頁數；當使用者關閉網頁分頁 5 秒後（排除 F5 重新整理），
//    背景微伺服器會自動呼叫 `StopApplication()` 安全退出並釋放所有記憶體與檔案鎖定。
// 3. 全記憶體零落地緩衝 (`MemoryBufferThreshold = int.MaxValue`) 與 NAS 多執行緒直連架構。
// ==============================================================================

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Hosting.Server.Features;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using PhotoToExcelTool.Models;
using PhotoToExcelTool.Services;
using YamlDotNet.Serialization;

var builder = WebApplication.CreateBuilder(new WebApplicationOptions
{
    Args = args,
    ContentRootPath = AppDomain.CurrentDomain.BaseDirectory
});

// WinExe 無主控台視窗模式下關閉預設 Console Logger，避免無效主控台控制代碼開銷
builder.Logging.ClearProviders();

// ------------------------------------------------------------------------------
// 1. 配置本機隨機 Port (http://127.0.0.1:0 防衝突)
// ------------------------------------------------------------------------------
builder.WebHost.UseUrls("http://127.0.0.1:0");

// ------------------------------------------------------------------------------
// 2. 提升大檔案上傳上限至 1GB，並強制 100% 純記憶體緩衝（禁止寫入 %TEMP% 暫存檔）
// ------------------------------------------------------------------------------
builder.Services.Configure<FormOptions>(options =>
{
    options.MultipartBodyLengthLimit = 1_073_741_824; // 1 GB
    options.MemoryBufferThreshold = int.MaxValue;     // 100% 記憶體緩衝，禁止落地寫入 %TEMP%
    options.ValueCountLimit = 8192;
});
builder.WebHost.ConfigureKestrel(serverOptions =>
{
    serverOptions.Limits.MaxRequestBodySize = 1_073_741_824; // 1 GB
});

// ------------------------------------------------------------------------------
// 3. 註冊核心服務與執行期狀態追蹤
// ------------------------------------------------------------------------------
builder.Services.AddSingleton<ExcelImageService>();

var authorizedRoots = new ConcurrentDictionary<string, byte>(StringComparer.OrdinalIgnoreCase);
int activeSessions = 0;
int activeBusyOperations = 0;
bool hasEverConnected = false;

var app = builder.Build();

// ------------------------------------------------------------------------------
// 4. 首頁路由：開發時若存在外部 插圖軟體.html 可即時熱載入，正式部署則直接由 .exe 內嵌資源高速供應
// ------------------------------------------------------------------------------
async Task ServeIndexHtml(HttpContext context)
{
    context.Response.Headers["Cache-Control"] = "no-cache, no-store, must-revalidate";
    context.Response.Headers["Pragma"] = "no-cache";
    context.Response.ContentType = "text/html; charset=utf-8";

    string? externalHtml = FindExternalHtmlOverridePath();
    if (externalHtml != null)
    {
        await context.Response.SendFileAsync(externalHtml);
        return;
    }

    using Stream? embeddedStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("插圖軟體.html");
    if (embeddedStream != null)
    {
        await embeddedStream.CopyToAsync(context.Response.Body);
        return;
    }

    context.Response.StatusCode = 404;
    await context.Response.WriteAsync("找不到內嵌介面資源 插圖軟體.html");
}

app.MapGet("/", ServeIndexHtml);
app.MapGet("/插圖軟體.html", ServeIndexHtml);
app.MapGet("/index.html", ServeIndexHtml);

// ------------------------------------------------------------------------------
// 4.1 分頁存活常駐連線 (SSE)：關閉瀏覽器分頁後自動結束無黑框背景程序
// ------------------------------------------------------------------------------
app.MapGet("/api/session", async (HttpContext context) =>
{
    hasEverConnected = true;
    Interlocked.Increment(ref activeSessions);

    context.Response.Headers["Content-Type"] = "text/event-stream";
    context.Response.Headers["Cache-Control"] = "no-cache";
    context.Response.Headers["Connection"] = "keep-alive";

    await context.Response.WriteAsync(": connected\n\n");
    await context.Response.Body.FlushAsync();

    try
    {
        while (!context.RequestAborted.IsCancellationRequested &&
               !app.Lifetime.ApplicationStopping.IsCancellationRequested)
        {
            await Task.Delay(TimeSpan.FromSeconds(25), context.RequestAborted);
            await context.Response.WriteAsync(": ping\n\n");
            await context.Response.Body.FlushAsync();
        }
    }
    catch (OperationCanceledException)
    {
        // 分頁關閉或重新整理時正常中斷
    }
    finally
    {
        if (Interlocked.Decrement(ref activeSessions) <= 0)
        {
            _ = Task.Run(async () =>
            {
                // 等待 5 秒以區分「按 F5 重新整理」與「真正關閉所有分頁」
                await Task.Delay(TimeSpan.FromSeconds(5));
                if (Volatile.Read(ref activeSessions) <= 0 &&
                    Volatile.Read(ref activeBusyOperations) <= 0)
                {
                    app.Lifetime.StopApplication();
                }
            });
        }
    }
});

// ==============================================================================
// API 路由契約 (Minimal API Endpoints)
// ==============================================================================

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

app.MapGet("/config.yml", () =>
{
    string configPath = GetConfigFilePath();
    if (!File.Exists(configPath))
    {
        return Results.NotFound("config.yml not found");
    }
    return Results.File(configPath, "text/yaml; charset=utf-8");
});

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

app.MapPost("/api/browse-folder", async (HttpRequest request) =>
{
    Interlocked.Increment(ref activeBusyOperations);
    try
    {
        string? initialDir = null;
        try
        {
            var body = await JsonSerializer.DeserializeAsync<BrowseFolderRequest>(request.Body);
            initialDir = body?.InitialDirectory;
        }
        catch
        {
        }

        string? selected = WindowsFolderPicker.PickFolder(initialDir);
        if (string.IsNullOrWhiteSpace(selected))
        {
            return Results.Ok(new { cancelled = true });
        }

        string cleanRoot = NormalizeRootPath(selected);
        authorizedRoots[cleanRoot] = 1;
        return Results.Ok(new { cancelled = false, rootPath = cleanRoot });
    }
    catch (Exception ex)
    {
        return Results.Problem(statusCode: 500, detail: ex.Message);
    }
    finally
    {
        Interlocked.Decrement(ref activeBusyOperations);
    }
}).DisableAntiforgery();

app.MapPost("/api/scan-folder", async (HttpRequest request, ExcelImageService imageService) =>
{
    ScanFolderRequest? req;
    try
    {
        req = await JsonSerializer.DeserializeAsync<ScanFolderRequest>(request.Body);
    }
    catch (Exception ex)
    {
        return Results.BadRequest(new { error = $"請求格式錯誤: {ex.Message}" });
    }

    if (req == null || string.IsNullOrWhiteSpace(req.RootPath))
    {
        return Results.BadRequest(new { error = "請提供有效的照片根資料夾路徑 (rootPath)。" });
    }

    string cleanRoot;
    try
    {
        cleanRoot = NormalizeRootPath(req.RootPath);
    }
    catch (Exception ex)
    {
        return Results.BadRequest(new { error = $"路徑格式不合法: {ex.Message}" });
    }

    if (!Directory.Exists(cleanRoot))
    {
        return Results.BadRequest(new { error = $"找不到指定的資料夾路徑：{cleanRoot}" });
    }

    var folders = new List<string>();
    if (req.Folders != null) folders.AddRange(req.Folders);

    var exts = new List<string>();
    if (req.AcceptExtensions != null) exts.AddRange(req.AcceptExtensions);

    AppConfig? parsedCfg = TryParseConfigOrFallback(req.ConfigYaml);
    if (parsedCfg != null)
    {
        foreach (var r in parsedCfg.Rules)
        {
            if (!string.IsNullOrWhiteSpace(r.IQC?.Folder)) folders.Add(r.IQC!.Folder);
            if (!string.IsNullOrWhiteSpace(r.FQC?.Folder)) folders.Add(r.FQC!.Folder);
        }
        if (exts.Count == 0 && parsedCfg.AcceptExtension != null)
        {
            exts.AddRange(parsedCfg.AcceptExtension);
        }
    }

    var sw = Stopwatch.StartNew();
    try
    {
        var files = imageService.ScanRuleFolders(cleanRoot, folders, exts);
        sw.Stop();

        authorizedRoots[cleanRoot] = 1;
        string rootFolderName = Path.GetFileName(cleanRoot.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
        if (string.IsNullOrWhiteSpace(rootFolderName)) rootFolderName = cleanRoot;

        return Results.Ok(new
        {
            rootPath = cleanRoot,
            rootFolderName,
            elapsedMs = sw.ElapsedMilliseconds,
            files
        });
    }
    catch (Exception ex)
    {
        return Results.BadRequest(new { error = ex.Message });
    }
}).DisableAntiforgery();

app.MapGet("/api/photo-file", (string path, HttpResponse response) =>
{
    if (string.IsNullOrWhiteSpace(path))
    {
        return Results.BadRequest("Missing path");
    }

    string fullPath;
    try
    {
        fullPath = Path.GetFullPath(path.Trim().Trim('"'));
    }
    catch
    {
        return Results.BadRequest("Invalid path");
    }

    bool isAuthorized = authorizedRoots.Keys.Any(root =>
        fullPath.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(fullPath, root, StringComparison.OrdinalIgnoreCase));

    if (!isAuthorized)
    {
        return Results.StatusCode(StatusCodes.Status403Forbidden);
    }

    string ext = Path.GetExtension(fullPath).ToLowerInvariant();
    string? contentType = ext switch
    {
        ".jpg" or ".jpeg" => "image/jpeg",
        ".png" => "image/png",
        ".gif" => "image/gif",
        ".bmp" => "image/bmp",
        ".webp" => "image/webp",
        _ => null
    };

    if (contentType == null || !File.Exists(fullPath))
    {
        return Results.NotFound();
    }

    response.Headers["Cache-Control"] = "private, max-age=300";

    var stream = new FileStream(
        fullPath,
        FileMode.Open,
        FileAccess.Read,
        FileShare.ReadWrite,
        bufferSize: 65536,
        options: FileOptions.SequentialScan | FileOptions.Asynchronous);

    return Results.Stream(stream, contentType);
});

app.MapPost("/api/export-excel", async (
    HttpRequest request,
    HttpResponse response,
    ExcelImageService imageService) =>
{
    Interlocked.Increment(ref activeBusyOperations);
    try
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

        AppConfig? config = TryParseConfigOrFallback(form["configYaml"]);
        if (config == null)
        {
            return Results.Problem(statusCode: 500, detail: "找不到或無法解析 config.yml 設定檔。");
        }

        byte[] fileBytes;
        List<string> warnings;

        string? rootFolderPath = form["rootFolderPath"].FirstOrDefault();
        if (!string.IsNullOrWhiteSpace(rootFolderPath))
        {
            string cleanRoot = NormalizeRootPath(rootFolderPath);
            if (!Directory.Exists(cleanRoot))
            {
                return Results.BadRequest(new { error = $"找不到照片根資料夾：{cleanRoot}" });
            }
            authorizedRoots[cleanRoot] = 1;
            (fileBytes, warnings) = imageService.ProcessExcelFromDirectory(templateFile, cleanRoot, config);
        }
        else
        {
            var photos = form.Files.GetFiles("photos");
            var relativePaths = form["relativePaths"].Select(s => s ?? string.Empty).ToList();
            (fileBytes, warnings) = imageService.ProcessExcel(templateFile, photos, relativePaths, config);
        }

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
    }
    finally
    {
        Interlocked.Decrement(ref activeBusyOperations);
    }
}).DisableAntiforgery();

// ------------------------------------------------------------------------------
// 5. 伺服器啟動完成事件：靜默開啟系統預設瀏覽器，並啟動閒置保護機制
// ------------------------------------------------------------------------------
app.Lifetime.ApplicationStarted.Register(() =>
{
    var serverAddressesFeature = app.Services.GetRequiredService<IServer>().Features.Get<IServerAddressesFeature>();
    var address = serverAddressesFeature?.Addresses.FirstOrDefault();
    if (!string.IsNullOrEmpty(address))
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = address,
                UseShellExecute = true
            });
        }
        catch
        {
        }
    }

    // 啟動保護：若 60 秒內瀏覽器從未連入，自動結束背景程序避免殘留
    _ = Task.Run(async () =>
    {
        await Task.Delay(TimeSpan.FromSeconds(60));
        if (!hasEverConnected && Volatile.Read(ref activeSessions) <= 0)
        {
            app.Lifetime.StopApplication();
        }
    });
});

app.Run();

// ==============================================================================
// 輔助方法與請求模型 (Local Helper Functions & Request DTOs)
// ==============================================================================

static string NormalizeRootPath(string rawPath)
{
    string cleaned = rawPath.Trim().Trim('"');
    string full = Path.GetFullPath(cleaned);
    return full.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
}

static string? FindExternalHtmlOverridePath()
{
    const string uiFileName = "插圖軟體.html";
    string baseDir = AppDomain.CurrentDomain.BaseDirectory;

    // 1. 若同級目錄下有放置外部 插圖軟體.html 則優先讀取
    string path1 = Path.Combine(baseDir, uiFileName);
    if (File.Exists(path1)) return path1;

    // 2. 本機開發環境下，若 PhotoToExcelTool\插圖軟體.html 存在則直接讀取方便即時修改預覽
    string path2 = Path.Combine(baseDir, "PhotoToExcelTool", uiFileName);
    if (File.Exists(path2)) return path2;

    return null;
}

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

static AppConfig? TryParseConfigOrFallback(string? rawYaml)
{
    if (!string.IsNullOrWhiteSpace(rawYaml))
    {
        try
        {
            var deserializer = new DeserializerBuilder()
                .IgnoreUnmatchedProperties()
                .Build();
            return deserializer.Deserialize<AppConfig>(rawYaml) ?? new AppConfig();
        }
        catch
        {
        }
    }

    string configPath = GetConfigFilePath();
    if (!File.Exists(configPath)) return null;

    try
    {
        return LoadConfig(configPath);
    }
    catch
    {
        return null;
    }
}

public record BrowseFolderRequest(
    [property: JsonPropertyName("initialDirectory")] string? InitialDirectory
);

public record ScanFolderRequest(
    [property: JsonPropertyName("rootPath")] string RootPath,
    [property: JsonPropertyName("folders")] List<string>? Folders,
    [property: JsonPropertyName("acceptExtensions")] List<string>? AcceptExtensions,
    [property: JsonPropertyName("configYaml")] string? ConfigYaml
);
