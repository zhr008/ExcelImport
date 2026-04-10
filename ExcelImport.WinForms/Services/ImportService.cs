using ExcelImport.Core.Models;
using ExcelImport.Core.Services;

namespace ExcelImport.Services;

public sealed class ImportService
{
    private readonly ConfigService _configService;
    private readonly ExcelReaderService _excelReaderService;
    private readonly RecordFormatterService _recordFormatterService;
    private readonly SqlServerService _sqlServerService;
    private readonly WebApiService _webApiService;
    private readonly LoggingService _loggingService;

    public ImportService(
        ConfigService configService,
        ExcelReaderService excelReaderService,
        RecordFormatterService recordFormatterService,
        SqlServerService sqlServerService,
        WebApiService webApiService,
        LoggingService loggingService)
    {
        _configService = configService;
        _excelReaderService = excelReaderService;
        _recordFormatterService = recordFormatterService;
        _sqlServerService = sqlServerService;
        _webApiService = webApiService;
        _loggingService = loggingService;
    }

    public async Task<ImportResult> RunTemplateAsync(AppSettings appSettings, TemplateTaskConfig template, CancellationToken cancellationToken)
    {
        var result = new ImportResult();
        _loggingService.Info($"开始处理模板: {template.Name}，间隔: {template.IntervalMinutes} 分钟，目录: {template.WatchPath}");

        if (!template.Enabled)
        {
            _loggingService.Warning($"模板未启用，跳过执行: {template.Name}");
            return result;
        }

        if (!appSettings.Database.Enabled && !appSettings.WebApi.Enabled)
        {
            throw new InvalidOperationException("未启用任何采集方式，请在 appsettings.json 中启用 Database 或 WebApi。");
        }

        if (!Directory.Exists(template.WatchPath))
        {
            _loggingService.Warning($"目录不存在: {template.WatchPath}");
            return result;
        }

        var temp = _configService.LoadTemplate(template.TemplateFile);
        var option = template.IncludeSubdirectories ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
        var files = Directory.GetFiles(template.WatchPath, template.FilePattern, option)
            .Where(file => !IsInSystemFolder(file, template.WatchPath))
            .ToList();

        result.FilesScanned = files.Count;
        if (files.Count == 0)
        {
            _loggingService.Info($"模板 {template.Name} 本轮未发现匹配文件。目录: {template.WatchPath}，模式: {template.FilePattern}");
        }
        _loggingService.Info($"开始执行模板 {template.Name}，扫描到 {files.Count} 个文件。");

        foreach (var file in files)
        {
            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                var rawRecords = _excelReaderService.Read(file, temp);
                var formattedRecords = _recordFormatterService.FormatRecords(rawRecords, temp);
                var records = formattedRecords.Select(x => x.Values).ToList();
                var importedCount = 0;

                if (appSettings.Database.Enabled)
                {
                    importedCount = await _sqlServerService.InsertAsync(appSettings.Database.ConnectionString, template.TargetTable, temp, records, cancellationToken);
                    var skippedCount = records.Count - importedCount;
                    _loggingService.Info($"数据库导入完成: {file}，总记录数: {records.Count}，插入: {importedCount}，重复跳过: {skippedCount}");
                }

                if (appSettings.WebApi.Enabled)
                {
                    var payload = new WebApiPayload
                    {
                        TemplateName = template.Name,
                        FileName = Path.GetFileName(file),
                        TemplateFile = template.TemplateFile,
                        TargetTable = template.TargetTable,
                        Records = records
                    };

                    var apiResult = await _webApiService.SendAsync(appSettings.WebApi, payload, cancellationToken);
                    importedCount = apiResult.InsertedCount;
                    var skippedCount = apiResult.RecordCount - apiResult.InsertedCount;
                    _loggingService.Info($"WebApi 导入成功: {file}，消息: {apiResult.Message}，接收记录数: {apiResult.RecordCount}，插入: {apiResult.InsertedCount}，重复跳过: {skippedCount}");
                }

                result.FilesSucceeded++;
                result.RecordsImported += importedCount;
                MoveToFolder(file, template.WatchPath, true, _loggingService);
            }
            catch (Exception ex)
            {
                result.FilesFailed++;
                MoveToFolder(file, template.WatchPath, false, _loggingService);
                _loggingService.Error($"文件导入失败: {file}", ex);
            }
        }

        _loggingService.Info($"模板执行完成: {template.Name}，扫描: {result.FilesScanned}，成功: {result.FilesSucceeded}，失败: {result.FilesFailed}，记录: {result.RecordsImported}");
        return result;
    }

    private static bool IsInSystemFolder(string filePath, string watchPath)
    {
        var relativePath = Path.GetRelativePath(watchPath, filePath);
        return relativePath.StartsWith("Succeed", StringComparison.OrdinalIgnoreCase)
               || relativePath.StartsWith("Failed", StringComparison.OrdinalIgnoreCase);
    }

    private static void MoveToFolder(string filePath, string watchPath, bool success, LoggingService loggingService)
    {
        var lastWriteDate = File.GetLastWriteTime(filePath).Date;
        if (lastWriteDate == DateTime.Today)
        {
            loggingService.Info($"文件修改日期为当天，保留原文件不移动: {filePath}");
            return;
        }

        var folderName = success ? "Succeed" : "Failed";
        var targetDirectory = Path.Combine(watchPath, folderName);
        Directory.CreateDirectory(targetDirectory);

        var targetPath = Path.Combine(targetDirectory, Path.GetFileName(filePath));
        if (File.Exists(targetPath))
        {
            var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);
            var extension = Path.GetExtension(filePath);
            targetPath = Path.Combine(targetDirectory, $"{fileNameWithoutExtension}_{DateTime.Now:yyyyMMddHHmmss}{extension}");
        }

        File.Move(filePath, targetPath);
    }
}
