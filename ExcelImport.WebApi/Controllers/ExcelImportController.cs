using ExcelImport.Core.Models;
using ExcelImport.Core.Services;
using ExcelImport.WebApi.Models;
using ExcelImport.WebApi.Services;
using Microsoft.AspNetCore.Mvc;

namespace ExcelImport.WebApi.Controllers;

[ApiController]
[Route("api/excel-import")]
public sealed class ExcelImportController : ControllerBase
{
    private readonly ILogger<ExcelImportController> _logger;
    private readonly ConfigService _configService;
    private readonly RecordFormatterService _recordFormatterService;
    private readonly SqlServerService _sqlServerService;
    private readonly string _connectionString;

    public ExcelImportController(
        ILogger<ExcelImportController> logger,
        ConfigService configService,
        RecordFormatterService recordFormatterService,
        SqlServerService sqlServerService,
        IConfiguration configuration)
    {
        _logger = logger;
        _configService = configService;
        _recordFormatterService = recordFormatterService;
        _sqlServerService = sqlServerService;
        _connectionString = configuration["Database:ConnectionString"] ?? string.Empty;
    }

    [HttpPost]
    public async Task<IActionResult> Import([FromBody] ExcelImportRequest request, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(_connectionString))
        {
            throw new InvalidOperationException("未配置 Database:ConnectionString。");
        }

        if (request.Records.Count == 0)
        {
            return BadRequest(new WebApiImportResponse
            {
                Success = false,
                Message = "Records 不能为空。"
            });
        }

        if (string.IsNullOrWhiteSpace(request.TemplateFile))
        {
            return BadRequest(new WebApiImportResponse
            {
                Success = false,
                Message = "TemplateFile 不能为空。"
            });
        }

        var temp = _configService.LoadTemplate(request.TemplateFile);
        var formattedRecords = _recordFormatterService.FormatRecords(request.Records, temp);
        var records = formattedRecords.Select(x => x.Values).ToList();

        _logger.LogInformation(
            "Received import request. Template={TemplateName}, File={FileName}, TemplateFile={TemplateFile}, TargetTable={TargetTable}, RecordCount={RecordCount}",
            request.TemplateName,
            request.FileName,
            request.TemplateFile,
            request.TargetTable,
            records.Count);

        var insertedCount = await _sqlServerService.InsertAsync(_connectionString, request.TargetTable, temp, records, cancellationToken);
        var skippedCount = records.Count - insertedCount;

        return Ok(new WebApiImportResponse
        {
            Success = true,
            Message = $"接收并按 Template 格式化入库成功，插入 {insertedCount} 条，重复跳过 {skippedCount} 条",
            TemplateName = request.TemplateName,
            FileName = request.FileName,
            TemplateFile = request.TemplateFile,
            TargetTable = request.TargetTable,
            RecordCount = records.Count,
            InsertedCount = insertedCount
        });
    }
}
