namespace ExcelImport.Core.Models;

public sealed class WebApiImportResponse
{
    public bool Success { get; set; }

    public string Message { get; set; } = string.Empty;

    public string TemplateName { get; set; } = string.Empty;

    public string FileName { get; set; } = string.Empty;

    public string TemplateFile { get; set; } = string.Empty;

    public string TargetTable { get; set; } = string.Empty;

    public int RecordCount { get; set; }

    public int InsertedCount { get; set; }
}
