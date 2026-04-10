namespace ExcelImport.Core.Models;

public sealed class WebApiPayload
{
    public string TemplateName { get; set; } = string.Empty;

    public string FileName { get; set; } = string.Empty;

    public string TemplateFile { get; set; } = string.Empty;

    public string TargetTable { get; set; } = string.Empty;

    public List<Dictionary<string, object?>> Records { get; set; } = [];
}
