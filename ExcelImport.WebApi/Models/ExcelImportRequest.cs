using ExcelImport.Core.Models;

namespace ExcelImport.WebApi.Models;

public sealed class ExcelImportRequest
{
    public string TemplateName { get; set; } = string.Empty;

    public string FileName { get; set; } = string.Empty;

    public string TemplateFile { get; set; } = string.Empty;

    public string TargetTable { get; set; } = string.Empty;

    public List<Dictionary<string, object?>> Records { get; set; } = [];

    public WebApiPayload ToPayload()
    {
        return new WebApiPayload
        {
            TemplateName = TemplateName,
            FileName = FileName,
            TemplateFile = TemplateFile,
            TargetTable = TargetTable,
            Records = Records
        };
    }
}
