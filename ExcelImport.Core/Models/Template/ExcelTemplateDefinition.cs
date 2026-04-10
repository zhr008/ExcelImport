namespace ExcelImport.Core.Models.Template;

public sealed class ExcelTemplateDefinition
{
    public ExcelTemplateReadMode ReadMode { get; set; }

    public int? StartRow { get; set; }

    public List<ExcelTemplateFieldDefinition> Fields { get; set; } = [];
}
