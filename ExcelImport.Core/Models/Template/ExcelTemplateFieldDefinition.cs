namespace ExcelImport.Core.Models.Template;

public sealed class ExcelTemplateFieldDefinition
{
    public string Name { get; set; } = string.Empty;

    public string? Column { get; set; }

    public string? Cell { get; set; }

    public string Type { get; set; } = "string";

    public bool Required { get; set; }

    public int? Length { get; set; }

    public bool IsKey { get; set; }

    public string? Format { get; set; }
}
