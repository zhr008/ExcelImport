namespace ExcelImport.Core.Models;

public sealed class DynamicEntityRecord
{
    public Dictionary<string, object?> Values { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}
