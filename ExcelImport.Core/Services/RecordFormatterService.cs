using ExcelImport.Core.Models;
using ExcelImport.Core.Models.Template;

namespace ExcelImport.Core.Services;

public sealed class RecordFormatterService
{
    public List<DynamicEntityRecord> FormatRecords(IEnumerable<Dictionary<string, object?>> records, ExcelTemplateDefinition template)
    {
        return records.Select(record => new DynamicEntityRecord
        {
            Values = FormatRecord(record, template)
        }).ToList();
    }

    public Dictionary<string, object?> FormatRecord(Dictionary<string, object?> record, ExcelTemplateDefinition template)
    {
        var formatted = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);

        foreach (var field in template.Fields)
        {
            record.TryGetValue(field.Name, out var value);
            formatted[field.Name] = ConvertValue(field, value);
        }

        return formatted;
    }

    private static object? ConvertValue(ExcelTemplateFieldDefinition field, object? value)
    {
        if (value is null)
        {
            if (field.Required)
            {
                throw new InvalidOperationException($"字段 {field.Name} 不允许为空。");
            }

            return null;
        }

        var text = value.ToString()?.Trim();
        if (string.IsNullOrWhiteSpace(text))
        {
            if (field.Required)
            {
                throw new InvalidOperationException($"字段 {field.Name} 不允许为空。");
            }

            return null;
        }

        object converted = field.Type.ToLowerInvariant() switch
        {
            "int" => int.Parse(text, System.Globalization.CultureInfo.InvariantCulture),
            "decimal" => decimal.Parse(text, System.Globalization.CultureInfo.InvariantCulture),
            "datetime" => DateTime.Parse(text, System.Globalization.CultureInfo.InvariantCulture),
            "bool" => ParseBoolean(text),
            "string" => text,
            _ => text
        };

        if (converted is string s && field.Length is int maxLength && s.Length > maxLength)
        {
            throw new InvalidOperationException($"字段 {field.Name} 长度超过限制 {maxLength}。");
        }

        return converted;
    }

    private static bool ParseBoolean(string text)
    {
        return text.Trim().ToUpperInvariant() switch
        {
            "TRUE" or "1" or "Y" or "YES" or "PASS" or "通过" => true,
            "FALSE" or "0" or "N" or "NO" or "FAIL" or "不通过"=> false,
            _ => bool.Parse(text)
        };
    }
}
