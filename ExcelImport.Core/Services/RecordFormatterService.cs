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

    public bool IsValidRecord(Dictionary<string, object?> record, ExcelTemplateDefinition template)
    {
        foreach (var field in template.Fields)
        {
            if (field.Required)
            {
                if (!record.TryGetValue(field.Name, out var value) || value is null)
                {
                    return false;
                }
                var text = value.ToString()?.Trim();
                if (string.IsNullOrWhiteSpace(text))
                {
                    return false;
                }
            }
        }
        return true;
    }

    private static object? ConvertValue(ExcelTemplateFieldDefinition field, object? value)
    {
        if (value is null)
        {
            if (field.Required)
            {
                return null;
            }

            return null;
        }

        var text = value.ToString()?.Trim();
        if (string.IsNullOrWhiteSpace(text))
        {
            if (field.Required)
            {
                return null;
            }

            return null;
        }

        object converted = field.Type.ToLowerInvariant() switch
        {
            "int" => int.Parse(text, System.Globalization.CultureInfo.InvariantCulture),
            "decimal" => decimal.Parse(text, System.Globalization.CultureInfo.InvariantCulture),
            "datetime" => ParseDateTime(field, text),
            "date" => ParseDate(field, text),
            "time" => ParseTime(text),
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

    private static string ParseDateTime(ExcelTemplateFieldDefinition field, string text)
    {
        var dateTime = DateTime.Parse(text, System.Globalization.CultureInfo.InvariantCulture);
        
        if (!string.IsNullOrWhiteSpace(field.Format))
        {
            return dateTime.ToString(ParseFormatString(field.Format));
        }
        
        return dateTime.ToString("yyyy-MM-dd HH:mm:ss");
    }

    private static string ParseDate(ExcelTemplateFieldDefinition field, string text)
    {
        var dateTime = DateTime.Parse(text, System.Globalization.CultureInfo.InvariantCulture);
        
        if (!string.IsNullOrWhiteSpace(field.Format))
        {
            return dateTime.ToString(ParseFormatString(field.Format));
        }
        
        return dateTime.ToString("yyyy-MM-dd");
    }

    private static string ParseTime(string text)
    {
        var dateTime = DateTime.Parse(text, System.Globalization.CultureInfo.InvariantCulture);
        return dateTime.ToString("HH:mm:ss");
    }

    private static string ParseFormatString(string format)
    {
        return format.ToLowerInvariant() switch
        {
            "yyyy-mm-dd" => "yyyy-MM-dd",
            "yyyy/mm/dd" => "yyyy/MM/dd",
            "yyyy-mm-dd hh:mm:ss" => "yyyy-MM-dd HH:mm:ss",
            "yyyy/mm/dd hh:mm:ss" => "yyyy/MM/dd HH:mm:ss",
            _ => format
        };
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
