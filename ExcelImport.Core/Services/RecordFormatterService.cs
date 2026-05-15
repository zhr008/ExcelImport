using ExcelImport.Core.Models;
using ExcelImport.Core.Models.Template;
using System.Text.RegularExpressions;

namespace ExcelImport.Core.Services;

public sealed class RecordFormatterService
{
    private const string NumericFilterPattern = @"[><=+\(\)@#\$%^&\*]";
    
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
                var text = CleanText(value.ToString());
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
            return field.Required ? null : null;
        }

        var text = CleanText(value.ToString());
        if (string.IsNullOrWhiteSpace(text))
        {
            return field.Required ? null : null;
        }


        object converted = field.Type.ToLowerInvariant() switch
        {
            "int" => int.Parse(FilterNumericText(text), System.Globalization.CultureInfo.InvariantCulture),
            "decimal" => ParseDecimal(field, text),
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

    private static string CleanText(string? text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return string.Empty;
        }

        return text.Replace("\r", "").Replace("\n", " ").Replace("\t", " ").Trim();
    }

    private static string FilterNumericText(string text)
    {
        return Regex.Replace(text, NumericFilterPattern, string.Empty);
    }

    private static decimal ParseDecimal(ExcelTemplateFieldDefinition field, string text)
    {
        var cleanedText = FilterNumericText(text);
        var value = decimal.Parse(cleanedText, System.Globalization.CultureInfo.InvariantCulture);
        
        var scale = ExtractDecimalScale(field);
        if (scale.HasValue)
        {
            value = decimal.Round(value, scale.Value);
        }
        
        return value;
    }

    private static int? ExtractDecimalScale(ExcelTemplateFieldDefinition field)
    {
        if (!string.IsNullOrWhiteSpace(field.Format))
        {
            var formatMatch = Regex.Match(field.Format, @"^(\d+),(\d+)$");
            if (formatMatch.Success)
            {
                return int.Parse(formatMatch.Groups[2].Value);
            }
        }
        
        return null;
    }

    private static string ParseDateTime(ExcelTemplateFieldDefinition field, string text)
    {
        var cleanedText = CleanText(text);
        
        if (DateTime.TryParse(cleanedText, System.Globalization.CultureInfo.InvariantCulture, 
            System.Globalization.DateTimeStyles.None, out var dateTime))
        {
            if (!string.IsNullOrWhiteSpace(field.Format))
            {
                return dateTime.ToString(ParseFormatString(field.Format));
            }
            
            return dateTime.ToString("yyyy-MM-dd HH:mm:ss");
        }
        
        throw new FormatException($"无法将 '{cleanedText}' 解析为 DateTime。");
    }

    private static string ParseDate(ExcelTemplateFieldDefinition field, string text)
    {
        var cleanedText = CleanText(text);
        
        if (DateTime.TryParse(cleanedText, System.Globalization.CultureInfo.InvariantCulture, 
            System.Globalization.DateTimeStyles.None, out var dateTime))
        {
            if (!string.IsNullOrWhiteSpace(field.Format))
            {
                return dateTime.ToString(ParseFormatString(field.Format));
            }
            
            return dateTime.ToString("yyyy-MM-dd");
        }
        
        throw new FormatException($"无法将 '{cleanedText}' 解析为 Date。");
    }

    private static string ParseTime(string text)
    {
        var cleanedText = CleanText(text);
        
        if (DateTime.TryParse(cleanedText, System.Globalization.CultureInfo.InvariantCulture, 
            System.Globalization.DateTimeStyles.None, out var dateTime))
        {
            return dateTime.ToString("HH:mm:ss");
        }
        
        throw new FormatException($"无法将 '{cleanedText}' 解析为 Time。");
    }

    private static string ParseFormatString(string format)
    {
        return format.ToLowerInvariant() switch
        {
            "yyyy-mm-dd" => "yyyy-MM-dd",
            "yyyy/mm/dd" => "yyyy/MM/dd",
            "yyyy-mm-dd hh:mm:ss" => "yyyy-MM-dd HH:mm:ss",
            "yyyy/mm/dd hh:mm:ss" => "yyyy/MM/dd HH:mm:ss",
            "yyyy年mm月dd日" => "yyyy年MM月dd日",
            "yyyy年mm月dd日 hh时mm分ss秒" => "yyyy年MM月dd日 HH时mm分ss秒",
            _ => format
        };
    }

    private static bool ParseBoolean(string text)
    {
        return text.Trim().ToUpperInvariant() switch
        {
            "TRUE" or "1" or "Y" or "YES" or "PASS" or "通过" => true,
            "FALSE" or "0" or "N" or "NO" or "FAIL" or "不通过" => false,
            _ => bool.Parse(text)
        };
    }
}
