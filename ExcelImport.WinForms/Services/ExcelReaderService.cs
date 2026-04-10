using System.Globalization;
using ClosedXML.Excel;
using ExcelImport.Core.Models.Template;

namespace ExcelImport.Services;

public sealed class ExcelReaderService
{
    public List<Dictionary<string, object?>> Read(string filePath, ExcelTemplateDefinition template)
    {
        using var workbook = new XLWorkbook(filePath);
        var worksheet = workbook.Worksheet(1);

        return template.ReadMode switch
        {
            ExcelTemplateReadMode.Row => ReadRows(worksheet, template),
            ExcelTemplateReadMode.Cell => [ReadCells(worksheet, template)],
            _ => throw new InvalidOperationException($"不支持的读取模式: {template.ReadMode}")
        };
    }

    private static List<Dictionary<string, object?>> ReadRows(IXLWorksheet worksheet, ExcelTemplateDefinition template)
    {
        if (template.StartRow is null or <= 0)
        {
            throw new InvalidOperationException("Row 模式必须设置有效的 StartRow。");
        }

        var records = new List<Dictionary<string, object?>>();
        var row = template.StartRow.Value;

        while (true)
        {
            var record = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            var hasAnyValue = false;

            foreach (var field in template.Fields)
            {
                if (string.IsNullOrWhiteSpace(field.Column))
                {
                    throw new InvalidOperationException($"字段 {field.Name} 未配置 Column。");
                }

                var cell = worksheet.Cell($"{field.Column}{row}");
                var rawValue = ReadRawValue(cell.Value);
                if (rawValue is not null)
                {
                    hasAnyValue = true;
                }

                record[field.Name] = rawValue;
            }

            if (!hasAnyValue)
            {
                break;
            }

            records.Add(record);
            row++;
        }

        return records;
    }

    private static Dictionary<string, object?> ReadCells(IXLWorksheet worksheet, ExcelTemplateDefinition template)
    {
        var record = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);

        foreach (var field in template.Fields)
        {
            if (string.IsNullOrWhiteSpace(field.Cell))
            {
                throw new InvalidOperationException($"字段 {field.Name} 未配置 Cell。");
            }

            var cell = worksheet.Cell(field.Cell);
            record[field.Name] = ReadRawValue(cell.Value);
        }

        return record;
    }

    private static object? ReadRawValue(XLCellValue rawValue)
    {
        if (rawValue.IsBlank)
        {
            return null;
        }

        var text = rawValue.ToString(CultureInfo.InvariantCulture).Trim();
        return string.IsNullOrWhiteSpace(text) ? null : text;
    }
}
