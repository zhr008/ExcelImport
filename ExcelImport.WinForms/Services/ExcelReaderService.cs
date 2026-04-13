using System.Globalization;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using ExcelImport.Core.Models.Template;

namespace ExcelImport.Services;

public sealed class ExcelReaderService
{
    private static readonly Regex ColumnPattern = new("^[A-Za-z]+$", RegexOptions.Compiled);
    private static readonly Regex CellPattern = new("^[A-Za-z]+[1-9][0-9]*$", RegexOptions.Compiled);
    private static readonly Regex StringLiteralPattern = new("^'([^']*)'$", RegexOptions.Compiled);
    private static readonly Regex NumericPattern = new(@"^-?\d+(\.\d+)?$", RegexOptions.Compiled);

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

                var rawValue = EvaluateExpression(worksheet, field.Column, field.Name, row);
                // 检查是否有有效值（非null且非空字符串）
                if (rawValue is not null)
                {
                    if (rawValue is string strValue && string.IsNullOrWhiteSpace(strValue))
                    {
                        // 空字符串视为无效值
                        record[field.Name] = null;
                    }
                    else
                    {
                        hasAnyValue = true;
                        record[field.Name] = rawValue;
                    }
                }
                else
                {
                    record[field.Name] = null;
                }
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

            record[field.Name] = EvaluateExpression(worksheet, field.Cell, field.Name, null);
        }

        return record;
    }

    private static object? EvaluateExpression(IXLWorksheet worksheet, string expression, string fieldName, int? currentRow)
    {
        try
        {
            // 处理简单的拼接和运算表达式
            var tokens = TokenizeExpression(expression);
            if (tokens.Count == 0)
            {
                return null;
            }

            // 计算表达式
            var result = CalculateExpression(worksheet, tokens, currentRow);
            return result;
        }
        catch (Exception)
        {
            // 运算出错时返回null
            return null;
        }
    }

    private static List<Token> TokenizeExpression(string expression)
    {
        var tokens = new List<Token>();
        var current = string.Empty;
        var inString = false;

        for (int i = 0; i < expression.Length; i++)
        {
            var c = expression[i];

            if (c == '\'')
            {
                inString = !inString;
                current += c;
            }
            else if (!inString && "&+-*/".Contains(c))
            {
                if (!string.IsNullOrWhiteSpace(current))
                {
                    tokens.Add(CreateToken(current.Trim()));
                    current = string.Empty;
                }
                tokens.Add(new Token { Type = TokenType.Operator, Value = c.ToString() });
            }
            else
            {
                current += c;
            }
        }

        if (!string.IsNullOrWhiteSpace(current))
        {
            tokens.Add(CreateToken(current.Trim()));
        }

        return tokens;
    }

    private static Token CreateToken(string value)
    {
        if (StringLiteralPattern.IsMatch(value))
        {
            var match = StringLiteralPattern.Match(value);
            return new Token { Type = TokenType.StringLiteral, Value = match.Groups[1].Value };
        }
        else if (NumericPattern.IsMatch(value))
        {
            return new Token { Type = TokenType.Numeric, Value = value };
        }
        else
        {
            return new Token { Type = TokenType.Reference, Value = value };
        }
    }

    private static object? CalculateExpression(IXLWorksheet worksheet, List<Token> tokens, int? currentRow)
    {
        if (tokens.Count == 0)
        {
            return null;
        }

        // 处理只有一个token的情况
        if (tokens.Count == 1)
        {
            return GetTokenValue(worksheet, tokens[0], currentRow);
        }

        // 处理多个token的表达式（支持&+-*/）
        var result = GetTokenValue(worksheet, tokens[0], currentRow);
        
        for (int i = 1; i < tokens.Count; i += 2)
        {
            if (i + 1 >= tokens.Count)
            {
                break;
            }

            var op = tokens[i].Value;
            var nextValue = GetTokenValue(worksheet, tokens[i + 1], currentRow);
            
            // 对于拼接操作，null值视为空字符串
            if (op == "&")
            {
                var leftStr = result?.ToString() ?? "";
                var rightStr = nextValue?.ToString() ?? "";
                result = leftStr + rightStr;
            }
            // 对于其他操作，需要两个值都不为null
            else
            {
                if (result == null || nextValue == null)
                {
                    return null;
                }
                result = ApplyOperation(result, nextValue, op);
                if (result == null)
                {
                    return null;
                }
            }
        }

        return result;
    }

    private static object? GetTokenValue(IXLWorksheet worksheet, Token token, int? currentRow)
    {
        switch (token.Type)
        {
            case TokenType.StringLiteral:
                return token.Value;
            case TokenType.Numeric:
                if (decimal.TryParse(token.Value, CultureInfo.InvariantCulture, out var num))
                {
                    return num;
                }
                return null;
            case TokenType.Reference:
                if (currentRow.HasValue && ColumnPattern.IsMatch(token.Value))
                {
                    // Row模式：列引用
                    return ReadRawValue(worksheet.Cell($"{token.Value}{currentRow.Value}").Value);
                }
                else if (CellPattern.IsMatch(token.Value))
                {
                    // Cell模式：单元格引用
                    return ReadRawValue(worksheet.Cell(token.Value).Value);
                }
                return null;
            default:
                return null;
        }
    }

    private static object? ApplyOperation(object left, object right, string op)
    {
        try
        {
            switch (op)
            {
                case "&":
                    // 字符串拼接
                    return left.ToString() + right.ToString();
                case "+":
                    // 加法
                    if (left is decimal leftDec1 && right is decimal rightDec1)
                    {
                        return leftDec1 + rightDec1;
                    }
                    else if (double.TryParse(left.ToString(), out var leftDouble1) && double.TryParse(right.ToString(), out var rightDouble1))
                    {
                        return leftDouble1 + rightDouble1;
                    }
                    return null;
                case "-":
                    // 减法
                    if (left is decimal leftDec2 && right is decimal rightDec2)
                    {
                        return leftDec2 - rightDec2;
                    }
                    else if (double.TryParse(left.ToString(), out var leftDouble2) && double.TryParse(right.ToString(), out var rightDouble2))
                    {
                        return leftDouble2 - rightDouble2;
                    }
                    return null;
                case "*":
                    // 乘法
                    if (left is decimal leftDec3 && right is decimal rightDec3)
                    {
                        return leftDec3 * rightDec3;
                    }
                    else if (double.TryParse(left.ToString(), out var leftDouble3) && double.TryParse(right.ToString(), out var rightDouble3))
                    {
                        return leftDouble3 * rightDouble3;
                    }
                    return null;
                case "/":
                    // 除法
                    if (left is decimal leftDec4 && right is decimal rightDec4 && rightDec4 != 0)
                    {
                        return leftDec4 / rightDec4;
                    }
                    else if (double.TryParse(left.ToString(), out var leftDouble4) && double.TryParse(right.ToString(), out var rightDouble4) && rightDouble4 != 0)
                    {
                        return leftDouble4 / rightDouble4;
                    }
                    return null;
                default:
                    return null;
            }
        }
        catch (Exception)
        {
            return null;
        }
    }

    private static object? ReadRawValue(XLCellValue rawValue)
    {
        if (rawValue.IsBlank)
        {
            return null;
        }

        // 处理日期时间类型
        if (rawValue.IsDateTime)
        {
            try
            {
                var dateTime = rawValue.GetDateTime();
                
                // 检查是否是时间值（Excel中纯时间值的日期部分通常为1900-01-00或1900-01-01）
                // 注意：1900-01-00是Excel的特殊表示，实际会转换为1899-12-31
                var dateOnly = dateTime.Date;
                if (dateOnly == new DateTime(1900, 1, 1) || dateOnly == new DateTime(1899, 12, 31))
                {
                    // 仅时间值，确保小时在0-23范围内
                    var safeHour = Math.Min(23, dateTime.Hour);
                    var timeValue = new DateTime(1900, 1, 1, safeHour, dateTime.Minute, dateTime.Second);
                    return timeValue.ToString("HH:mm:ss");
                }
                // 检查是否是完整的日期时间值（时间部分不为00:00:00）
                else if (dateTime.TimeOfDay != TimeSpan.Zero)
                {
                    // 完整日期时间值，返回完整的日期时间字符串
                    return dateTime.ToString("yyyy-MM-dd HH:mm:ss");
                }
                // 纯日期值，格式化为yyyy-MM-dd
                else
                {
                    return dateTime.ToString("yyyy-MM-dd");
                }
            }
            catch (Exception)
            {
                // 日期解析失败，返回null
                return null;
            }
        }

        var text = rawValue.ToString(CultureInfo.InvariantCulture).Trim();
        return string.IsNullOrWhiteSpace(text) ? null : text;
    }

    private class Token
    {
        public TokenType Type { get; set; }
        public string Value { get; set; } = string.Empty;
    }

    private enum TokenType
    {
        Reference,
        StringLiteral,
        Numeric,
        Operator
    }
}
