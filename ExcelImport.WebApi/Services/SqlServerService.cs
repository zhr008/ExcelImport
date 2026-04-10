using System.Data;
using System.Text.RegularExpressions;
using ExcelImport.Core.Models.Template;
using Microsoft.Data.SqlClient;

namespace ExcelImport.WebApi.Services;

public sealed class SqlServerService
{
    private static readonly Regex TableNamePattern = new(@"^[A-Za-z0-9_\.\[\]]+$", RegexOptions.Compiled);

    public async Task<int> InsertAsync(string connectionString, string tableName, ExcelTemplateDefinition template, IReadOnlyList<Dictionary<string, object?>> records, CancellationToken cancellationToken)
    {
        if (records.Count == 0)
        {
            return 0;
        }

        if (string.IsNullOrWhiteSpace(connectionString))
        {
            throw new InvalidOperationException("未配置 Database:ConnectionString。");
        }

        if (!IsSafeTableName(tableName))
        {
            throw new InvalidOperationException($"非法的目标表名: {tableName}");
        }

        var keyFields = template.Fields.Where(field => field.IsKey).Select(field => field.Name).ToArray();

        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancellationToken);

        var inserted = 0;
        foreach (var record in records)
        {
            await using var command = connection.CreateCommand();
            command.CommandType = CommandType.Text;
            command.CommandText = BuildInsertCommand(tableName, record.Keys, keyFields);

            foreach (var pair in record)
            {
                command.Parameters.AddWithValue($"@{pair.Key}", pair.Value ?? DBNull.Value);
            }

            inserted += await command.ExecuteNonQueryAsync(cancellationToken);
        }

        return inserted;
    }

    private static bool IsSafeTableName(string tableName)
    {
        return !string.IsNullOrWhiteSpace(tableName) && TableNamePattern.IsMatch(tableName);
    }

    private static string BuildInsertCommand(string tableName, IEnumerable<string> columns, IReadOnlyCollection<string> keyFields)
    {
        var columnList = columns.ToArray();
        var safeColumns = columnList.Select(column => $"[{column}]").ToArray();
        var parameters = columnList.Select(column => $"@{column}").ToArray();

        if (keyFields.Count == 0)
        {
            return $"INSERT INTO {tableName} ({string.Join(", ", safeColumns)}) VALUES ({string.Join(", ", parameters)})";
        }

        var keyPredicates = keyFields
            .Select(field => $"([{field}] = @{field} OR ([{field}] IS NULL AND @{field} IS NULL))")
            .ToArray();

        return $"INSERT INTO {tableName} ({string.Join(", ", safeColumns)}) SELECT {string.Join(", ", parameters)} WHERE NOT EXISTS (SELECT 1 FROM {tableName} WHERE {string.Join(" AND ", keyPredicates)})";
    }
}
