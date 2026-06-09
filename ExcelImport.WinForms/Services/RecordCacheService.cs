using Microsoft.Data.Sqlite;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace ExcelImport.Services;

/// <summary>
/// 记录缓存服务，用于本地缓存已插入的记录，避免重复插入
/// </summary>
public class RecordCacheService
{
    private readonly string _dbPath;
    private readonly LoggingService _loggingService;
    private readonly SemaphoreSlim _semaphore = new(1, 1);

    public RecordCacheService(LoggingService loggingService)
    {
        _loggingService = loggingService;
        _dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "cache", "imported_records.db");
        InitializeDatabase();
    }

    /// <summary>
    /// 初始化数据库表
    /// </summary>
    private void InitializeDatabase()
    {
        var directory = Path.GetDirectoryName(_dbPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        using var connection = new SqliteConnection($"Data Source={_dbPath}");
        connection.Open();

        var command = connection.CreateCommand();
        command.CommandText = @"
            CREATE TABLE IF NOT EXISTS ImportedRecords (
                TemplateName TEXT NOT NULL,
                RecordHash TEXT NOT NULL,
                ImportedAt TEXT NOT NULL,
                PRIMARY KEY (TemplateName, RecordHash)
            );
            CREATE INDEX IF NOT EXISTS IX_ImportedRecords_ImportedAt ON ImportedRecords(ImportedAt);
        ";
        command.ExecuteNonQuery();
    }

    /// <summary>
    /// 计算记录的哈希值
    /// </summary>
    public string ComputeRecordHash(string templateName, Dictionary<string, object?> record)
    {
        // 对键值对进行排序，确保相同内容的记录生成相同的哈希
        var sortedRecord = record.OrderBy(kv => kv.Key).ToDictionary(kv => kv.Key, kv => kv.Value);
        var json = JsonSerializer.Serialize(sortedRecord);
        
        using var sha256 = SHA256.Create();
        var bytes = Encoding.UTF8.GetBytes($"{templateName}:{json}");
        var hash = sha256.ComputeHash(bytes);
        return Convert.ToHexString(hash).ToLowerInvariant();
    }

    /// <summary>
    /// 检查记录是否已存在
    /// </summary>
    public async Task<bool> ExistsAsync(string templateName, Dictionary<string, object?> record)
    {
        await _semaphore.WaitAsync();
        try
        {
            var hash = ComputeRecordHash(templateName, record);

            using var connection = new SqliteConnection($"Data Source={_dbPath}");
            await connection.OpenAsync();

            var command = connection.CreateCommand();
            command.CommandText = @"
                SELECT COUNT(1) 
                FROM ImportedRecords 
                WHERE TemplateName = @TemplateName AND RecordHash = @RecordHash;
            ";
            command.Parameters.AddWithValue("@TemplateName", templateName);
            command.Parameters.AddWithValue("@RecordHash", hash);

            var result = await command.ExecuteScalarAsync();
            return Convert.ToInt32(result) > 0;
        }
        catch (Exception ex)
        {
            _loggingService.Error($"检查记录缓存失败: {ex.Message}", ex);
            return false; // 出错时返回false，允许插入
        }
        finally
        {
            _semaphore.Release();
        }
    }

    /// <summary>
    /// 批量检查记录是否存在
    /// </summary>
    public async Task<HashSet<string>> BatchExistsAsync(string templateName, List<Dictionary<string, object?>> records)
    {
        await _semaphore.WaitAsync();
        try
        {
            var hashes = records.Select(r => ComputeRecordHash(templateName, r)).ToHashSet();

            using var connection = new SqliteConnection($"Data Source={_dbPath}");
            await connection.OpenAsync();

            var command = connection.CreateCommand();
            command.CommandText = @"
                SELECT RecordHash 
                FROM ImportedRecords 
                WHERE TemplateName = @TemplateName AND RecordHash IN ({0});
            ";
            command.Parameters.AddWithValue("@TemplateName", templateName);

            // 构建IN子句
            var paramNames = hashes.Select((_, i) => $"@hash{i}").ToArray();
            command.CommandText = string.Format(command.CommandText, string.Join(", ", paramNames));

            for (int i = 0; i < hashes.Count; i++)
            {
                command.Parameters.AddWithValue(paramNames[i], hashes.ElementAt(i));
            }

            var existingHashes = new HashSet<string>();
            using var reader = await command.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                existingHashes.Add(reader.GetString(0));
            }

            return existingHashes;
        }
        catch (Exception ex)
        {
            _loggingService.Error($"批量检查记录缓存失败: {ex.Message}", ex);
            return new HashSet<string>(); // 出错时返回空集合，允许所有记录插入
        }
        finally
        {
            _semaphore.Release();
        }
    }

    /// <summary>
    /// 添加记录到缓存
    /// </summary>
    public async Task AddAsync(string templateName, Dictionary<string, object?> record)
    {
        await _semaphore.WaitAsync();
        try
        {
            var hash = ComputeRecordHash(templateName, record);

            using var connection = new SqliteConnection($"Data Source={_dbPath}");
            await connection.OpenAsync();

            var command = connection.CreateCommand();
            command.CommandText = @"
                INSERT OR IGNORE INTO ImportedRecords (TemplateName, RecordHash, ImportedAt)
                VALUES (@TemplateName, @RecordHash, @ImportedAt);
            ";
            command.Parameters.AddWithValue("@TemplateName", templateName);
            command.Parameters.AddWithValue("@RecordHash", hash);
            command.Parameters.AddWithValue("@ImportedAt", DateTime.UtcNow.ToString("o"));

            await command.ExecuteNonQueryAsync();
        }
        catch (Exception ex)
        {
            _loggingService.Error($"添加记录缓存失败: {ex.Message}", ex);
        }
        finally
        {
            _semaphore.Release();
        }
    }

    /// <summary>
    /// 批量添加记录到缓存
    /// </summary>
    public async Task BatchAddAsync(string templateName, List<Dictionary<string, object?>> records)
    {
        if (records.Count == 0) return;

        await _semaphore.WaitAsync();
        try
        {
            using var connection = new SqliteConnection($"Data Source={_dbPath}");
            await connection.OpenAsync();

            using var transaction = connection.BeginTransaction();

            foreach (var record in records)
            {
                var hash = ComputeRecordHash(templateName, record);

                var command = connection.CreateCommand();
                command.Transaction = transaction;
                command.CommandText = @"
                    INSERT OR IGNORE INTO ImportedRecords (TemplateName, RecordHash, ImportedAt)
                    VALUES (@TemplateName, @RecordHash, @ImportedAt);
                ";
                command.Parameters.AddWithValue("@TemplateName", templateName);
                command.Parameters.AddWithValue("@RecordHash", hash);
                command.Parameters.AddWithValue("@ImportedAt", DateTime.UtcNow.ToString("o"));

                await command.ExecuteNonQueryAsync();
            }

            await transaction.CommitAsync();
        }
        catch (Exception ex)
        {
            _loggingService.Error($"批量添加记录缓存失败: {ex.Message}", ex);
        }
        finally
        {
            _semaphore.Release();
        }
    }

    /// <summary>
    /// 清理过期的缓存记录
    /// </summary>
    public async Task CleanupExpiredRecordsAsync(int daysToKeep = 30)
    {
        await _semaphore.WaitAsync();
        try
        {
            var cutoffDate = DateTime.UtcNow.AddDays(-daysToKeep).ToString("o");

            using var connection = new SqliteConnection($"Data Source={_dbPath}");
            await connection.OpenAsync();

            var command = connection.CreateCommand();
            command.CommandText = @"
                DELETE FROM ImportedRecords 
                WHERE ImportedAt < @CutoffDate;
            ";
            command.Parameters.AddWithValue("@CutoffDate", cutoffDate);

            var deleted = await command.ExecuteNonQueryAsync();
            _loggingService.Info($"已清理 {deleted} 条过期缓存记录（保留 {daysToKeep} 天）");
        }
        catch (Exception ex)
        {
            _loggingService.Error($"清理过期缓存记录失败: {ex.Message}", ex);
        }
        finally
        {
            _semaphore.Release();
        }
    }

    /// <summary>
    /// 获取缓存统计信息
    /// </summary>
    public async Task<(int TotalRecords, int OldestDays)> GetStatisticsAsync()
    {
        await _semaphore.WaitAsync();
        try
        {
            using var connection = new SqliteConnection($"Data Source={_dbPath}");
            await connection.OpenAsync();

            var command = connection.CreateCommand();
            command.CommandText = @"
                SELECT 
                    COUNT(*) as Total,
                    MIN(ImportedAt) as Oldest
                FROM ImportedRecords;
            ";

            using var reader = await command.ExecuteReaderAsync();
            if (await reader.ReadAsync())
            {
                var total = reader.GetInt32(0);
                var oldest = reader.IsDBNull(1) ? (DateTime?)null : DateTime.Parse(reader.GetString(1));
                var oldestDays = oldest.HasValue ? (int)(DateTime.UtcNow - oldest.Value).TotalDays : 0;
                return (total, oldestDays);
            }

            return (0, 0);
        }
        catch (Exception ex)
        {
            _loggingService.Error($"获取缓存统计信息失败: {ex.Message}", ex);
            return (0, 0);
        }
        finally
        {
            _semaphore.Release();
        }
    }
}