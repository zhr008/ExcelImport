namespace ExcelImport.Core.Models;

public sealed class DatabaseOptions
{
    public bool Enabled { get; set; }

    public string ConnectionString { get; set; } = string.Empty;
}
