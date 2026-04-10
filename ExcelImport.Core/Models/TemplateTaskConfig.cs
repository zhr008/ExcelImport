namespace ExcelImport.Core.Models;

public sealed class TemplateTaskConfig
{
    public string Name { get; set; } = string.Empty;

    public bool Enabled { get; set; }

    public string WatchPath { get; set; } = string.Empty;

    public bool IncludeSubdirectories { get; set; } = true;

    public int IntervalMinutes { get; set; } = 5;

    public string TemplateFile { get; set; } = string.Empty;

    public string TargetTable { get; set; } = string.Empty;

    public string FilePattern { get; set; } = "*.xlsx";
}
