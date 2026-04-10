namespace ExcelImport.Core.Models;

public sealed class AppSettings
{
    public DatabaseOptions Database { get; set; } = new();

    public WebApiOptions WebApi { get; set; } = new();

    public bool StartWithWindows { get; set; }

    public string LogDirectory { get; set; } = "Logs";

    public List<TemplateTaskConfig> Templates { get; set; } = [];
}
