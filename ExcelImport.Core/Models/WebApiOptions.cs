namespace ExcelImport.Core.Models;

public sealed class WebApiOptions
{
    public bool Enabled { get; set; }

    public string BaseUrl { get; set; } = string.Empty;

    public string Endpoint { get; set; } = "/api/excel-import";
}
