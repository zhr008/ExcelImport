namespace ExcelImport.Core.Models;

public sealed class ImportResult
{
    public int FilesScanned { get; set; }

    public int FilesSucceeded { get; set; }

    public int FilesFailed { get; set; }

    public int RecordsImported { get; set; }
}
