using System.Text.Json;
using System.Text.Json.Serialization;
using ExcelImport.Core.Models.Template;
using ExcelImport.Core.Services;

namespace ExcelImport.WebApi.Services;

public sealed class ConfigService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase, allowIntegerValues: true) }
    };

    private readonly string _templatePath;

    public ConfigService()
    {
        _templatePath = new SharedPathResolver(AppContext.BaseDirectory).ResolveTemplateDirectory();
    }

    public ExcelTemplateDefinition LoadTemplate(string file)
    {
        if (string.IsNullOrWhiteSpace(file) || file.Contains(".."))
        {
            throw new InvalidOperationException("非法的 File。");
        }

        var path = Path.Combine(_templatePath, file);
        if (!File.Exists(path))
        {
            throw new FileNotFoundException($"未找到共享 Template 文件: {file}", path);
        }

        var json = File.ReadAllText(path);
        return JsonSerializer.Deserialize<ExcelTemplateDefinition>(json, JsonOptions)
               ?? throw new InvalidOperationException($"无法读取 Template: {file}");
    }

}
