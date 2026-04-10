using System.Text.Json;
using System.Text.Json.Serialization;
using ExcelImport.Core.Models;
using ExcelImport.Core.Models.Template;
using ExcelImport.Core.Services;

namespace ExcelImport.Services;

public sealed class ConfigService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase, allowIntegerValues: true) }
    };

    private readonly string _baseDirectory;
    private readonly string _templatePath;

    public ConfigService(string baseDirectory)
    {
        _baseDirectory = baseDirectory;
        _templatePath = new SharedPathResolver(baseDirectory).ResolveTemplateDirectory();
    }

    public string AppSettingsPath => Path.Combine(_baseDirectory, "appsettings.json");

    public string TemplatePath => _templatePath;

    public string BaseDirectory => _baseDirectory;

    public AppSettings LoadAppSettings()
    {
        EnsureDefaultFiles();

        var json = File.ReadAllText(AppSettingsPath);
        return JsonSerializer.Deserialize<AppSettings>(json, JsonOptions) ?? new AppSettings();
    }

    public void SaveAppSettings(AppSettings settings)
    {
        Directory.CreateDirectory(_baseDirectory);
        var json = JsonSerializer.Serialize(settings, JsonOptions);
        File.WriteAllText(AppSettingsPath, json);
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

    public void EnsureDefaultFiles()
    {
        Directory.CreateDirectory(_baseDirectory);
        Directory.CreateDirectory(Path.Combine(_baseDirectory, "Logs"));

        if (!File.Exists(AppSettingsPath))
        {
            throw new FileNotFoundException($"未找到配置文件: {AppSettingsPath}", AppSettingsPath);
        }
    }
}
