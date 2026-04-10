namespace ExcelImport.Core.Services;

public sealed class SharedPathResolver
{
    private readonly string _baseDirectory;

    public SharedPathResolver(string baseDirectory)
    {
        _baseDirectory = Path.GetFullPath(baseDirectory);
    }

    public string BaseDirectory => _baseDirectory;

    public string ResolveTemplateDirectory()
    {
        var deployedTemplatePath = Path.GetFullPath(Path.Combine(_baseDirectory, "..", "Template"));
        if (Directory.Exists(deployedTemplatePath))
        {
            return deployedTemplatePath;
        }

        var projectRootPath = ResolveProjectRootPath();
        if (projectRootPath is not null)
        {
            return Path.Combine(projectRootPath, "Template");
        }

        return Path.Combine(_baseDirectory, "Template");
    }

    public string? ResolveProjectRootPath()
    {
        var current = new DirectoryInfo(_baseDirectory);
        while (current is not null)
        {
            if (Directory.Exists(Path.Combine(current.FullName, "ExcelImport.WinForms")) &&
                Directory.Exists(Path.Combine(current.FullName, "ExcelImport.WebApi")) &&
                Directory.Exists(Path.Combine(current.FullName, "Template")))
            {
                return current.FullName;
            }

            current = current.Parent;
        }

        return null;
    }
}
