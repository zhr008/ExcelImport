using ExcelImport.Core.Models;
using log4net;
using log4net.Config;

namespace ExcelImport.Services;

public enum LogLevel
{
    Debug,
    Info,
    Warning,
    Error
}

public sealed class LoggingService
{
    private readonly string _baseDirectory;
    private readonly Func<AppSettings> _settingsAccessor;
    private readonly ILog _logger;

    public LoggingService(string baseDirectory, Func<AppSettings> settingsAccessor)
    {
        _baseDirectory = baseDirectory;
        _settingsAccessor = settingsAccessor;
        var repository = LogManager.CreateRepository(Guid.NewGuid().ToString("N"));
        XmlConfigurator.Configure(repository, new FileInfo(Path.Combine(baseDirectory, "log4net.config")));
        _logger = LogManager.GetLogger(repository.Name, "ExcelImport");
    }

    public event Action<string>? LogWritten;

    public void Info(string message) => Write(LogLevel.Info, message, null);

    public void Warning(string message) => Write(LogLevel.Warning, message, null);

    public void Error(string message, Exception? exception = null)
    {
        var finalMessage = exception is null ? message : $"{message}{Environment.NewLine}{exception}";
        Write(LogLevel.Error, finalMessage, exception);
    }

    public string ReadRecentLines(int maxLines = 200)
    {
        var path = GetLogFilePath();
        if (!File.Exists(path))
        {
            return string.Empty;
        }

        var lines = File.ReadAllLines(path);
        return string.Join(Environment.NewLine, lines.TakeLast(maxLines));
    }

    private void Write(LogLevel level, string message, Exception? exception)
    {
        var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}";

        switch (level)
        {
            case LogLevel.Debug:
                _logger.Debug(message, exception);
                break;
            case LogLevel.Info:
                _logger.Info(message, exception);
                break;
            case LogLevel.Warning:
                _logger.Warn(message, exception);
                break;
            case LogLevel.Error:
                _logger.Error(message, exception);
                break;
        }

        LogWritten?.Invoke(line);
    }

    private string GetLogFilePath()
    {
        var settings = _settingsAccessor();
        var relativeDirectory = string.IsNullOrWhiteSpace(settings.LogDirectory) ? "Logs" : settings.LogDirectory;
        return Path.Combine(_baseDirectory, relativeDirectory, $"{DateTime.Today:yyyy-MM-dd}.log");
    }
}
