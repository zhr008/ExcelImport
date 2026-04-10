using Microsoft.Win32;

namespace ExcelImport.Services;

public sealed class StartupService
{
    private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string AppName = "ExcelImport";
    private const string StartupMinimizedArgument = "--startup-minimized";

    public bool IsEnabled()
    {
        using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, false);
        return key?.GetValue(AppName) is string;
    }

    public void SetEnabled(bool enabled)
    {
        using var key = Registry.CurrentUser.CreateSubKey(RunKeyPath);

        if (enabled)
        {
            var exePath = Application.ExecutablePath;
            key.SetValue(AppName, $"\"{exePath}\" {StartupMinimizedArgument}");
            return;
        }

        key.DeleteValue(AppName, false);
    }
}
