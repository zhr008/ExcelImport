using System.Runtime.InteropServices;
using System.Threading;

namespace ExcelImport;

static class Program
{
    private const string StartupMinimizedArgument = "--startup-minimized";
    private const string SingleInstanceMutexName = "ExcelImport.WinForms.SingleInstance";
    private const string ActivationMessageName = "ExcelImport.WinForms.ActivateExistingInstance";
    private const uint HwndBroadcast = 0xffff;

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern uint RegisterWindowMessage(string lpString);

    [DllImport("user32.dll", SetLastError = false)]
    private static extern nint PostMessage(nint hWnd, uint msg, nint wParam, nint lParam);

    /// <summary>
    ///  The main entry point for the application.
    /// </summary>
    [STAThread]
    static void Main(string[] args)
    {
        using var mutex = new Mutex(true, SingleInstanceMutexName, out var createdNew);
        var activationMessageId = RegisterWindowMessage(ActivationMessageName);
        if (!createdNew)
        {
            PostMessage((nint)HwndBroadcast, activationMessageId, nint.Zero, nint.Zero);
            return;
        }

        ApplicationConfiguration.Initialize();
        var startMinimizedToTray = args.Any(arg => string.Equals(arg, StartupMinimizedArgument, StringComparison.OrdinalIgnoreCase));
        Application.Run(new Form1(startMinimizedToTray, activationMessageId));
    }
}
