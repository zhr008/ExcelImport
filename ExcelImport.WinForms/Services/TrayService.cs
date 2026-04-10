namespace ExcelImport.Services;

public sealed class TrayService : IDisposable
{
    private readonly NotifyIcon _notifyIcon;

    public TrayService(Icon icon, Action showAction, Action runNowAction, Action exitAction)
    {
        var menu = new ContextMenuStrip();
        menu.Items.Add("显示主界面", null, (_, _) => showAction());
        menu.Items.Add("立即执行", null, (_, _) => runNowAction());
        menu.Items.Add("退出", null, (_, _) => exitAction());

        _notifyIcon = new NotifyIcon
        {
            Icon = icon,
            Visible = true,
            Text = "Excel Import",
            ContextMenuStrip = menu
        };

        _notifyIcon.DoubleClick += (_, _) => showAction();
    }

    public void Dispose()
    {
        _notifyIcon.Visible = false;
        _notifyIcon.Dispose();
    }
}
