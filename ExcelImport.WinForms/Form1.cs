using System.ComponentModel;
using ExcelImport.Core.Models;
using ExcelImport.Core.Services;
using ExcelImport.Services;

namespace ExcelImport;

public partial class Form1 : Form
{
    private readonly string _baseDirectory = AppContext.BaseDirectory;
    private readonly bool _startMinimizedToTray;
    private readonly uint _activationMessageId;
    private readonly ConfigService _configService;
    private readonly StartupService _startupService;
    private readonly ExcelReaderService _excelReaderService;
    private readonly RecordFormatterService _recordFormatterService;
    private readonly SqlServerService _sqlServerService;
    private readonly WebApiService _webApiService;
    private readonly LoggingService _loggingService;
    private readonly ImportService _importService;
    private readonly SchedulerService _schedulerService;
    private readonly TrayService _trayService;
    private AppSettings _settings = new();
    private BindingList<TemplateTaskConfig> _templateBindingList = [];
    private bool _allowClose;

    public Form1(bool startMinimizedToTray = false, uint activationMessageId = 0)
    {
        _startMinimizedToTray = startMinimizedToTray;
        _activationMessageId = activationMessageId;
        InitializeComponent();

        _configService = new ConfigService(_baseDirectory);
        _startupService = new StartupService();
        _excelReaderService = new ExcelReaderService();
        _recordFormatterService = new RecordFormatterService();
        _webApiService = new WebApiService();
        _loggingService = new LoggingService(_baseDirectory, () => _settings);
        _sqlServerService = new SqlServerService(_loggingService);
        _loggingService.LogWritten += AppendLogLine;
        _importService = new ImportService(_configService, _excelReaderService, _recordFormatterService, _sqlServerService, _webApiService, _loggingService);
        _schedulerService = new SchedulerService(_importService, _loggingService);
        _trayService = new TrayService(Icon ?? SystemIcons.Application, ShowFromTray, RunNowFromTray, ExitApplication);

        Load += Form1_Load;
        Shown += Form1_Shown;
        Resize += Form1_Resize;
        FormClosing += Form1_FormClosing;
    }

    private void Form1_Load(object? sender, EventArgs e)
    {
        _settings = _configService.LoadAppSettings();
        _loggingService.Info($"配置文件路径: {_configService.AppSettingsPath}");
        _loggingService.Info($"程序目录: {_configService.BaseDirectory}");
        _loggingService.Info($"Template 目录: {_configService.TemplatePath}");
        _settings.StartWithWindows = _startupService.IsEnabled();
        BindSettings();
        _loggingService.Info($"已加载 {_settings.Templates.Count} 个模板，启用 {_settings.Templates.Count(t => t.Enabled)} 个。");
        _schedulerService.Start(_settings);
        RefreshLog();
    }

    private void BindSettings()
    {
        chkStartWithWindows.Checked = _settings.StartWithWindows;

        _templateBindingList = new BindingList<TemplateTaskConfig>(_settings.Templates);
        dgvTemplates.AutoGenerateColumns = false;
        dgvTemplates.DataSource = _templateBindingList;
    }

    private void SaveSettings()
    {
        dgvTemplates.EndEdit();
        _settings.StartWithWindows = chkStartWithWindows.Checked;
        _settings.Templates = _templateBindingList.ToList();

        _configService.SaveAppSettings(_settings);
        _loggingService.Info($"配置已写入: {_configService.AppSettingsPath}");
        _startupService.SetEnabled(_settings.StartWithWindows);
        _schedulerService.Start(_settings);
        _loggingService.Info("配置已保存。");
        RefreshLog();
    }

    private void btnSave_Click(object sender, EventArgs e)
    {
        try
        {
            SaveSettings();
            MessageBox.Show("配置已保存。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            _loggingService.Error("保存配置失败。", ex);
            RefreshLog();
            MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void btnAdd_Click(object sender, EventArgs e)
    {
        _templateBindingList.Add(new TemplateTaskConfig
        {
            Name = "NewTemplate",
            Enabled = false,
            WatchPath = string.Empty,
            IncludeSubdirectories = true,
            IntervalMinutes = 5,
            TemplateFile = "row-template.json",
            TargetTable = "dbo.ExcelImports",
            FilePattern = "*.xlsx"
        });
    }

    private void btnDelete_Click(object sender, EventArgs e)
    {
        if (dgvTemplates.CurrentRow?.DataBoundItem is TemplateTaskConfig config)
        {
            _templateBindingList.Remove(config);
        }
    }

    private void btnSelectFolder_Click(object sender, EventArgs e)
    {
        if (!TryGetCurrentTemplate(out var template))
        {
            MessageBox.Show("请先选择一条模板记录。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        using var dialog = new FolderBrowserDialog();
        if (Directory.Exists(template.WatchPath))
        {
            dialog.SelectedPath = template.WatchPath;
        }

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            template.WatchPath = dialog.SelectedPath;
            dgvTemplates.Refresh();
        }
    }

    private void btnSelect_Click(object sender, EventArgs e)
    {
        if (!TryGetCurrentTemplate(out var template))
        {
            MessageBox.Show("请先选择一条模板记录。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        using var dialog = new OpenFileDialog
        {
            Filter = "JSON Files (*.json)|*.json",
            Title = "选择 Template 文件"
        };

        var templateDirectory = _configService.TemplatePath;
        if (Directory.Exists(templateDirectory))
        {
            dialog.InitialDirectory = templateDirectory;
        }

        var currentTemplateFullPath = GetSharedTemplateFullPath(template.TemplateFile);
        if (File.Exists(currentTemplateFullPath))
        {
            dialog.FileName = currentTemplateFullPath;
        }

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            template.TemplateFile = GetTemplatePath(dialog.FileName);
            dgvTemplates.Refresh();
        }
    }

    private async void btnRunNow_Click(object sender, EventArgs e)
    {
        await RunSelectedTemplateAsync();
    }

    private async Task RunSelectedTemplateAsync()
    {
        if (dgvTemplates.CurrentRow?.DataBoundItem is not TemplateTaskConfig template)
        {
            return;
        }

        try
        {
            btnRunNow.Enabled = false;
            dgvTemplates.EndEdit();
            _settings.Templates = _templateBindingList.ToList();

            var result = await _schedulerService.RunOnceAsync(_settings, template);
            RefreshLog();
            MessageBox.Show($"执行完成。文件: {result.FilesScanned}，成功: {result.FilesSucceeded}，失败: {result.FilesFailed}，记录: {result.RecordsImported}", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            _loggingService.Error("手动执行失败。", ex);
            RefreshLog();
            MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            btnRunNow.Enabled = true;
        }
    }

    private void btnRefreshLog_Click(object sender, EventArgs e)
    {
        RefreshLog();
    }

    private void btnClearLog_Click(object sender, EventArgs e)
    {
        txtLogs.Clear();
    }

    private void AppendLogLine(string line)
    {
        if (IsDisposed || Disposing || txtLogs.IsDisposed)
        {
            return;
        }

        if (InvokeRequired)
        {
            BeginInvoke(() => AppendLogLine(line));
            return;
        }

        if (txtLogs.TextLength > 0)
        {
            txtLogs.AppendText(Environment.NewLine);
        }

        txtLogs.AppendText(line);
        txtLogs.SelectionStart = txtLogs.TextLength;
        txtLogs.ScrollToCaret();
    }

    private void RefreshLog()
    {
        txtLogs.Text = _loggingService.ReadRecentLines();
        txtLogs.SelectionStart = txtLogs.TextLength;
        txtLogs.ScrollToCaret();
    }

    private bool TryGetCurrentTemplate(out TemplateTaskConfig template)
    {
        if (dgvTemplates.CurrentRow?.DataBoundItem is TemplateTaskConfig config)
        {
            template = config;
            return true;
        }

        template = null!;
        return false;
    }

    private string GetTemplatePath(string fullPath)
    {
        var relativePath = Path.GetRelativePath(_configService.TemplatePath, fullPath);
        if (relativePath.StartsWith(".."))
        {
            throw new InvalidOperationException("Template 文件必须位于共享 Template 目录下。");
        }

        return relativePath;
    }

    private string GetSharedTemplateFullPath(string path)
    {
        return Path.GetFullPath(Path.Combine(_configService.TemplatePath, path));
    }

    protected override void WndProc(ref Message m)
    {
        if (_activationMessageId != 0 && m.Msg == _activationMessageId)
        {
            ShowFromTray();
            return;
        }

        base.WndProc(ref m);
    }

    private void Form1_Shown(object? sender, EventArgs e)
    {
        if (!_startMinimizedToTray)
        {
            return;
        }

        WindowState = FormWindowState.Minimized;
        Hide();
    }

    private void Form1_Resize(object? sender, EventArgs e)
    {
        if (WindowState == FormWindowState.Minimized)
        {
            Hide();
        }
    }

    private void ShowFromTray()
    {
        Show();
        if (WindowState == FormWindowState.Minimized)
        {
            WindowState = FormWindowState.Normal;
        }

        BringToFront();
        Activate();
    }

    private async void RunNowFromTray()
    {
        await RunSelectedTemplateAsync();
    }

    private void ExitApplication()
    {
        _allowClose = true;
        Close();
    }

    private void Form1_FormClosing(object? sender, FormClosingEventArgs e)
    {
        if (_allowClose)
        {
            _loggingService.LogWritten -= AppendLogLine;
            _schedulerService.Dispose();
            _trayService.Dispose();
            return;
        }

        e.Cancel = true;
        Hide();
    }
}
