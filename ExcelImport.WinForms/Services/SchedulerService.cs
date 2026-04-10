using System.Collections.Concurrent;
using System.Threading;
using ExcelImport.Core.Models;

namespace ExcelImport.Services;

public sealed class SchedulerService : IDisposable
{
    private readonly ImportService _importService;
    private readonly LoggingService _loggingService;
    private readonly ConcurrentDictionary<string, System.Threading.Timer> _timers = new(StringComparer.OrdinalIgnoreCase);
    private readonly ConcurrentDictionary<string, SemaphoreSlim> _locks = new(StringComparer.OrdinalIgnoreCase);

    public SchedulerService(ImportService importService, LoggingService loggingService)
    {
        _importService = importService;
        _loggingService = loggingService;
    }

    public void Start(AppSettings settings)
    {
        _loggingService.Info($"启动调度器，模板总数: {settings.Templates.Count}，启用数: {settings.Templates.Count(t => t.Enabled)}。");
        Stop();

        foreach (var template in settings.Templates.Where(t => t.Enabled))
        {
            var semaphore = new SemaphoreSlim(1, 1);
            _locks[template.Name] = semaphore;

            _loggingService.Info($"注册模板定时任务: {template.Name}，间隔: {Math.Max(1, template.IntervalMinutes)} 分钟，目录: {template.WatchPath}，模板文件: {template.TemplateFile}");
            var timer = new System.Threading.Timer(_ => QueueTemplateExecution(settings, template), null, TimeSpan.Zero, TimeSpan.FromMinutes(Math.Max(1, template.IntervalMinutes)));
            _timers[template.Name] = timer;
        }
    }

    public void Stop()
    {
        if (_timers.Count > 0)
        {
            _loggingService.Info($"停止调度器，清理 {_timers.Count} 个定时任务。");
        }

        foreach (var timer in _timers.Values)
        {
            timer.Dispose();
        }

        foreach (var semaphore in _locks.Values)
        {
            semaphore.Dispose();
        }

        _timers.Clear();
        _locks.Clear();
    }

    public async Task<ImportResult> RunOnceAsync(AppSettings settings, TemplateTaskConfig template, CancellationToken cancellationToken = default)
    {
        return await _importService.RunTemplateAsync(settings, template, cancellationToken);
    }

    private void QueueTemplateExecution(AppSettings settings, TemplateTaskConfig template)
    {
        _ = ExecuteTemplateAsync(settings, template);
    }

    private async Task ExecuteTemplateAsync(AppSettings settings, TemplateTaskConfig template)
    {
        _loggingService.Info($"触发模板定时执行: {template.Name}");

        if (!_locks.TryGetValue(template.Name, out var semaphore))
        {
            _loggingService.Warning($"模板 {template.Name} 未找到调度锁，跳过执行。");
            return;
        }

        if (!await semaphore.WaitAsync(0))
        {
            _loggingService.Warning($"模板 {template.Name} 上一轮尚未结束，本轮跳过。");
            return;
        }

        try
        {
            await _importService.RunTemplateAsync(settings, template, CancellationToken.None);
        }
        catch (Exception ex)
        {
            _loggingService.Error($"模板执行异常: {template.Name}", ex);
        }
        finally
        {
            semaphore.Release();
        }
    }

    public void Dispose()
    {
        Stop();
    }
}
