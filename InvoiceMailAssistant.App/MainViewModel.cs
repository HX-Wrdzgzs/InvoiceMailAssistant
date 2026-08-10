using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Input;

namespace InvoiceMailAssistant.App;

public sealed class MainViewModel : INotifyPropertyChanged, IDisposable
{
    private readonly string _dataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "InvoiceMailAssistant");
    private readonly MailboxService _mailbox = new();
    private readonly ExcelWriter _excel = new();
    private readonly InvoiceParser _parser = new();
    private readonly SqliteInvoiceRepository _repository;
    private readonly DpapiCredentialStore _credentials;
    private readonly SemaphoreSlim _checkGate = new(1, 1);
    private readonly CancellationTokenSource _lifetime = new();
    private readonly string _settingsPath;
    private AppSettings _settings = new();
    private string _password = string.Empty;
    private string _statusText = "正在初始化";
    private string _connectionText = "未连接";
    private DateTime? _lastChecked;
    private bool _listening = true;
    private int _customHistoryDays = 7;

    public event PropertyChangedEventHandler? PropertyChanged;
    public event Action<int>? NewApplicationsFound;
    public ObservableCollection<InvoiceApplication> Records { get; } = [];
    public ObservableCollection<InvoiceApplication> PendingRecords { get; } = [];
    public ICommand CheckNowCommand { get; }
    public ICommand TestConnectionCommand { get; }
    public ICommand SaveSettingsCommand { get; }
    public ICommand RetryPendingCommand { get; }
    public ICommand ToggleListeningCommand { get; }
    public ICommand Scan7DaysCommand { get; }
    public ICommand Scan30DaysCommand { get; }
    public ICommand ScanCustomCommand { get; }
    public ICommand ReparseFailedCommand { get; }

    public MainViewModel()
    {
        _settingsPath = Path.Combine(_dataDir, "settings.json");
        _repository = new SqliteInvoiceRepository(Path.Combine(_dataDir, "invoice-mail.db"));
        _credentials = new DpapiCredentialStore(Path.Combine(_dataDir, "credentials"));
        CheckNowCommand = new AsyncRelayCommand(() => CheckNowAsync(false));
        TestConnectionCommand = new AsyncRelayCommand(TestConnectionAsync);
        SaveSettingsCommand = new AsyncRelayCommand(SaveSettingsAsync);
        RetryPendingCommand = new AsyncRelayCommand(RetryPendingAsync);
        Scan7DaysCommand = new AsyncRelayCommand(() => ScanHistoryAsync(7));
        Scan30DaysCommand = new AsyncRelayCommand(() => ScanHistoryAsync(30));
        ScanCustomCommand = new AsyncRelayCommand(() => ScanHistoryAsync(CustomHistoryDays));
        ReparseFailedCommand = new AsyncRelayCommand(ReparseFailedAsync);
        ToggleListeningCommand = new RelayCommand(() =>
        {
            Listening = !Listening;
            StatusText = Listening ? "自动监听已继续" : "自动监听已暂停";
        });
    }

    public string EmailAccount { get => _settings.EmailAccount; set { _settings.EmailAccount = value; OnPropertyChanged(); } }
    public string Password { get => _password; set { _password = value; OnPropertyChanged(); } }
    public string ImapHost { get => _settings.ImapHost; set { _settings.ImapHost = value; OnPropertyChanged(); } }
    public int ImapPort { get => _settings.ImapPort; set { _settings.ImapPort = value; OnPropertyChanged(); } }
    public string ExcelPath { get => _settings.ExcelPath; set { _settings.ExcelPath = value; OnPropertyChanged(); } }
    public string WorksheetName { get => _settings.WorksheetName; set { _settings.WorksheetName = value; OnPropertyChanged(); } }
    public int PollSeconds { get => _settings.PollSeconds; set { _settings.PollSeconds = Math.Clamp(value, 30, 300); OnPropertyChanged(); } }
    public bool RunAtStartup { get => _settings.RunAtStartup; set { _settings.RunAtStartup = value; OnPropertyChanged(); } }
    public int CustomHistoryDays { get => _customHistoryDays; set { _customHistoryDays = Math.Clamp(value, 1, 3650); OnPropertyChanged(); } }
    public string StatusText { get => _statusText; private set { _statusText = value; OnPropertyChanged(); } }
    public string ConnectionText { get => _connectionText; private set { _connectionText = value; OnPropertyChanged(); } }
    public bool Listening { get => _listening; set { _listening = value; OnPropertyChanged(); OnPropertyChanged(nameof(ListeningButtonText)); } }
    public string ListeningButtonText => Listening ? "暂停监听" : "继续监听";
    public string LastCheckedText => _lastChecked is null ? "尚未检查" : _lastChecked.Value.ToString("yyyy-MM-dd HH:mm:ss");
    public int TodayReceived => Records.Count(x => x.MailReceivedAt.LocalDateTime.Date == DateTime.Today);
    public int TodayCompleted => Records.Count(x => x.ProcessingStatus == ProcessingStatus.Completed && x.UpdatedAt.LocalDateTime.Date == DateTime.Today);
    public int PendingCount => Records.Count(x => x.ProcessingStatus is ProcessingStatus.PendingExcel or ProcessingStatus.ExcelFailed);
    public int FailedCount => Records.Count(x => x.ProcessingStatus is ProcessingStatus.ParseFailed or ProcessingStatus.MailFailed);

    public async Task InitializeAsync()
    {
        Directory.CreateDirectory(_dataDir);
        try
        {
            _settings = await AppSettings.LoadAsync(_settingsPath);
        }
        catch (Exception ex) when (ex is InvalidDataException or System.Text.Json.JsonException or IOException)
        {
            _settings = new AppSettings();
            StatusText = $"设置文件无法读取，已使用默认设置：{ex.Message}";
        }

        await _repository.InitializeAsync();
        try { StartupManager.SetEnabled(_settings.RunAtStartup); }
        catch (Exception ex) { StatusText = ex.Message; }
        if (!string.IsNullOrWhiteSpace(_settings.EmailAccount))
        {
            await EnsureMonitorStartAsync();
            try
            {
                _password = await _credentials.LoadAsync(_settings.EmailAccount) ?? string.Empty;
            }
            catch (Exception ex) when (ex is CryptographicException or IOException)
            {
                _password = string.Empty;
                StatusText = $"邮箱凭据无法读取，请重新输入：{ex.Message}";
            }
        }
        OnPropertyChanged(string.Empty);
        await RefreshRecordsAsync();
        await RetryPendingAsync();
        StatusText = "就绪";
        _ = RunPollingLoopAsync(_lifetime.Token);
    }

    private async Task SaveSettingsAsync()
    {
        try
        {
            var previousMailbox = _settings.MailboxIdentity;
            var normalizedAccount = EmailAccount.Trim().ToLowerInvariant();
            if (!string.IsNullOrWhiteSpace(EmailAccount))
                await EnsureMonitorStartAsync();
            else
            {
                _settings.MailboxIdentity = string.Empty;
                _settings.MonitorFromUtc = null;
            }

            await _settings.SaveAsync(_settingsPath);
            StartupManager.SetEnabled(_settings.RunAtStartup);
            if (!string.IsNullOrWhiteSpace(previousMailbox) && !string.Equals(previousMailbox, normalizedAccount, StringComparison.OrdinalIgnoreCase))
                await _credentials.DeleteAsync(previousMailbox, _lifetime.Token);
            if (!string.IsNullOrWhiteSpace(EmailAccount) && !string.IsNullOrWhiteSpace(Password))
                await _credentials.SaveAsync(EmailAccount, Password);
            StatusText = "设置已保存";
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            StatusText = $"设置保存失败：{ex.Message}";
        }
    }

    private async Task TestConnectionAsync()
    {
        try
        {
            var password = await ResolvePasswordAsync();
            await _mailbox.TestConnectionAsync(EmailAccount, password, ImapHost, ImapPort, _lifetime.Token);
            ConnectionText = $"已连接：{EmailAccount.Trim()}";
            StatusText = "邮箱连接正常";
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            ConnectionText = $"连接失败：{EmailAccount.Trim()}";
            StatusText = $"邮箱连接失败：{ex.Message}";
        }
    }

    private async Task CheckNowAsync(bool automatic, DateTimeOffset? monitorFromOverride = null)
    {
        if (!await _checkGate.WaitAsync(0)) return;
        try
        {
            if (string.IsNullOrWhiteSpace(EmailAccount) || string.IsNullOrWhiteSpace(ExcelPath))
            {
                if (!automatic) StatusText = "请先填写邮箱账户和 Excel 文件路径";
                return;
            }

            if (!automatic) StatusText = "正在检查邮件";
            var password = await ResolvePasswordAsync();
            var monitorFromUtc = monitorFromOverride ?? await EnsureMonitorStartAsync();
            var messages = await _mailbox.FetchCandidateMessagesAsync(EmailAccount, password, ImapHost, ImapPort, monitorFromUtc, 200, _lifetime.Token);
            var added = 0;

            foreach (var mail in messages)
            {
                if (!string.IsNullOrWhiteSpace(mail.FetchError))
                {
                    var failedFetch = CreateFetchFailed(mail, mail.FetchError);
                    await _repository.TryInsertAsync(failedFetch, CreateFailureHash(mail), _lifetime.Token);
                    continue;
                }

                var parsed = _parser.Parse(mail, EmailAccount.Trim().ToLowerInvariant());
                if (!parsed.Success || parsed.Application is null)
                {
                    var failed = CreateParseFailed(mail, parsed.Error ?? "邮件解析失败");
                    var failureHash = CreateFailureHash(mail);
                    await _repository.TryInsertAsync(failed, failureHash, _lifetime.Token);
                    continue;
                }

                var app = parsed.Application;
                var hash = Deduplication.CreateFallbackHash(app);
                var insertedId = await _repository.TryInsertAsync(app, hash, _lifetime.Token);
                if (insertedId is null) continue;

                app.Id = insertedId.Value;
                await _repository.UpdateStatusAsync(app.Id, ProcessingStatus.PendingExcel, cancellationToken: _lifetime.Token);
                await TryWriteExcelAsync(app);
                added++;
            }

            var repaired = await ReparseFailedCoreAsync();
            await RetryPendingCoreAsync();
            _lastChecked = DateTime.Now;
            ConnectionText = $"已连接：{EmailAccount.Trim()}";
            if (added > 0) NewApplicationsFound?.Invoke(added);
            if (!automatic || added > 0 || repaired > 0)
                StatusText = $"检查完成，本次发现 {added} 条新申请，重新处理 {repaired} 条失败记录";
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            ConnectionText = $"连接失败：{EmailAccount.Trim()}";
            StatusText = $"检查失败：{ex.Message}";
        }
        finally
        {
            _checkGate.Release();
            OnPropertyChanged(nameof(LastCheckedText));
            await RefreshRecordsAsync();
        }
    }

    private async Task RetryPendingAsync()
    {
        if (!await _checkGate.WaitAsync(0)) return;
        try
        {
            await RetryPendingCoreAsync();
            await RefreshRecordsAsync();
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            StatusText = $"重试等待项失败：{ex.Message}";
        }
        finally
        {
            _checkGate.Release();
        }
    }

    private async Task RetryPendingCoreAsync()
    {
        if (string.IsNullOrWhiteSpace(ExcelPath) || !File.Exists(ExcelPath)) return;
        var pending = await _repository.GetPendingExcelAsync(_lifetime.Token);
        foreach (var item in pending)
            await TryWriteExcelAsync(item);
    }

    private async Task ScanHistoryAsync(int days)
    {
        days = Math.Clamp(days, 1, 3650);
        if (string.IsNullOrWhiteSpace(EmailAccount) || string.IsNullOrWhiteSpace(ExcelPath))
        {
            StatusText = "请先填写邮箱账户和 Excel 文件路径";
            return;
        }

        StatusText = $"正在补扫最近 {days} 天邮件";
        await CheckNowAsync(false, DateTimeOffset.UtcNow.AddDays(-days));
        StatusText = $"历史补扫完成：最近 {days} 天";
    }

    private async Task ReparseFailedAsync()
    {
        if (!await _checkGate.WaitAsync(0)) return;
        try
        {
            var repaired = await ReparseFailedCoreAsync();

            StatusText = $"已重新解析 {repaired} 条失败记录";
            await RefreshRecordsAsync();
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            StatusText = $"重新解析失败：{ex.Message}";
        }
        finally
        {
            _checkGate.Release();
        }
    }

    private async Task<int> ReparseFailedCoreAsync()
    {
        var failed = await _repository.GetParseFailedAsync(_lifetime.Token);
        var repaired = 0;
        foreach (var item in failed)
        {
            var mail = new MailEnvelope(
                item.ImapUid,
                item.MessageId,
                item.MailFrom,
                item.MailSubject,
                item.MailReceivedAt,
                item.NormalizedBody,
                item.UidValidity,
                item.MailboxName);
            var parsed = _parser.Parse(mail, item.MailboxIdentity);
            if (!parsed.Success || parsed.Application is null) continue;

            var app = parsed.Application;
            app.Id = item.Id;
            app.ExcelRow = item.ExcelRow;
            app.CreatedAt = item.CreatedAt;
            await _repository.UpdateParsedAsync(app.Id, app, Deduplication.CreateFallbackHash(app), _lifetime.Token);
            await TryWriteExcelAsync(app);
            repaired++;
        }

        return repaired;
    }

    private async Task TryWriteExcelAsync(InvoiceApplication app)
    {
        try
        {
            await PlanAndWriteExcelAsync(app);
        }
        catch (ExcelRowOccupiedException)
        {
            try
            {
                // The persisted row is only a recovery hint. If a user filled it
                // after planning, recompute a safe row before retrying.
                await PlanAndWriteExcelAsync(app);
            }
            catch (IOException ex)
            {
                await _repository.UpdateStatusAsync(app.Id, ProcessingStatus.PendingExcel, ex.Message, cancellationToken: _lifetime.Token);
            }
            catch (Exception ex)
            {
                await _repository.UpdateStatusAsync(app.Id, ProcessingStatus.ExcelFailed, ex.Message, cancellationToken: _lifetime.Token);
            }
        }
        catch (IOException ex)
        {
            await _repository.UpdateStatusAsync(app.Id, ProcessingStatus.PendingExcel, ex.Message, cancellationToken: _lifetime.Token);
        }
        catch (Exception ex)
        {
            await _repository.UpdateStatusAsync(app.Id, ProcessingStatus.ExcelFailed, ex.Message, cancellationToken: _lifetime.Token);
        }
    }

    private async Task PlanAndWriteExcelAsync(InvoiceApplication app)
    {
        using var writeLock = await _excel.AcquireWriteLockAsync(_lifetime.Token);
        var plannedRow = _excel.ResolveTargetRow(app, ExcelPath, WorksheetName);
        app.ExcelRow = plannedRow;
        await _repository.UpdateStatusAsync(app.Id, ProcessingStatus.PendingExcel, excelRow: plannedRow, cancellationToken: _lifetime.Token);

        var row = await _excel.WriteAsync(app, ExcelPath, WorksheetName, _lifetime.Token, writeLock);
        await _repository.UpdateStatusAsync(app.Id, ProcessingStatus.Completed, excelRow: row, cancellationToken: _lifetime.Token);
    }

    private async Task RunPollingLoopAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            try
            {
                await Task.Delay(TimeSpan.FromSeconds(Math.Clamp(PollSeconds, 30, 300)), cancellationToken);
                if (Listening) await CheckNowAsync(true);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                break;
            }
            catch (Exception ex)
            {
                StatusText = $"自动监听异常：{ex.Message}";
                try { await Task.Delay(TimeSpan.FromSeconds(30), cancellationToken); }
                catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { break; }
            }
        }
    }

    private async Task RefreshRecordsAsync()
    {
        var items = await _repository.GetRecentAsync(200, _lifetime.Token);
        Records.Clear();
        foreach (var item in items) Records.Add(item);
        var pending = await _repository.GetPendingExcelAsync(_lifetime.Token);
        PendingRecords.Clear();
        foreach (var item in pending) PendingRecords.Add(item);
        OnPropertyChanged(nameof(TodayReceived));
        OnPropertyChanged(nameof(TodayCompleted));
        OnPropertyChanged(nameof(PendingCount));
        OnPropertyChanged(nameof(FailedCount));
        OnPropertyChanged(nameof(PendingRecords));
    }

    private async Task<DateTimeOffset> EnsureMonitorStartAsync()
    {
        if (AppSettings.EnsureMonitorStart(_settings, EmailAccount, DateTimeOffset.UtcNow, out var monitorFromUtc))
            await _settings.SaveAsync(_settingsPath);
        return monitorFromUtc;
    }

    private async Task<string> ResolvePasswordAsync()
    {
        if (!string.IsNullOrWhiteSpace(Password)) return Password;
        var password = await _credentials.LoadAsync(EmailAccount, _lifetime.Token);
        if (string.IsNullOrWhiteSpace(password)) throw new InvalidOperationException("尚未保存邮箱密码或客户端专用密码。");
        return password;
    }

    private InvoiceApplication CreateParseFailed(MailEnvelope mail, string error) => new()
    {
        MessageId = mail.MessageId.Trim(),
        ImapUid = mail.Uid,
        UidValidity = mail.UidValidity,
        MailboxName = mail.MailboxName,
        MailboxIdentity = EmailAccount.Trim().ToLowerInvariant(),
        MailReceivedAt = mail.ReceivedAt,
        MailSubject = mail.Subject,
        MailFrom = mail.FromAddress,
        NormalizedBody = InvoiceParser.NormalizeBody(mail.BodyText),
        ProcessingStatus = ProcessingStatus.ParseFailed,
        ErrorMessage = error
    };

    private InvoiceApplication CreateFetchFailed(MailEnvelope mail, string error) => new()
    {
        MessageId = mail.MessageId.Trim(),
        ImapUid = mail.Uid,
        UidValidity = mail.UidValidity,
        MailboxName = mail.MailboxName,
        MailboxIdentity = EmailAccount.Trim().ToLowerInvariant(),
        MailReceivedAt = mail.ReceivedAt,
        MailSubject = mail.Subject,
        MailFrom = mail.FromAddress,
        NormalizedBody = InvoiceParser.NormalizeBody(mail.BodyText),
        ProcessingStatus = ProcessingStatus.MailFailed,
        ErrorMessage = error
    };

    private static string CreateFailureHash(MailEnvelope mail)
    {
        var source = $"FAIL|{mail.MailboxName}|{mail.FromAddress}|{mail.Subject}|{mail.ReceivedAt:O}|{InvoiceParser.NormalizeBody(mail.BodyText)}";
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(source)));
    }

    private void OnPropertyChanged([CallerMemberName] string? name = null)
    {
        if (string.IsNullOrEmpty(name))
        {
            foreach (var property in GetType().GetProperties().Where(x => x.CanRead))
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property.Name));
            return;
        }
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public void Dispose()
    {
        _lifetime.Cancel();
        _mailbox.Dispose();
        _lifetime.Dispose();
        _checkGate.Dispose();
    }
}

public sealed class AsyncRelayCommand(Func<Task> execute, Func<bool>? canExecute = null) : ICommand
{
    private bool _running;
    public event EventHandler? CanExecuteChanged;
    public bool CanExecute(object? parameter) => !_running && (canExecute?.Invoke() ?? true);
    public async void Execute(object? parameter)
    {
        if (!CanExecute(parameter)) return;
        _running = true;
        CanExecuteChanged?.Invoke(this, EventArgs.Empty);
        try { await execute(); }
        finally { _running = false; CanExecuteChanged?.Invoke(this, EventArgs.Empty); }
    }
}

public sealed class RelayCommand(Action execute) : ICommand
{
    event EventHandler? ICommand.CanExecuteChanged
    {
        add { }
        remove { }
    }

    public bool CanExecute(object? parameter) => true;
    public void Execute(object? parameter) => execute();
}
