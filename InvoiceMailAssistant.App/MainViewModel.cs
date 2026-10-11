using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Input;

namespace InvoiceMailAssistant.App;

public sealed class MainViewModel : INotifyPropertyChanged, IAsyncDisposable
{
    private readonly string _dataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "InvoiceMailAssistant");
    private readonly MailboxService _mailbox = new();
    private readonly ExcelWriter _excel = new();
    private readonly ExcelAssignmentCoordinator _excelAssignments;
    private readonly InvoiceParser _parser = new();
    private readonly SqliteInvoiceRepository _repository;
    private readonly DpapiCredentialStore _credentials;
    private readonly SemaphoreSlim _checkGate = new(1, 1);
    private readonly CancellationTokenSource _lifetime = new();
    private readonly object _lifecycleSync = new();
    private readonly string _settingsPath;
    private AppSettings _settings = new();
    private string _password = string.Empty;
    private string _statusText = "正在初始化";
    private string _connectionText = "未连接";
    private DateTime? _lastChecked;
    private bool _listening = true;
    private string? _reconciledExcelTargetKey;
    private bool _reconciliationPending;
    private Task? _pollingTask;
    private Task? _stopTask;
    public event PropertyChangedEventHandler? PropertyChanged;
    public event Action<IReadOnlyList<InvoiceSourceCount>>? NewApplicationsFound;
    private IReadOnlyList<InvoiceApplication> _records = Array.Empty<InvoiceApplication>();
    private IReadOnlyList<InvoiceApplication> _pendingRecords = Array.Empty<InvoiceApplication>();
    private IReadOnlyList<InvoiceSourceStatistics> _sourceStatistics = InvoiceSourceCatalog.CreateEmptyStatistics();
    public IReadOnlyList<InvoiceApplication> Records { get => _records; private set { _records = value; OnPropertyChanged(); } }
    public IReadOnlyList<InvoiceApplication> PendingRecords { get => _pendingRecords; private set { _pendingRecords = value; OnPropertyChanged(); } }
    public IReadOnlyList<InvoiceSourceStatistics> SourceStatistics { get => _sourceStatistics; private set { _sourceStatistics = value; OnPropertyChanged(); } }
    public ICommand CheckNowCommand { get; }
    public ICommand TestConnectionCommand { get; }
    public ICommand SaveSettingsCommand { get; }
    public ICommand RetryPendingCommand { get; }
    public ICommand ToggleListeningCommand { get; }
    public ICommand Scan2DaysCommand { get; }
    public ICommand ReparseFailedCommand { get; }

    public MainViewModel()
    {
        _settingsPath = Path.Combine(_dataDir, "settings.json");
        _repository = new SqliteInvoiceRepository(Path.Combine(_dataDir, "invoice-mail.db"));
        _excelAssignments = new ExcelAssignmentCoordinator(_repository, _excel, _dataDir);
        _credentials = new DpapiCredentialStore(Path.Combine(_dataDir, "credentials"));
        CheckNowCommand = new AsyncRelayCommand(() => CheckNowAsync(false));
        TestConnectionCommand = new AsyncRelayCommand(TestConnectionAsync);
        SaveSettingsCommand = new AsyncRelayCommand(SaveSettingsAsync);
        RetryPendingCommand = new AsyncRelayCommand(RetryPendingAsync);
        Scan2DaysCommand = new AsyncRelayCommand(() => ScanHistoryAsync(2));
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
    public string YuntianhuaWorksheetName { get => _settings.YuntianhuaWorksheetName; set { _settings.YuntianhuaWorksheetName = value; OnPropertyChanged(); } }
    public int PollSeconds { get => _settings.PollSeconds; set { _settings.PollSeconds = Math.Clamp(value, 30, 300); OnPropertyChanged(); } }
    public bool RunAtStartup { get => _settings.RunAtStartup; set { _settings.RunAtStartup = value; OnPropertyChanged(); } }
    public string StatusText { get => _statusText; private set { _statusText = value; OnPropertyChanged(); } }
    public string ConnectionText { get => _connectionText; private set { _connectionText = value; OnPropertyChanged(); } }
    public bool Listening { get => _listening; set { _listening = value; OnPropertyChanged(); OnPropertyChanged(nameof(ListeningButtonText)); } }
    public string ListeningButtonText => Listening ? "暂停监听" : "继续监听";
    public string LastCheckedText => _lastChecked is null ? "尚未检查" : _lastChecked.Value.ToString("yyyy-MM-dd HH:mm:ss");
    public int TodayReceived => SourceStatistics.Sum(x => x.TodayReceived);
    public int TodayCompleted => SourceStatistics.Sum(x => x.TodayCompleted);
    public int PendingCount => SourceStatistics.Sum(x => x.Pending);
    public int FailedCount => SourceStatistics.Sum(x => x.Failed);

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
        var assignmentsReady = await EnsureExcelAssignmentsAsync();
        await RefreshRecordsAsync();
        if (assignmentsReady)
        {
            await RepairCompletedRepeatedFormRecordsAsync();
            await RetryPendingAsync();
            StatusText = "就绪";
        }
        _pollingTask = RunPollingLoopAsync(_lifetime.Token);
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
            _reconciledExcelTargetKey = null;
            _reconciliationPending = false;
            StartupManager.SetEnabled(_settings.RunAtStartup);
            if (!string.IsNullOrWhiteSpace(previousMailbox) && !string.Equals(previousMailbox, normalizedAccount, StringComparison.OrdinalIgnoreCase))
                await _credentials.DeleteAsync(previousMailbox, _lifetime.Token);
            if (!string.IsNullOrWhiteSpace(EmailAccount) && !string.IsNullOrWhiteSpace(Password))
                await _credentials.SaveAsync(EmailAccount, Password);
            var assignmentsReady = await EnsureExcelAssignmentsAsync();
            if (assignmentsReady)
                await RepairCompletedRepeatedFormRecordsAsync();
            await RefreshRecordsAsync();
            if (assignmentsReady) StatusText = "设置已保存";
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
            if (!await EnsureExcelAssignmentsAsync()) return;
            var password = await ResolvePasswordAsync();
            var monitorFromUtc = monitorFromOverride ?? await EnsureMonitorStartAsync();
            var repairedExisting = await RepairCompletedRepeatedFormRecordsAsync();
            var messages = await _mailbox.FetchCandidateMessagesAsync(EmailAccount, password, ImapHost, ImapPort, monitorFromUtc, 200, _lifetime.Token);
            var addedBySource = new Dictionary<InvoiceSource, int>();

            foreach (var mail in messages)
            {
                if (!string.IsNullOrWhiteSpace(mail.FetchError))
                {
                    var failedFetch = CreateFetchFailed(mail, mail.FetchError);
                    var insertedFailure = await _repository.TryInsertAsync(failedFetch, CreateFailureHash(mail), _lifetime.Token);
                    if (insertedFailure is not null) IncrementSourceCount(addedBySource, mail.Source);
                    continue;
                }

                var parsed = _parser.Parse(mail, EmailAccount.Trim().ToLowerInvariant());
                if (!parsed.Success || parsed.Application is null)
                {
                    var failed = CreateParseFailed(mail, parsed.Error ?? "邮件解析失败");
                    var failureHash = CreateFailureHash(mail);
                    var insertedFailure = await _repository.TryInsertAsync(failed, failureHash, _lifetime.Token);
                    if (insertedFailure is not null) IncrementSourceCount(addedBySource, mail.Source);
                    continue;
                }

                var app = parsed.Application;
                var hash = Deduplication.CreateFallbackHash(app);
                var insertedId = await _repository.TryInsertAsync(app, hash, _lifetime.Token);
                if (insertedId is null) continue;

                app.Id = insertedId.Value;
                await _repository.UpdateStatusAsync(app.Id, ProcessingStatus.PendingExcel, cancellationToken: _lifetime.Token);
                await TryWriteExcelAsync(app);
                IncrementSourceCount(addedBySource, app.Source);
            }

            var repaired = await ReparseFailedCoreAsync();
            await RetryPendingCoreAsync();
            _lastChecked = DateTime.Now;
            ConnectionText = $"已连接：{EmailAccount.Trim()}";
            var sourceCounts = InvoiceSourceCatalog.CreateCounts(addedBySource);
            var added = sourceCounts.Sum(x => x.Count);
            if (added > 0) NewApplicationsFound?.Invoke(sourceCounts);
            if (!automatic || added > 0 || repaired > 0 || repairedExisting > 0
                || StatusText.StartsWith("检查失败", StringComparison.Ordinal)
                || StatusText.StartsWith("自动监听异常", StringComparison.Ordinal))
                StatusText = $"检查完成，本次发现 {added} 条新申请，修复 {repairedExisting} 条历史记录，重新处理 {repaired} 条失败记录";
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
            try
            {
                await RefreshRecordsAsync();
            }
            catch (OperationCanceledException) when (_lifetime.IsCancellationRequested)
            {
            }
            catch (Exception ex)
            {
                StatusText = $"记录刷新失败：{ex.Message}";
            }
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

    private async Task<int> RepairCompletedRepeatedFormRecordsAsync()
    {
        if (string.IsNullOrWhiteSpace(ExcelPath) || !File.Exists(ExcelPath)) return 0;

        var candidates = await _repository.GetCompletedRepeatedFormRecordsAsync(_lifetime.Token);
        var repaired = 0;
        foreach (var original in candidates)
        {
            var parsed = _parser.Parse(
                new MailEnvelope(
                    original.ImapUid,
                    original.MessageId,
                    original.MailFrom,
                    original.MailSubject,
                    original.MailReceivedAt,
                    original.NormalizedBody,
                    original.UidValidity,
                    original.MailboxName,
                    Source: original.Source),
                original.MailboxIdentity);
            if (!parsed.Success || parsed.Application is null) continue;

            var corrected = parsed.Application;
            if (BusinessFieldsEqual(original, corrected)) continue;

            corrected.Id = original.Id;
            corrected.ExcelRow = original.ExcelRow;
            corrected.ExcelTargetKey = original.ExcelTargetKey;
            corrected.CreatedAt = original.CreatedAt;
            try
            {
                // Repair the existing row only after confirming its A/B/C/D/F
                // values still match the old database record. This prevents a
                // manual edit from being mistaken for the stale parser output.
                await _excel.RepairExistingRowAsync(original, corrected, ExcelPath, GetWorksheetName(corrected.Source), _lifetime.Token, corrected.Source);
                await _repository.UpdateParsedAsync(original.Id, corrected, Deduplication.CreateFallbackHash(corrected), _lifetime.Token);
                await _repository.UpdateStatusAsync(original.Id, ProcessingStatus.Completed, excelRow: corrected.ExcelRow, cancellationToken: _lifetime.Token);
                repaired++;
            }
            catch (ExcelRowOccupiedException)
            {
                // Keep the old completed record and leave the user's row
                // untouched. The mismatch must be resolved manually.
            }
            catch (IOException)
            {
                // A locked/missing workbook is handled by the normal pending
                // workflow after the user restores access to the file.
            }
        }

        return repaired;
    }

    private static bool BusinessFieldsEqual(InvoiceApplication first, InvoiceApplication second)
        => string.Equals(first.CompanyName, second.CompanyName, StringComparison.Ordinal)
            && string.Equals(first.CreditCode, second.CreditCode, StringComparison.Ordinal)
            && first.Amount == second.Amount
            && first.ApplyTime == second.ApplyTime
            && string.Equals(first.InvoiceType, second.InvoiceType, StringComparison.Ordinal)
            && string.Equals(first.Recipient, second.Recipient, StringComparison.Ordinal)
            && string.Equals(first.Phone, second.Phone, StringComparison.Ordinal)
            && string.Equals(first.Address, second.Address, StringComparison.Ordinal)
            && string.Equals(first.Email, second.Email, StringComparison.OrdinalIgnoreCase)
            && string.Equals(first.Remark, second.Remark, StringComparison.Ordinal);

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
                item.MailboxName,
                Source: item.Source);
            var parsed = _parser.Parse(mail, item.MailboxIdentity);
            if (!parsed.Success || parsed.Application is null) continue;

            var app = parsed.Application;
            app.Id = item.Id;
            app.ExcelRow = item.ExcelRow;
            app.ExcelTargetKey = item.ExcelTargetKey;
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
            var worksheetName = GetWorksheetName(app.Source);
            var row = await _excelAssignments.WriteApplicationAsync(app, ExcelPath, worksheetName, _lifetime.Token, app.Source);
            app.ExcelRow = row;
            app.ExcelTargetKey = ExcelTargetIdentity.Create(ExcelPath, worksheetName);
            app.ProcessingStatus = ProcessingStatus.Completed;
            app.ErrorMessage = string.Empty;
        }
        catch (ExcelRowOccupiedException ex)
        {
            app.ProcessingStatus = ProcessingStatus.PendingExcel;
            app.ErrorMessage = ex.Message;
            await _repository.UpdateStatusAsync(app.Id, ProcessingStatus.PendingExcel, ex.Message, cancellationToken: _lifetime.Token);
        }
        catch (IOException ex)
        {
            app.ProcessingStatus = ProcessingStatus.PendingExcel;
            app.ErrorMessage = ex.Message;
            await _repository.UpdateStatusAsync(app.Id, ProcessingStatus.PendingExcel, ex.Message, cancellationToken: _lifetime.Token);
        }
        catch (Exception ex)
        {
            app.ProcessingStatus = ProcessingStatus.ExcelFailed;
            app.ErrorMessage = ex.Message;
            await _repository.UpdateStatusAsync(app.Id, ProcessingStatus.ExcelFailed, ex.Message, cancellationToken: _lifetime.Token);
        }
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
        var pending = await _repository.GetPendingExcelAsync(_lifetime.Token);
        var sourceStatistics = await _repository.GetSourceStatisticsAsync(DateTimeOffset.Now, _lifetime.Token);
        Records = items;
        PendingRecords = pending;
        SourceStatistics = sourceStatistics;
        OnPropertyChanged(nameof(TodayReceived));
        OnPropertyChanged(nameof(TodayCompleted));
        OnPropertyChanged(nameof(PendingCount));
        OnPropertyChanged(nameof(FailedCount));
    }

    private string GetWorksheetName(InvoiceSource source)
        => (source == InvoiceSource.Yuntianhua ? _settings.YuntianhuaWorksheetName : _settings.WorksheetName).Trim();

    private async Task<bool> EnsureExcelAssignmentsAsync()
    {
        if (string.IsNullOrWhiteSpace(ExcelPath)) return true;

        var targets = InvoiceSourceCatalog.All
            .Select(definition => (definition.Source, WorksheetName: GetWorksheetName(definition.Source).Trim()))
            .Where(target => target.WorksheetName.Length > 0)
            .GroupBy(target => target.WorksheetName, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .ToArray();
        var targetSignature = string.Join("|", targets.Select(target => $"{target.Source}:{ExcelTargetIdentity.Create(ExcelPath, target.WorksheetName)}"));
        if (string.Equals(_reconciledExcelTargetKey, targetSignature, StringComparison.Ordinal) && !_reconciliationPending)
            return true;

        try
        {
            var mappedExistingRows = 0;
            var writtenRows = 0;
            foreach (var target in targets)
            {
                var result = await _excelAssignments.ReconcileAsync(ExcelPath, target.WorksheetName, _lifetime.Token, target.Source);
                mappedExistingRows += result.MappedExistingRows;
                writtenRows += result.WrittenRows;
                if (result.PendingRows > 0)
                {
                    _reconciledExcelTargetKey = targetSignature;
                    _reconciliationPending = true;
                    StatusText = $"历史 Excel 核对有 {result.PendingRows} 条等待处理，请关闭占用的登记表后重试";
                    return false;
                }
            }

            _reconciledExcelTargetKey = targetSignature;
            _reconciliationPending = false;
            if (mappedExistingRows > 0 || writtenRows > 0)
                StatusText = $"已自动核对历史 Excel：复用 {mappedExistingRows} 行，补写 {writtenRows} 行";
            return true;
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            StatusText = $"历史 Excel 核对暂停：{ex.Message}";
            return false;
        }
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
        Source = mail.Source,
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
        Source = mail.Source,
        MailReceivedAt = mail.ReceivedAt,
        MailSubject = mail.Subject,
        MailFrom = mail.FromAddress,
        NormalizedBody = InvoiceParser.NormalizeBody(mail.BodyText),
        ProcessingStatus = ProcessingStatus.MailFailed,
        ErrorMessage = error
    };

    private static string CreateFailureHash(MailEnvelope mail)
    {
        var source = $"FAIL|{mail.Source}|{mail.MailboxName}|{mail.FromAddress}|{mail.Subject}|{mail.ReceivedAt:O}|{InvoiceParser.NormalizeBody(mail.BodyText)}";
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(source)));
    }

    private static void IncrementSourceCount(Dictionary<InvoiceSource, int> counts, InvoiceSource source)
        => counts[source] = counts.TryGetValue(source, out var current) ? current + 1 : 1;

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

    public Task StopAsync()
    {
        lock (_lifecycleSync)
            return _stopTask ??= StopCoreAsync();
    }

    private async Task StopCoreAsync()
    {
        _lifetime.Cancel();
        if (_pollingTask is not null)
        {
            try
            {
                await _pollingTask.ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (_lifetime.IsCancellationRequested)
            {
            }
        }

        await _mailbox.DisposeAsync().ConfigureAwait(false);
        _checkGate.Dispose();
        _lifetime.Dispose();
    }

    public ValueTask DisposeAsync() => new(StopAsync());
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
