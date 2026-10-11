using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace InvoiceMailAssistant.App;

public static class ExcelTargetIdentity
{
    public static string Create(string filePath, string worksheetName)
    {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("Excel 文件路径不能为空。", nameof(filePath));
        if (string.IsNullOrWhiteSpace(worksheetName)) throw new ArgumentException("工作表名称不能为空。", nameof(worksheetName));

        var fullPath = Path.TrimEndingDirectorySeparator(Path.GetFullPath(filePath.Trim()));
        if (OperatingSystem.IsWindows()) fullPath = fullPath.ToUpperInvariant();
        var source = $"{fullPath}\0{worksheetName.Trim().ToUpperInvariant()}";
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(source)));
    }
}

public sealed record ExcelRowSnapshot(
    int RowNumber,
    int? Month,
    int? Day,
    string CompanyName,
    string CreditCode,
    decimal? Amount,
    string Email)
{
    public bool Matches(InvoiceApplication application)
    {
        if (!string.Equals(CompanyName.Trim(), application.CompanyName.Trim(), StringComparison.Ordinal)) return false;
        if (!string.Equals(CreditCode.Trim(), application.CreditCode.Trim(), StringComparison.Ordinal)) return false;
        if (Amount is null || Amount.Value != application.Amount) return false;
        if (!string.Equals(Email.Trim(), application.Email.Trim(), StringComparison.OrdinalIgnoreCase)) return false;
        return (Month is null || Day is null)
            || (Month == application.ApplyTime.Month && Day == application.ApplyTime.Day);
    }
}

public sealed record ExcelWorkbookSnapshot(int LastUsedRow, IReadOnlyList<ExcelRowSnapshot> Rows)
{
    public ExcelRowSnapshot? FindRow(int rowNumber)
        => Rows.FirstOrDefault(row => row.RowNumber == rowNumber);

    public IReadOnlyList<int> FindMatchingRows(InvoiceApplication application, IReadOnlySet<int> claimedRows)
        => Rows
            .Where(row => !claimedRows.Contains(row.RowNumber) && row.Matches(application))
            .Select(row => row.RowNumber)
            .OrderBy(row => row)
            .ToArray();
}

public sealed record ExcelReconciliationResult(
    int MappedExistingRows,
    int WrittenRows,
    int PendingRows,
    string? BackupDirectory)
{
    public bool HasChanges => MappedExistingRows > 0 || WrittenRows > 0 || PendingRows > 0;
    public bool IsReady => PendingRows == 0;
}

/// <summary>
/// Coordinates SQLite row ownership and the Excel package update. The
/// database reservation is deliberately made before the package is changed;
/// a process crash therefore leaves enough information for the next retry to
/// safely reuse the same row.
/// </summary>
public sealed class ExcelAssignmentCoordinator
{
    private readonly SqliteInvoiceRepository _repository;
    private readonly ExcelWriter _writer;
    private readonly string _dataDirectory;

    public ExcelAssignmentCoordinator(SqliteInvoiceRepository repository, ExcelWriter writer, string dataDirectory)
    {
        _repository = repository;
        _writer = writer;
        _dataDirectory = dataDirectory;
    }

    public async Task<int> WriteApplicationAsync(
        InvoiceApplication application,
        string filePath,
        string worksheetName,
        CancellationToken cancellationToken = default,
        InvoiceSource? source = null)
    {
        var effectiveSource = source ?? application.Source;
        var targetKey = ExcelTargetIdentity.Create(filePath, worksheetName);
        if (!string.Equals(application.ExcelTargetKey, targetKey, StringComparison.Ordinal))
        {
            application.ExcelTargetKey = targetKey;
            application.ExcelRow = null;
        }

        using var writeLock = await _writer.AcquireWriteLockAsync(cancellationToken).ConfigureAwait(false);
        for (var attempt = 0; attempt < 2; attempt++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var claimedRows = await _repository.GetClaimedExcelRowsAsync(targetKey, application.Id, cancellationToken).ConfigureAwait(false);
            var plannedRow = await _writer.ResolveTargetRowAsync(application, filePath, worksheetName, claimedRows, cancellationToken, effectiveSource).ConfigureAwait(false);
            if (!await _repository.TryReserveExcelRowAsync(application.Id, targetKey, plannedRow, cancellationToken).ConfigureAwait(false))
            {
                application.ExcelRow = null;
                continue;
            }

            application.ExcelTargetKey = targetKey;
            application.ExcelRow = plannedRow;
            try
            {
                var row = await _writer.WriteAsync(application, filePath, worksheetName, cancellationToken, writeLock, claimedRows, effectiveSource).ConfigureAwait(false);
                application.ExcelRow = row;
                application.ProcessingStatus = ProcessingStatus.Completed;
                application.ErrorMessage = string.Empty;
                await _repository.UpdateStatusAsync(application.Id, ProcessingStatus.Completed, excelRow: row, cancellationToken: cancellationToken).ConfigureAwait(false);
                return row;
            }
            catch (ExcelRowOccupiedException) when (attempt == 0)
            {
                await _repository.ClearExcelAssignmentAsync(application.Id, "Excel 目标行在写入前被占用，正在重新规划。", cancellationToken).ConfigureAwait(false);
                application.ExcelRow = null;
                application.ExcelTargetKey = string.Empty;
            }
        }

        await _repository.ClearExcelAssignmentAsync(application.Id, "Excel 目标行连续发生冲突，已保留在等待处理。", cancellationToken).ConfigureAwait(false);
        application.ExcelRow = null;
        application.ExcelTargetKey = string.Empty;
        throw new ExcelRowOccupiedException("Excel 目标行连续发生冲突，申请已保留在等待队列。");
    }

    public async Task<ExcelReconciliationResult> ReconcileAsync(
        string filePath,
        string worksheetName,
        CancellationToken cancellationToken = default,
        InvoiceSource source = InvoiceSource.Sinotrans)
    {
        var legacy = await _repository.GetLegacyExcelAssignmentsAsync(source, cancellationToken).ConfigureAwait(false);
        var targetKey = ExcelTargetIdentity.Create(filePath, worksheetName);
        var currentAssignments = await _repository.GetExcelAssignmentsAsync(targetKey, cancellationToken).ConfigureAwait(false);
        if (legacy.Count == 0 && currentAssignments.Count == 0)
            return new ExcelReconciliationResult(0, 0, 0, null);

        using var writeLock = await _writer.AcquireWriteLockAsync(cancellationToken).ConfigureAwait(false);
        // Re-read after acquiring the cross-process lock. Another process may
        // have completed the migration while this instance was waiting.
        legacy = await _repository.GetLegacyExcelAssignmentsAsync(source, cancellationToken).ConfigureAwait(false);
        currentAssignments = await _repository.GetExcelAssignmentsAsync(targetKey, cancellationToken).ConfigureAwait(false);
        if (legacy.Count == 0 && currentAssignments.Count == 0)
            return new ExcelReconciliationResult(0, 0, 0, null);

        var snapshot = await _writer.ReadSnapshotAsync(filePath, worksheetName, cancellationToken, source).ConfigureAwait(false);
        var invalidCurrent = currentAssignments
            .Where(application => application.ProcessingStatus != ProcessingStatus.Completed
                || application.ExcelRow is not int row
                || snapshot.FindRow(row)?.Matches(application) != true)
            .ToDictionary(application => application.Id);
        var work = legacy
            .Concat(invalidCurrent.Values)
            .GroupBy(application => application.Id)
            .Select(group => group.First())
            .OrderBy(application => application.Id)
            .ToArray();
        if (work.Length == 0)
            return new ExcelReconciliationResult(0, 0, 0, null);

        // Nothing is changed in either store until both backup artifacts have
        // been created successfully.
        var backupDirectory = await CreateBackupAsync(filePath, cancellationToken).ConfigureAwait(false);

        var claimedRows = new HashSet<int>(await _repository.GetClaimedExcelRowsAsync(targetKey, cancellationToken: cancellationToken).ConfigureAwait(false));
        var nextAppendRow = Math.Max(snapshot.LastUsedRow + 1, 2);
        var mappedExisting = 0;
        var written = 0;
        var pending = 0;

        foreach (var application in work)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var candidateClaims = new HashSet<int>(claimedRows);
            if (string.Equals(application.ExcelTargetKey, targetKey, StringComparison.Ordinal)
                && application.ExcelRow is int ownedRow)
                candidateClaims.Remove(ownedRow);

            var selected = SelectCandidateRow(application, snapshot, candidateClaims, ref nextAppendRow, allowReservedHint: string.Equals(application.ExcelTargetKey, targetKey, StringComparison.Ordinal));
            if (selected is null)
            {
                await _repository.ClearExcelAssignmentAsync(application.Id, "没有找到可安全接管的 Excel 行，请检查登记表后重试。", cancellationToken).ConfigureAwait(false);
                application.ExcelRow = null;
                application.ExcelTargetKey = string.Empty;
                application.ProcessingStatus = ProcessingStatus.PendingExcel;
                application.ErrorMessage = "没有找到可安全接管的 Excel 行，请检查登记表后重试。";
                pending++;
                continue;
            }

            var row = selected.Value;
            if (!await _repository.TryReserveExcelRowAsync(application.Id, targetKey, row, cancellationToken).ConfigureAwait(false))
            {
                // A reservation can lose a race with another process even
                // after the local lock and snapshot. Re-plan once from the
                // current ownership set without touching the workbook.
                claimedRows = new HashSet<int>(await _repository.GetClaimedExcelRowsAsync(targetKey, cancellationToken: cancellationToken).ConfigureAwait(false));
                candidateClaims = new HashSet<int>(claimedRows);
                if (string.Equals(application.ExcelTargetKey, targetKey, StringComparison.Ordinal)
                    && application.ExcelRow is int retryOwnedRow)
                    candidateClaims.Remove(retryOwnedRow);
                row = SelectCandidateRow(application, snapshot, candidateClaims, ref nextAppendRow, allowReservedHint: string.Equals(application.ExcelTargetKey, targetKey, StringComparison.Ordinal)) ?? 0;
                if (row < 2 || !await _repository.TryReserveExcelRowAsync(application.Id, targetKey, row, cancellationToken).ConfigureAwait(false))
                {
                    await _repository.ClearExcelAssignmentAsync(application.Id, "Excel 行预留发生冲突，已保留在等待处理。", cancellationToken).ConfigureAwait(false);
                    application.ExcelRow = null;
                    application.ExcelTargetKey = string.Empty;
                    application.ProcessingStatus = ProcessingStatus.PendingExcel;
                    application.ErrorMessage = "Excel 行预留发生冲突，已保留在等待处理。";
                    pending++;
                    continue;
                }
            }

            var otherClaimedRows = new HashSet<int>(claimedRows);
            if (string.Equals(application.ExcelTargetKey, targetKey, StringComparison.Ordinal)
                && application.ExcelRow is int currentOwnedRow)
                otherClaimedRows.Remove(currentOwnedRow);
            claimedRows.Add(row);
            application.ExcelRow = row;
            application.ExcelTargetKey = targetKey;
            application.ProcessingStatus = ProcessingStatus.PendingExcel;
            application.ErrorMessage = string.Empty;
            try
            {
                var existing = snapshot.FindRow(row);
                if (existing?.Matches(application) == true)
                {
                    await _repository.UpdateStatusAsync(application.Id, ProcessingStatus.Completed, excelRow: row, cancellationToken: cancellationToken).ConfigureAwait(false);
                    application.ProcessingStatus = ProcessingStatus.Completed;
                    mappedExisting++;
                    continue;
                }

                var outputRow = await _writer.WriteAsync(application, filePath, worksheetName, cancellationToken, writeLock, otherClaimedRows, source).ConfigureAwait(false);
                application.ExcelRow = outputRow;
                application.ProcessingStatus = ProcessingStatus.Completed;
                await _repository.UpdateStatusAsync(application.Id, ProcessingStatus.Completed, excelRow: outputRow, cancellationToken: cancellationToken).ConfigureAwait(false);
                written++;
            }
            catch (ExcelRowOccupiedException ex)
            {
                await _repository.ClearExcelAssignmentAsync(application.Id, ex.Message, cancellationToken).ConfigureAwait(false);
                application.ExcelRow = null;
                application.ExcelTargetKey = string.Empty;
                application.ProcessingStatus = ProcessingStatus.PendingExcel;
                application.ErrorMessage = ex.Message;
                pending++;
            }
            catch (IOException ex)
            {
                // Keep the reservation so a retry can safely reuse the same
                // planned row after the workbook is unlocked.
                await _repository.UpdateStatusAsync(application.Id, ProcessingStatus.PendingExcel, ex.Message, excelRow: row, cancellationToken: cancellationToken).ConfigureAwait(false);
                application.ProcessingStatus = ProcessingStatus.PendingExcel;
                application.ErrorMessage = ex.Message;
                pending++;
            }
            catch (Exception ex)
            {
                await _repository.UpdateStatusAsync(application.Id, ProcessingStatus.PendingExcel, ex.Message, excelRow: row, cancellationToken: cancellationToken).ConfigureAwait(false);
                application.ProcessingStatus = ProcessingStatus.PendingExcel;
                application.ErrorMessage = ex.Message;
                pending++;
            }
        }

        return new ExcelReconciliationResult(mappedExisting, written, pending, backupDirectory);
    }

    private static int? SelectCandidateRow(
        InvoiceApplication application,
        ExcelWorkbookSnapshot snapshot,
        IReadOnlySet<int> claimedRows,
        ref int nextAppendRow,
        bool allowReservedHint = false)
    {
        if (application.ExcelRow is int preferred && preferred >= 2
            && !claimedRows.Contains(preferred)
            && (snapshot.FindRow(preferred)?.Matches(application) == true
                || (allowReservedHint && snapshot.FindRow(preferred) is null)))
            return preferred;

        var matching = snapshot.FindMatchingRows(application, claimedRows);
        if (matching.Count > 0) return matching[0];

        while (claimedRows.Contains(nextAppendRow)) nextAppendRow++;
        return nextAppendRow++;
    }

    private async Task<string> CreateBackupAsync(string filePath, CancellationToken cancellationToken)
    {
        if (!File.Exists(filePath)) throw new FileNotFoundException("Excel 登记表不存在，无法执行历史核对。", filePath);

        var stamp = DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture);
        var backupDirectory = Path.Combine(_dataDirectory, "backups", stamp);
        if (Directory.Exists(backupDirectory))
            backupDirectory = Path.Combine(_dataDirectory, "backups", $"{stamp}-{DateTime.Now:fff}");
        Directory.CreateDirectory(backupDirectory);

        var workbookBackup = Path.Combine(backupDirectory, Path.GetFileName(filePath));
        var databaseBackup = Path.Combine(backupDirectory, "invoice-mail.db");
        await Task.Run(() => File.Copy(filePath, workbookBackup, false), cancellationToken).ConfigureAwait(false);
        await _repository.BackupToAsync(databaseBackup, cancellationToken).ConfigureAwait(false);
        return backupDirectory;
    }
}
