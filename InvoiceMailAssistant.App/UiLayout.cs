using System.IO;
using System.Text.Json;

namespace InvoiceMailAssistant.App;

public sealed class UiLayoutStore : IDisposable
{
    private sealed class LayoutDocument
    {
        public Dictionary<string, double> Columns { get; set; } = new(StringComparer.Ordinal);
    }

    private readonly string _filePath;
    private readonly object _sync = new();
    private readonly SemaphoreSlim _saveGate = new(1, 1);
    private Dictionary<string, double> _columns;
    private CancellationTokenSource? _debounceCancellation;
    private Task? _pendingSave;
    private bool _disposed;

    private UiLayoutStore(string filePath, Dictionary<string, double> columns)
    {
        _filePath = filePath;
        _columns = columns;
    }

    public static async Task<UiLayoutStore> LoadAsync(string filePath, CancellationToken cancellationToken = default)
    {
        var columns = new Dictionary<string, double>(StringComparer.Ordinal);
        try
        {
            if (File.Exists(filePath))
            {
                await using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, useAsync: true);
                var document = await JsonSerializer.DeserializeAsync<LayoutDocument>(stream, cancellationToken: cancellationToken).ConfigureAwait(false);
                if (document?.Columns is not null)
                {
                    foreach (var pair in document.Columns)
                    {
                        if (double.IsFinite(pair.Value)) columns[pair.Key] = pair.Value;
                    }
                }
            }
        }
        catch (JsonException)
        {
            // A damaged layout file only affects optional presentation state;
            // use defaults and let the next resize write a fresh document.
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // The UI must still start when the optional file is temporarily
            // unavailable (for example, while a backup tool is scanning it).
        }

        return new UiLayoutStore(filePath, columns);
    }

    public double GetWidth(string key, double defaultWidth, double minimum, double maximum)
    {
        var fallback = Clamp(defaultWidth, minimum, maximum);
        lock (_sync)
        {
            return _columns.TryGetValue(key, out var value) && double.IsFinite(value)
                ? Clamp(value, minimum, maximum)
                : fallback;
        }
    }

    public void SetWidthDebounced(string key, double width, double minimum, double maximum)
    {
        var value = Clamp(width, minimum, maximum);
        CancellationToken token;
        lock (_sync)
        {
            if (_disposed) return;
            _columns[key] = value;
            _debounceCancellation?.Cancel();
            _debounceCancellation?.Dispose();
            _debounceCancellation = new CancellationTokenSource();
            token = _debounceCancellation.Token;
            _pendingSave = SaveAfterDelayAsync(token);
        }
    }

    public async Task FlushAsync(CancellationToken cancellationToken = default)
    {
        Task? pending;
        lock (_sync) pending = _pendingSave;
        if (pending is not null)
        {
            try { await pending.ConfigureAwait(false); }
            catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested) { }
        }

        cancellationToken.ThrowIfCancellationRequested();
        await SaveNowAsync(cancellationToken).ConfigureAwait(false);
    }

    private async Task SaveAfterDelayAsync(CancellationToken cancellationToken)
    {
        try
        {
            await Task.Delay(250, cancellationToken).ConfigureAwait(false);
            await SaveNowAsync(cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // A later resize or an explicit flush can retry the optional file.
        }
    }

    private async Task SaveNowAsync(CancellationToken cancellationToken)
    {
        Dictionary<string, double> snapshot;
        lock (_sync) snapshot = new Dictionary<string, double>(_columns, StringComparer.Ordinal);

        var directory = Path.GetDirectoryName(_filePath);
        if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory);
        var temporaryPath = $"{_filePath}.{Guid.NewGuid():N}.tmp";
        await _saveGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            var document = new LayoutDocument { Columns = snapshot };
            await using (var stream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None, 4096, useAsync: true))
            {
                await JsonSerializer.SerializeAsync(stream, document, new JsonSerializerOptions { WriteIndented = true }, cancellationToken).ConfigureAwait(false);
                await stream.FlushAsync(cancellationToken).ConfigureAwait(false);
                stream.Flush(true);
            }

            if (File.Exists(_filePath)) File.Replace(temporaryPath, _filePath, null, true);
            else File.Move(temporaryPath, _filePath);
        }
        finally
        {
            _saveGate.Release();
            if (File.Exists(temporaryPath))
            {
                try { File.Delete(temporaryPath); }
                catch (IOException) { }
            }
        }
    }

    private static double Clamp(double value, double minimum, double maximum)
    {
        if (minimum > maximum) (minimum, maximum) = (maximum, minimum);
        if (!double.IsFinite(value)) return Math.Clamp(120d, minimum, maximum);
        return Math.Clamp(value, minimum, maximum);
    }

    public void Dispose()
    {
        lock (_sync)
        {
            if (_disposed) return;
            _disposed = true;
            _debounceCancellation?.Cancel();
        }

        try { FlushAsync().GetAwaiter().GetResult(); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
        catch (OperationCanceledException) { }
        _debounceCancellation?.Dispose();
        _saveGate.Dispose();
    }
}
