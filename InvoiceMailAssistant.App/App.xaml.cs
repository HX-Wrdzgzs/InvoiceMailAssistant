using System.Threading;
using System.Windows;

namespace InvoiceMailAssistant.App;

public partial class App : Application
{
    private Mutex? _singleInstanceMutex;

    protected override void OnStartup(StartupEventArgs e)
    {
        _singleInstanceMutex = new Mutex(true, "Local\\InvoiceMailAssistant.SingleInstance", out var createdNew);
        if (!createdNew)
        {
            Shutdown();
            return;
        }

        base.OnStartup(e);
    }

    protected override void OnExit(ExitEventArgs e)
    {
        try { _singleInstanceMutex?.ReleaseMutex(); }
        catch (ApplicationException) { }
        _singleInstanceMutex?.Dispose();
        base.OnExit(e);
    }
}
