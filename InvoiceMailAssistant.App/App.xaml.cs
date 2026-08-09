using System.Threading;
using System.Windows;
using System.Windows.Threading;

namespace InvoiceMailAssistant.App;

public partial class App : System.Windows.Application
{
    private Mutex? _singleInstanceMutex;
    private EventWaitHandle? _activationEvent;
    private RegisteredWaitHandle? _activationRegistration;
    private MainWindow? _mainWindow;

    protected override void OnStartup(StartupEventArgs e)
    {
        _activationEvent = new EventWaitHandle(false, EventResetMode.AutoReset, "Local\\InvoiceMailAssistant.Activate");
        _singleInstanceMutex = new Mutex(true, "Local\\InvoiceMailAssistant.SingleInstance", out var createdNew);
        if (!createdNew)
        {
            _activationEvent.Set();
            _activationEvent.Dispose();
            _activationEvent = null;
            Shutdown();
            return;
        }

        base.OnStartup(e);
    }

    protected override void OnExit(ExitEventArgs e)
    {
        try { _singleInstanceMutex?.ReleaseMutex(); }
        catch (ApplicationException) { }
        _activationRegistration?.Unregister(null);
        _activationEvent?.Dispose();
        _singleInstanceMutex?.Dispose();
        base.OnExit(e);
    }

    internal void RegisterMainWindow(MainWindow window)
    {
        _mainWindow = window;
        if (_activationEvent is null) return;
        _activationRegistration = ThreadPool.RegisterWaitForSingleObject(
            _activationEvent,
            (_, _) => Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(ActivateMainWindow)),
            null,
            Timeout.Infinite,
            false);
    }

    private void ActivateMainWindow()
    {
        if (_mainWindow is null) return;
        _mainWindow.Show();
        if (_mainWindow.WindowState == WindowState.Minimized) _mainWindow.WindowState = WindowState.Normal;
        _mainWindow.Activate();
        _mainWindow.Topmost = true;
        _mainWindow.Topmost = false;
        _mainWindow.Focus();
    }
}
