using Microsoft.Win32;
using System.ComponentModel;
using System.IO;
using System.Windows.Controls;
using System.Windows;

namespace InvoiceMailAssistant.App;

public partial class MainWindow : Window
{
    private readonly MainViewModel _viewModel = new();
    private readonly System.Windows.Forms.NotifyIcon _notifyIcon;
    private readonly List<ColumnLayoutBinding> _columnLayouts = [];
    private UiLayoutStore? _layoutStore;
    private bool _layoutReady;
    private bool _allowExit;
    private bool _shutdownStarted;
    private Task? _initializeTask;

    private sealed record ColumnLayoutBinding(DataGridColumn Column, string Key, double DefaultWidth, double MinimumWidth, double MaximumWidth);

    public MainWindow()
    {
        InitializeComponent();
        DataContext = _viewModel;
        RegisterColumnLayouts();
        _notifyIcon = CreateNotifyIcon();
        if (System.Windows.Application.Current is App app) app.RegisterMainWindow(this);
        CredentialInput.PasswordChanged += (_, _) => _viewModel.Password = CredentialInput.Password;
        Loaded += (_, _) => _initializeTask ??= InitializeAsync();
        Closing += MainWindow_Closing;
        Closed += MainWindow_Closed;
        _viewModel.NewApplicationsFound += ViewModel_NewApplicationsFound;
    }

    private void RegisterColumnLayouts()
    {
        RegisterGrid(OverviewGrid, "overview", ["applyTime", "processedTime", "source", "company", "creditCode", "amount", "status", "excelRow"], [145, 145, 90, 260, 190, 90, 110, 80], [48, 48, 60, 140, 140, 48, 48, 48], [800, 800, 240, 800, 800, 800, 800, 800]);
        RegisterGrid(RecordsGrid, "records", ["applyTime", "processedTime", "source", "company", "creditCode", "amount", "email", "status", "excelRow"], [145, 145, 90, 260, 190, 90, 220, 110, 80], [48, 48, 60, 140, 140, 48, 160, 48, 48], [800, 800, 240, 800, 800, 800, 800, 800, 800]);
        RegisterGrid(PendingGrid, "pending", ["applyTime", "source", "company", "creditCode", "email", "amount", "status", "excelRow", "error"], [145, 90, 260, 190, 220, 90, 110, 80, 360], [48, 60, 140, 140, 160, 48, 48, 48, 200], [800, 240, 800, 800, 800, 800, 800, 800, 800]);
    }

    private void RegisterGrid(DataGrid grid, string page, IReadOnlyList<string> names, IReadOnlyList<double> defaults, IReadOnlyList<double> minimums, IReadOnlyList<double> maximums)
    {
        for (var index = 0; index < grid.Columns.Count; index++)
        {
            var column = grid.Columns[index];
            var binding = new ColumnLayoutBinding(column, $"{page}.{names[index]}", defaults[index], minimums[index], maximums[index]);
            _columnLayouts.Add(binding);
            DependencyPropertyDescriptor.FromProperty(DataGridColumn.WidthProperty, typeof(DataGridColumn))?.AddValueChanged(column, ColumnWidthChanged);
        }
    }

    private void ApplyColumnLayouts()
    {
        if (_layoutStore is null) return;
        foreach (var binding in _columnLayouts)
        {
            var width = _layoutStore.GetWidth(binding.Key, binding.DefaultWidth, binding.MinimumWidth, binding.MaximumWidth);
            binding.Column.Width = new DataGridLength(width, DataGridLengthUnitType.Pixel);
        }
    }

    private void ColumnWidthChanged(object? sender, EventArgs e)
    {
        if (!_layoutReady || _layoutStore is null || sender is not DataGridColumn column) return;
        var binding = _columnLayouts.FirstOrDefault(item => ReferenceEquals(item.Column, column));
        if (binding is null || !double.IsFinite(column.ActualWidth) || column.ActualWidth <= 0) return;

        var width = Math.Clamp(column.ActualWidth, binding.MinimumWidth, binding.MaximumWidth);
        if (Math.Abs(width - column.ActualWidth) > 0.5)
            column.Width = new DataGridLength(width, DataGridLengthUnitType.Pixel);
        _layoutStore.SetWidthDebounced(binding.Key, width, binding.MinimumWidth, binding.MaximumWidth);
    }

    private void ChooseExcel_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Title = "选择开票登记表",
            Filter = "Excel 工作簿 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = false,
            FileName = _viewModel.ExcelPath
        };

        if (dialog.ShowDialog(this) == true)
            _viewModel.ExcelPath = dialog.FileName;
    }

    private System.Windows.Forms.NotifyIcon CreateNotifyIcon()
    {
        var exitItem = new System.Windows.Forms.ToolStripMenuItem("退出");
        exitItem.Click += (_, _) =>
        {
            _allowExit = true;
            _notifyIcon.Visible = false;
            Close();
        };
        var showItem = new System.Windows.Forms.ToolStripMenuItem("显示窗口");
        showItem.Click += (_, _) => ShowFromTray();
        var menu = new System.Windows.Forms.ContextMenuStrip();
        menu.Items.Add(showItem);
        menu.Items.Add(new System.Windows.Forms.ToolStripSeparator());
        menu.Items.Add(exitItem);

        var notifyIcon = new System.Windows.Forms.NotifyIcon
        {
            Icon = System.Drawing.SystemIcons.Application,
            Text = "开票邮件助手",
            Visible = true,
            ContextMenuStrip = menu
        };
        notifyIcon.DoubleClick += (_, _) => ShowFromTray();
        return notifyIcon;
    }

    private void ShowFromTray()
    {
        Show();
        if (WindowState == WindowState.Minimized) WindowState = WindowState.Normal;
        Activate();
        Topmost = true;
        Topmost = false;
        Focus();
    }

    private void MainWindow_Closing(object? sender, CancelEventArgs e)
    {
        if (!_allowExit)
        {
            e.Cancel = true;
            Hide();
            _notifyIcon.ShowBalloonTip(1500, "开票邮件助手", "程序仍在后台监听，可从系统托盘打开。", System.Windows.Forms.ToolTipIcon.Info);
            return;
        }

        if (_shutdownStarted) return;
        e.Cancel = true;
        _shutdownStarted = true;
        _notifyIcon.Visible = false;
        _ = ShutdownAsync();
    }

    private async Task InitializeAsync()
    {
        try
        {
            var dataDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "InvoiceMailAssistant");
            _layoutStore = await UiLayoutStore.LoadAsync(Path.Combine(dataDirectory, "ui-layout.json"));
            ApplyColumnLayouts();
            _layoutReady = true;
            await _viewModel.InitializeAsync();
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"初始化失败：{ex.Message}", "开票邮件助手", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async Task ShutdownAsync()
    {
        try
        {
            if (_initializeTask is not null) await _initializeTask;
            await _viewModel.StopAsync();
            if (_layoutStore is not null) await _layoutStore.FlushAsync();
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"退出前保存或关闭后台任务失败：{ex.Message}", "开票邮件助手", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        finally
        {
            _allowExit = true;
            Close();
        }
    }

    private void ViewModel_NewApplicationsFound(IReadOnlyList<InvoiceSourceCount> counts)
        => _notifyIcon.ShowBalloonTip(2500, "开票邮件助手", SourceNotificationFormatter.Format(counts), System.Windows.Forms.ToolTipIcon.Info);

    private void MainWindow_Closed(object? sender, EventArgs e)
    {
        _viewModel.NewApplicationsFound -= ViewModel_NewApplicationsFound;
        _layoutStore?.Dispose();
        _notifyIcon.Visible = false;
        _notifyIcon.Dispose();
    }
}
