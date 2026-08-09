using Microsoft.Win32;
using System.ComponentModel;
using System.Windows;

namespace InvoiceMailAssistant.App;

public partial class MainWindow : Window
{
    private readonly MainViewModel _viewModel = new();
    private readonly System.Windows.Forms.NotifyIcon _notifyIcon;
    private bool _allowExit;

    public MainWindow()
    {
        InitializeComponent();
        DataContext = _viewModel;
        _notifyIcon = CreateNotifyIcon();
        if (System.Windows.Application.Current is App app) app.RegisterMainWindow(this);
        CredentialInput.PasswordChanged += (_, _) => _viewModel.Password = CredentialInput.Password;
        Loaded += async (_, _) =>
        {
            try { await _viewModel.InitializeAsync(); }
            catch (Exception ex) { System.Windows.MessageBox.Show($"初始化失败：{ex.Message}", "开票邮件助手", MessageBoxButton.OK, MessageBoxImage.Error); }
        };
        Closing += MainWindow_Closing;
        Closed += MainWindow_Closed;
        _viewModel.NewApplicationsFound += ViewModel_NewApplicationsFound;
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
        if (_allowExit) return;
        e.Cancel = true;
        Hide();
        _notifyIcon.ShowBalloonTip(1500, "开票邮件助手", "程序仍在后台监听，可从系统托盘打开。", System.Windows.Forms.ToolTipIcon.Info);
    }

    private void ViewModel_NewApplicationsFound(int count)
        => _notifyIcon.ShowBalloonTip(2500, "开票邮件助手", $"发现 {count} 条新申请，已写入或进入等待队列。", System.Windows.Forms.ToolTipIcon.Info);

    private void MainWindow_Closed(object? sender, EventArgs e)
    {
        _viewModel.NewApplicationsFound -= ViewModel_NewApplicationsFound;
        _notifyIcon.Visible = false;
        _notifyIcon.Dispose();
        _viewModel.Dispose();
    }
}
