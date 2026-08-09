using Microsoft.Win32;
using System.Windows;

namespace InvoiceMailAssistant.App;

public partial class MainWindow : Window
{
    private readonly MainViewModel _viewModel = new();

    public MainWindow()
    {
        InitializeComponent();
        DataContext = _viewModel;
        CredentialInput.PasswordChanged += (_, _) => _viewModel.Password = CredentialInput.Password;
        Loaded += async (_, _) => await _viewModel.InitializeAsync();
        Closed += (_, _) => _viewModel.Dispose();
    }

    private void ChooseExcel_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
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
}
