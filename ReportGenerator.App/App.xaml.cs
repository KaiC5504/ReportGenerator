using System.Windows;
using ReportGenerator.App.Services;
using ReportGenerator.App.ViewModels;
using ReportGenerator.Core.Abstractions;
using ReportGenerator.Core.Services;

namespace ReportGenerator.App;

public partial class App : Application
{
    private FileLogService? _logService;

    protected override async void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        _logService = new FileLogService();
        _logService.LogInfo("Application starting.");

        DispatcherUnhandledException += (_, args) =>
        {
            _logService.LogError("Unhandled dispatcher exception.", args.Exception);
            MessageBox.Show(
                "An unexpected error occurred. Details were written to the local app log.",
                "Report Generator",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            args.Handled = true;
        };

        ITemplateStorageService templateStorageService = new TemplateStorageService();
        var validationService = new TemplateValidationService();
        var viewModel = new MainWindowViewModel(templateStorageService);
        var updateService = new UpdateService(_logService, new UpdateSettingsStore());
        var sessionStateStore = new SessionStateStore();
        viewModel.UpdateSettings = await updateService.GetSettingsAsync();

        var window = new MainWindow(
            viewModel,
            templateStorageService,
            new ExcelImportService(),
            new ReportBuilder(validationService),
            new PdfExportService(),
            new DocxExportService(),
            new ReportPreviewService(),
            updateService,
            sessionStateStore,
            _logService);

        MainWindow = window;
        window.Show();
    }
}
