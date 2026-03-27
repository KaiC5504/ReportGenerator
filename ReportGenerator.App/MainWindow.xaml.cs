using System.IO;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using ReportGenerator.App.Models;
using ReportGenerator.App.Services;
using ReportGenerator.App.ViewModels;
using ReportGenerator.App.Views;
using ReportGenerator.Core.Abstractions;
using ReportGenerator.Core.Services;

namespace ReportGenerator.App;

public partial class MainWindow : Window
{
    private readonly MainWindowViewModel _viewModel;
    private readonly ITemplateStorageService _templateStorageService;
    private readonly IExcelImportService _excelImportService;
    private readonly IReportBuilder _reportBuilder;
    private readonly IPdfExportService _pdfExportService;
    private readonly IDocxExportService _docxExportService;
    private readonly ReportPreviewService _reportPreviewService;
    private readonly UpdateService _updateService;
    private readonly SessionStateStore _sessionStateStore;
    private readonly FileLogService _logService;

    public MainWindow(
        MainWindowViewModel viewModel,
        ITemplateStorageService templateStorageService,
        IExcelImportService excelImportService,
        IReportBuilder reportBuilder,
        IPdfExportService pdfExportService,
        IDocxExportService docxExportService,
        ReportPreviewService reportPreviewService,
        UpdateService updateService,
        SessionStateStore sessionStateStore,
        FileLogService logService)
    {
        InitializeComponent();

        _viewModel = viewModel;
        _templateStorageService = templateStorageService;
        _excelImportService = excelImportService;
        _reportBuilder = reportBuilder;
        _pdfExportService = pdfExportService;
        _docxExportService = docxExportService;
        _reportPreviewService = reportPreviewService;
        _updateService = updateService;
        _sessionStateStore = sessionStateStore;
        _logService = logService;

        DataContext = _viewModel;
        Loaded += OnLoadedAsync;
        Closing += OnClosing;
    }

    private async void OnLoadedAsync(object sender, RoutedEventArgs e)
    {
        try
        {
            await RestoreSessionAsync();

            if (_viewModel.ImportedWorkbook is not null)
            {
                await GeneratePreviewAsync();
            }
        }
        catch (Exception exception)
        {
            _logService.LogError("Failed to restore the previous session.", exception);
            _viewModel.StatusMessage = "Ready.";
        }

        try
        {
            var settings = await _updateService.GetSettingsAsync();
            _viewModel.UpdateSettings = settings;
            _ = AutoCheckForUpdatesAsync(settings);
        }
        catch (Exception exception)
        {
            _logService.LogError("Failed to initialize update settings.", exception);
        }
    }

    private void OnClosing(object? sender, CancelEventArgs e)
    {
        try
        {
            _sessionStateStore.Save(new SessionState
            {
                CurrentTemplatePath = _viewModel.CurrentTemplatePath,
                ExcelFilePath = _viewModel.ImportedWorkbook?.SourcePath,
                IsEasyMode = _viewModel.IsEasyMode,
                Template = _viewModel.BuildTemplate()
            });
        }
        catch (Exception exception)
        {
            _logService.LogError("Failed to save session state.", exception);
        }
    }

    private async Task AutoCheckForUpdatesAsync(Models.UpdateSettings settings)
    {
        try
        {
            await Task.Delay(TimeSpan.FromSeconds(settings.StartupDelaySeconds));
            await CheckForUpdatesAsync(false);
        }
        catch (Exception exception)
        {
            _logService.LogError("Automatic update check failed.", exception);
        }
    }

    private async Task RestoreSessionAsync()
    {
        var sessionState = _sessionStateStore.Load();
        if (sessionState?.Template is null)
        {
            return;
        }

        _viewModel.ApplyTemplate(sessionState.Template);
        _viewModel.CurrentTemplatePath = sessionState.CurrentTemplatePath;
        _viewModel.IsEasyMode = sessionState.IsEasyMode;
        _viewModel.ClearPreview();

        if (string.IsNullOrWhiteSpace(sessionState.ExcelFilePath))
        {
            _viewModel.ClearImportedWorkbook();
            _viewModel.StatusMessage = "Previous template restored.";
            return;
        }

        if (!File.Exists(sessionState.ExcelFilePath))
        {
            _viewModel.ClearImportedWorkbook();
            _viewModel.StatusMessage = "Previous template restored. The Excel file could not be found.";
            _logService.LogInfo($"Saved Excel file was not found during session restore: {sessionState.ExcelFilePath}");
            return;
        }

        var workbook = await _excelImportService.ImportAsync(sessionState.ExcelFilePath, _viewModel.BuildTemplate().ImportSettings);
        _viewModel.SetImportedWorkbook(workbook);
        _viewModel.StatusMessage = $"Restored previous session from {Path.GetFileName(sessionState.ExcelFilePath)}.";
    }

    private async void NewTemplateButton_OnClick(object sender, RoutedEventArgs e)
    {
        _viewModel.ResetTemplate();
        _viewModel.ClearPreview();
        _viewModel.StatusMessage = "Template reset to the default layout.";
        await Task.CompletedTask;
    }

    private async void OpenTemplateButton_OnClick(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Template Files (*.json)|*.json|All Files (*.*)|*.*",
            Title = "Open Report Template"
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        await RunBusyAsync("Loading template...", async () =>
        {
            var template = await _templateStorageService.LoadAsync(dialog.FileName);
            _viewModel.ApplyTemplate(template);
            _viewModel.CurrentTemplatePath = dialog.FileName;
            _viewModel.ClearPreview();
            _viewModel.StatusMessage = "Template loaded.";
        });
    }

    private async void SaveTemplateButton_OnClick(object sender, RoutedEventArgs e)
    {
        var dialog = new SaveFileDialog
        {
            Filter = "Template Files (*.json)|*.json|All Files (*.*)|*.*",
            FileName = $"{SanitizeFileName(_viewModel.TemplateName)}.json",
            Title = "Save Report Template"
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        await RunBusyAsync("Saving template...", async () =>
        {
            await _templateStorageService.SaveAsync(_viewModel.BuildTemplate(), dialog.FileName);
            _viewModel.CurrentTemplatePath = dialog.FileName;
            _viewModel.StatusMessage = "Template saved.";
        });
    }

    private async void OpenExcelButton_OnClick(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel Files (*.xlsx)|*.xlsx",
            Title = "Open Excel Workbook"
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        await RunBusyAsync("Loading Excel workbook...", async () =>
        {
            var workbook = await _excelImportService.ImportAsync(dialog.FileName, _viewModel.BuildTemplate().ImportSettings);
            _viewModel.SetImportedWorkbook(workbook);
            _viewModel.ClearPreview();
            _viewModel.StatusMessage = $"Loaded {workbook.Rows.Count} rows from {workbook.SheetName}.";
        });
    }

    private async void GeneratePreviewButton_OnClick(object sender, RoutedEventArgs e)
    {
        await GeneratePreviewAsync();
    }

    private async Task GeneratePreviewAsync()
    {
        if (_viewModel.ImportedWorkbook is null)
        {
            MessageBox.Show(this, "Load an Excel workbook before generating the report.", "Report Generator", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        await RunBusyAsync("Generating preview...", async () =>
        {
            var template = _viewModel.BuildTemplate();
            _viewModel.IsBusyProgressIndeterminate = true;
            _viewModel.BusyProgressValue = 0;
            _viewModel.StatusMessage = "Building report pages...";
            var report = await Task.Run(() => _reportBuilder.Build(template, _viewModel.ImportedWorkbook));
            var progress = new Progress<PreviewGenerationProgress>(value =>
            {
                _viewModel.IsBusyProgressIndeterminate = false;
                _viewModel.BusyProgressValue = value.Percentage;
                _viewModel.StatusMessage = $"Rendering preview page {value.CurrentPage} of {value.TotalPages}...";
            });

            var previewDocument = await _reportPreviewService.CreateDocumentAsync(report, progress);
            _viewModel.SetPreview(report, previewDocument);
            _viewModel.StatusMessage = $"Preview generated with {report.Pages.Count} pages.";
        });
    }

    private async void ExportPdfButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (!await EnsurePreviewAsync())
        {
            return;
        }

        var dialog = new SaveFileDialog
        {
            Filter = "PDF Files (*.pdf)|*.pdf",
            FileName = $"{SanitizeFileName(_viewModel.TemplateName)}.pdf",
            Title = "Export PDF"
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        await RunBusyAsync("Exporting PDF...", async () =>
        {
            await _pdfExportService.ExportAsync(_viewModel.CurrentReport!, dialog.FileName);
            _viewModel.StatusMessage = $"PDF exported to {dialog.FileName}.";
        });
    }

    private async void ExportWordButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (!await EnsurePreviewAsync())
        {
            return;
        }

        var dialog = new SaveFileDialog
        {
            Filter = "Word Files (*.docx)|*.docx",
            FileName = $"{SanitizeFileName(_viewModel.TemplateName)}.docx",
            Title = "Export Word"
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        await RunBusyAsync("Exporting Word document...", async () =>
        {
            await _docxExportService.ExportAsync(_viewModel.CurrentReport!, dialog.FileName);
            _viewModel.StatusMessage = $"Word document exported to {dialog.FileName}.";
        });
    }

    private async void PrintButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (!await EnsurePreviewAsync())
        {
            return;
        }

        var dialog = new PrintDialog();
        if (dialog.ShowDialog() == true)
        {
            dialog.PrintDocument(_viewModel.PreviewDocument!.DocumentPaginator, _viewModel.TemplateName);
            _viewModel.StatusMessage = "Report sent to the printer.";
        }
    }

    private async void CheckUpdatesButton_OnClick(object sender, RoutedEventArgs e)
    {
        await CheckForUpdatesAsync(true);
    }

    private void EnableEasyModeButton_OnClick(object sender, RoutedEventArgs e)
    {
        _viewModel.EnableEasyMode();
        _viewModel.StatusMessage = "Easy mode enabled.";
    }

    private void DisableEasyModeButton_OnClick(object sender, RoutedEventArgs e)
    {
        _viewModel.DisableEasyMode();
        _viewModel.StatusMessage = "Full mode restored.";
    }

    private async Task CheckForUpdatesAsync(bool userInitiated)
    {
        var result = await _updateService.CheckForUpdatesAsync(userInitiated);
        _viewModel.StatusMessage = result.Message;

        if (result.IsUpdateAvailable)
        {
            var window = new UpdatePromptWindow(result.AvailableVersion, result.ReleaseNotes) { Owner = this };
            if (window.ShowDialog() == true)
            {
                await RunBusyAsync("Downloading update...", async () =>
                {
                    await _updateService.DownloadAndApplyUpdateAsync(
                        result,
                        progress => Dispatcher.Invoke(() => _viewModel.StatusMessage = $"Downloading update... {progress}%"));
                });

                Close();
                return;
            }
        }
        else if (userInitiated)
        {
            MessageBox.Show(this, result.Message, "Report Generator", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    private void AddHeaderBlockButton_OnClick(object sender, RoutedEventArgs e)
    {
        _viewModel.AddHeaderBlock();
    }

    private void RemoveHeaderBlockButton_OnClick(object sender, RoutedEventArgs e)
    {
        _viewModel.RemoveSelectedHeaderBlock();
    }

    private void AddDetailColumnButton_OnClick(object sender, RoutedEventArgs e)
    {
        _viewModel.AddDetailColumn();
    }

    private void RemoveDetailColumnButton_OnClick(object sender, RoutedEventArgs e)
    {
        _viewModel.RemoveSelectedDetailColumn();
    }

    private void AddFooterBlockButton_OnClick(object sender, RoutedEventArgs e)
    {
        _viewModel.AddFooterBlock();
    }

    private void RemoveFooterBlockButton_OnClick(object sender, RoutedEventArgs e)
    {
        _viewModel.RemoveSelectedFooterBlock();
    }

    private async Task<bool> EnsurePreviewAsync()
    {
        if (_viewModel.CurrentReport is not null && _viewModel.PreviewDocument is not null)
        {
            return true;
        }

        await GeneratePreviewAsync();
        return _viewModel.CurrentReport is not null && _viewModel.PreviewDocument is not null;
    }

    private async Task RunBusyAsync(string status, Func<Task> action)
    {
        try
        {
            _viewModel.IsBusy = true;
            _viewModel.IsBusyProgressIndeterminate = true;
            _viewModel.BusyProgressValue = 0;
            _viewModel.StatusMessage = status;
            await action();
        }
        catch (ReportValidationException validationException)
        {
            _logService.LogError("Template validation failed.", validationException);
            var errors = string.Join(Environment.NewLine, validationException.ValidationResult.Issues.Select(issue => $"- {issue.Message}"));
            MessageBox.Show(this, errors, "Template Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
            _viewModel.StatusMessage = "Template validation failed.";
        }
        catch (Exception exception)
        {
            _logService.LogError(status, exception);
            MessageBox.Show(
                this,
                "The operation could not be completed. Details were written to the local app log.",
                "Report Generator",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            _viewModel.StatusMessage = "Operation failed.";
        }
        finally
        {
            _viewModel.BusyProgressValue = 0;
            _viewModel.IsBusyProgressIndeterminate = true;
            _viewModel.IsBusy = false;
        }
    }

    private static string SanitizeFileName(string value)
    {
        var invalidCharacters = Path.GetInvalidFileNameChars();
        return string.Concat(value.Select(character => invalidCharacters.Contains(character) ? '_' : character));
    }
}
