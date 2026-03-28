using System.Collections.ObjectModel;
using System.Data;
using System.Windows.Documents;
using CommunityToolkit.Mvvm.ComponentModel;
using ReportGenerator.App.Models;
using ReportGenerator.Core.Abstractions;
using ReportGenerator.Core.Models;

namespace ReportGenerator.App.ViewModels;

public sealed partial class MainWindowViewModel : ObservableObject
{
    private readonly ITemplateStorageService _templateStorageService;

    [ObservableProperty]
    private string _templateName = "New Report Template";

    [ObservableProperty]
    private ReportPageSize _selectedPageSize = ReportPageSize.A4;

    [ObservableProperty]
    private ReportOrientation _selectedOrientation = ReportOrientation.Portrait;

    [ObservableProperty]
    private double _marginTopMm = 18;

    [ObservableProperty]
    private double _marginRightMm = 14;

    [ObservableProperty]
    private double _marginBottomMm = 18;

    [ObservableProperty]
    private double _marginLeftMm = 14;

    [ObservableProperty]
    private int _rowsPerPage = 30;

    [ObservableProperty]
    private bool _headerOnlyOnFirstPage;

    [ObservableProperty]
    private double _detailHeaderFontSize = 10;

    [ObservableProperty]
    private double _detailContentFontSize = 10;

    [ObservableProperty]
    private double _contentRowSpacing;

    [ObservableProperty]
    private int _groupEveryRows;

    [ObservableProperty]
    private double _groupSpacingRows = 1;

    [ObservableProperty]
    private bool _isEasyMode;

    [ObservableProperty]
    private ImportMappingMode _selectedMappingMode = ImportMappingMode.HeaderName;

    [ObservableProperty]
    private bool _headerRowEnabled = true;

    [ObservableProperty]
    private string _excelFilePath = "No Excel file loaded.";

    [ObservableProperty]
    private string _sheetName = "-";

    [ObservableProperty]
    private int _rowCount;

    [ObservableProperty]
    private int _columnCount;

    [ObservableProperty]
    private DataView? _sampleRowsView;

    [ObservableProperty]
    private IDocumentPaginatorSource? _previewDocument;

    [ObservableProperty]
    private string _statusMessage = "Ready.";

    [ObservableProperty]
    private bool _isBusy;

    [ObservableProperty]
    private bool _isBusyProgressIndeterminate = true;

    [ObservableProperty]
    private double _busyProgressValue;

    [ObservableProperty]
    private string _currentTemplatePath = "Unsaved template";

    [ObservableProperty]
    private string _currentReportSummary = "Preview has not been generated yet.";

    [ObservableProperty]
    private int _selectedMainTabIndex;

    [ObservableProperty]
    private UpdateSettings? _updateSettings;

    [ObservableProperty]
    private HeaderRowEditorViewModel? _selectedHeaderRow;

    [ObservableProperty]
    private ReportBlockEditorViewModel? _selectedFooterBlock;

    [ObservableProperty]
    private DetailColumnEditorViewModel? _selectedDetailColumn;

    public MainWindowViewModel(ITemplateStorageService templateStorageService)
    {
        _templateStorageService = templateStorageService;
        PageSizes = Enum.GetValues<ReportPageSize>();
        Orientations = Enum.GetValues<ReportOrientation>();
        MappingModes = Enum.GetValues<ImportMappingMode>();
        Alignments = Enum.GetValues<ReportTextAlignment>();
        BlockTypes = Enum.GetValues<ReportBlockType>();

        ResetTemplate();
    }

    public ObservableCollection<HeaderRowEditorViewModel> HeaderRows { get; } = [];

    public ObservableCollection<ReportBlockEditorViewModel> FooterBlocks { get; } = [];

    public ObservableCollection<DetailColumnEditorViewModel> DetailColumns { get; } = [];

    public ImportedWorkbook? ImportedWorkbook { get; private set; }

    public PagedReport? CurrentReport { get; private set; }

    public IReadOnlyList<ReportPageSize> PageSizes { get; }

    public IReadOnlyList<ReportOrientation> Orientations { get; }

    public IReadOnlyList<ImportMappingMode> MappingModes { get; }

    public IReadOnlyList<ReportTextAlignment> Alignments { get; }

    public IReadOnlyList<ReportBlockType> BlockTypes { get; }

    public IDocumentPaginatorSource? PreviewDocumentForFullMode => IsEasyMode ? null : PreviewDocument;

    public IDocumentPaginatorSource? PreviewDocumentForEasyMode => IsEasyMode ? PreviewDocument : null;

    public void ResetTemplate()
    {
        ApplyTemplate(_templateStorageService.CreateDefaultTemplate());
        CurrentTemplatePath = "Unsaved template";
    }

    partial void OnIsEasyModeChanged(bool value)
    {
        SelectedMainTabIndex = value ? 2 : 0;
        OnPropertyChanged(nameof(PreviewDocumentForFullMode));
        OnPropertyChanged(nameof(PreviewDocumentForEasyMode));
    }

    partial void OnPreviewDocumentChanged(IDocumentPaginatorSource? value)
    {
        OnPropertyChanged(nameof(PreviewDocumentForFullMode));
        OnPropertyChanged(nameof(PreviewDocumentForEasyMode));
    }

    public void ApplyTemplate(ReportTemplate template)
    {
        TemplateName = template.Name;
        SelectedPageSize = template.PageSettings.PageSize;
        SelectedOrientation = template.PageSettings.Orientation;
        MarginTopMm = template.PageSettings.MarginTopMm;
        MarginRightMm = template.PageSettings.MarginRightMm;
        MarginBottomMm = template.PageSettings.MarginBottomMm;
        MarginLeftMm = template.PageSettings.MarginLeftMm;
        RowsPerPage = template.PageSettings.RowsPerPage;
        HeaderOnlyOnFirstPage = template.PageSettings.HeaderOnlyOnFirstPage;
        DetailHeaderFontSize = template.DetailTable.HeaderFontSize;
        DetailContentFontSize = template.DetailTable.ContentFontSize;
        ContentRowSpacing = template.DetailTable.ContentRowSpacing;
        GroupEveryRows = template.DetailTable.GroupEveryRows;
        GroupSpacingRows = template.DetailTable.GroupSpacingRows;
        SelectedMappingMode = template.ImportSettings.MappingMode;
        HeaderRowEnabled = template.ImportSettings.HeaderRowEnabled;

        HeaderRows.Clear();
        foreach (var row in template.HeaderBlocks
                     .GroupBy(block => Math.Max(1, block.Row))
                     .OrderBy(group => group.Key))
        {
            HeaderRows.Add(HeaderRowEditorViewModel.FromBlocks(row));
        }
        RenumberHeaderRows();

        DetailColumns.Clear();
        foreach (var column in template.DetailTable.Columns)
        {
            DetailColumns.Add(DetailColumnEditorViewModel.FromModel(column));
        }

        FooterBlocks.Clear();
        foreach (var block in template.FooterBlocks)
        {
            FooterBlocks.Add(ReportBlockEditorViewModel.FromModel(block));
        }
    }

    public ReportTemplate BuildTemplate()
    {
        return new ReportTemplate
        {
            Name = TemplateName,
            PageSettings = new PageSettings
            {
                PageSize = SelectedPageSize,
                Orientation = SelectedOrientation,
                MarginTopMm = MarginTopMm,
                MarginRightMm = MarginRightMm,
                MarginBottomMm = MarginBottomMm,
                MarginLeftMm = MarginLeftMm,
                RowsPerPage = RowsPerPage,
                HeaderOnlyOnFirstPage = HeaderOnlyOnFirstPage
            },
            ImportSettings = new ImportSettings
            {
                MappingMode = SelectedMappingMode,
                HeaderRowEnabled = HeaderRowEnabled,
                FirstSheetOnly = true
            },
            HeaderBlocks = new Collection<ReportBlock>(HeaderRows.SelectMany(row => row.ToBlocks()).ToList()),
            DetailTable = new DetailTableDefinition
            {
                Columns = new Collection<DetailColumnDefinition>(DetailColumns.Select(column => column.ToModel()).ToList()),
                RepeatHeaderOnEveryPage = true,
                HeaderFontSize = DetailHeaderFontSize,
                ContentFontSize = DetailContentFontSize,
                ContentRowSpacing = ContentRowSpacing,
                GroupEveryRows = GroupEveryRows,
                GroupSpacingRows = GroupSpacingRows
            },
            FooterBlocks = new Collection<ReportBlock>(FooterBlocks.Select(block => block.ToModel()).ToList())
        };
    }

    public void SetImportedWorkbook(ImportedWorkbook workbook)
    {
        ImportedWorkbook = workbook;
        ExcelFilePath = workbook.SourcePath;
        SheetName = workbook.SheetName;
        RowCount = workbook.Rows.Count;
        ColumnCount = workbook.Headers.Count;
        SampleRowsView = CreateSampleView(workbook);
    }

    public void ClearImportedWorkbook()
    {
        ImportedWorkbook = null;
        ExcelFilePath = "No Excel file loaded.";
        SheetName = "-";
        RowCount = 0;
        ColumnCount = 0;
        SampleRowsView = null;
    }

    public void SetPreview(PagedReport report, IDocumentPaginatorSource previewDocument)
    {
        CurrentReport = report;
        PreviewDocument = previewDocument;
        CurrentReportSummary = $"{report.Pages.Count} pages generated from {report.TotalRows} rows.";
    }

    public void ClearPreview()
    {
        CurrentReport = null;
        PreviewDocument = null;
        CurrentReportSummary = "Preview has not been generated yet.";
    }

    public void AddHeaderRow()
    {
        var row = HeaderRowEditorViewModel.CreateEmpty(HeaderRows.Count + 1);
        HeaderRows.Add(row);
        RenumberHeaderRows();
        SelectedHeaderRow = row;
    }

    public void MoveSelectedHeaderRowUp()
    {
        if (SelectedHeaderRow is null)
        {
            return;
        }

        var index = HeaderRows.IndexOf(SelectedHeaderRow);
        if (index <= 0)
        {
            return;
        }

        HeaderRows.Move(index, index - 1);
        RenumberHeaderRows();
    }

    public void MoveSelectedHeaderRowDown()
    {
        if (SelectedHeaderRow is null)
        {
            return;
        }

        var index = HeaderRows.IndexOf(SelectedHeaderRow);
        if (index < 0 || index >= HeaderRows.Count - 1)
        {
            return;
        }

        HeaderRows.Move(index, index + 1);
        RenumberHeaderRows();
    }

    public void RemoveSelectedHeaderRow()
    {
        if (SelectedHeaderRow is null)
        {
            return;
        }

        var removedIndex = HeaderRows.IndexOf(SelectedHeaderRow);
        if (removedIndex < 0)
        {
            return;
        }

        HeaderRows.RemoveAt(removedIndex);
        RenumberHeaderRows();

        if (HeaderRows.Count == 0)
        {
            SelectedHeaderRow = null;
            return;
        }

        SelectedHeaderRow = HeaderRows[Math.Min(removedIndex, HeaderRows.Count - 1)];
    }

    public void AddFooterBlock()
    {
        FooterBlocks.Add(new ReportBlockEditorViewModel
        {
            Row = FooterBlocks.Count == 0 ? 1 : FooterBlocks.Max(block => block.Row) + 1
        });
    }

    public void RemoveSelectedFooterBlock()
    {
        if (SelectedFooterBlock is not null)
        {
            FooterBlocks.Remove(SelectedFooterBlock);
        }
    }

    public void AddDetailColumn()
    {
        DetailColumns.Add(new DetailColumnEditorViewModel());
    }

    public void EnableEasyMode()
    {
        IsEasyMode = true;
    }

    public void DisableEasyMode()
    {
        IsEasyMode = false;
    }

    public void RemoveSelectedDetailColumn()
    {
        if (SelectedDetailColumn is not null)
        {
            DetailColumns.Remove(SelectedDetailColumn);
        }
    }

    private void RenumberHeaderRows()
    {
        for (var index = 0; index < HeaderRows.Count; index++)
        {
            HeaderRows[index].RowNumber = index + 1;
        }
    }

    private static DataView CreateSampleView(ImportedWorkbook workbook)
    {
        var table = new DataTable();
        foreach (var header in workbook.Headers)
        {
            table.Columns.Add(header);
        }

        foreach (var row in workbook.Rows.Take(100))
        {
            var values = new object[workbook.Headers.Count];
            for (var index = 0; index < workbook.Headers.Count; index++)
            {
                values[index] = row.GetCell(index) ?? string.Empty;
            }

            table.Rows.Add(values);
        }

        return table.DefaultView;
    }
}
