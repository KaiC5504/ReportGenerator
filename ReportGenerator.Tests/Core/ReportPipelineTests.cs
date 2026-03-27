using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using ReportGenerator.Core.Models;
using ReportGenerator.Core.Services;

namespace ReportGenerator.Tests.Core;

public sealed class ReportPipelineTests
{
    private readonly TemplateStorageService _templateStorageService = new();
    private readonly ExcelImportService _excelImportService = new();
    private readonly TemplateValidationService _validationService = new();
    private readonly PdfExportService _pdfExportService = new();
    private readonly DocxExportService _docxExportService = new();

    [Fact]
    public async Task TemplateStorage_RoundTripsTemplate()
    {
        var template = _templateStorageService.CreateDefaultTemplate();
        template.Name = "Warehouse Report";
        template.PageSettings.RowsPerPage = 42;
        template.PageSettings.HeaderOnlyOnFirstPage = true;
        template.DetailTable.HeaderFontSize = 14;
        template.DetailTable.ContentFontSize = 12;
        template.DetailTable.GroupEveryRows = 5;
        template.DetailTable.Columns[0].Source = "A";

        var filePath = CreateTempFile(".json");

        await _templateStorageService.SaveAsync(template, filePath);
        var loaded = await _templateStorageService.LoadAsync(filePath);

        Assert.Equal("Warehouse Report", loaded.Name);
        Assert.Equal(42, loaded.PageSettings.RowsPerPage);
        Assert.True(loaded.PageSettings.HeaderOnlyOnFirstPage);
        Assert.Equal(14, loaded.DetailTable.HeaderFontSize);
        Assert.Equal(12, loaded.DetailTable.ContentFontSize);
        Assert.Equal(5, loaded.DetailTable.GroupEveryRows);
        Assert.Equal("A", loaded.DetailTable.Columns[0].Source);
    }

    [Fact]
    public async Task ExcelImport_UsesHeaderNames_WhenHeaderRowEnabled()
    {
        var workbookPath = CreateWorkbook(new[]
        {
            new[] { "ItemId", "Quantity", "Quality", "Price" },
            new[] { "SKU-001", "4", "High", "10.50" },
            new[] { "SKU-002", "8", "Low", "15.75" }
        });

        var workbook = await _excelImportService.ImportAsync(workbookPath, new ImportSettings
        {
            HeaderRowEnabled = true,
            MappingMode = ImportMappingMode.HeaderName
        });

        Assert.Equal(["ItemId", "Quantity", "Quality", "Price"], workbook.Headers);
        Assert.Equal(2, workbook.Rows.Count);
        Assert.Equal("SKU-001", workbook.Rows[0].GetCell(0));
    }

    [Fact]
    public async Task ExcelImport_GeneratesColumnHeaders_WhenHeaderRowDisabled()
    {
        var workbookPath = CreateWorkbook(new[]
        {
            new[] { "SKU-001", "4", "High", "10.50" },
            new[] { "SKU-002", "8", "Low", "15.75" }
        });

        var workbook = await _excelImportService.ImportAsync(workbookPath, new ImportSettings
        {
            HeaderRowEnabled = false,
            MappingMode = ImportMappingMode.ColumnLetter
        });

        Assert.Equal(["Column A", "Column B", "Column C", "Column D"], workbook.Headers);
        Assert.Equal(2, workbook.Rows.Count);
    }

    [Fact]
    public async Task ReportBuilder_PaginatesByConfiguredRowsPerPage()
    {
        var workbookPath = CreateWorkbook(new[]
        {
            new[] { "ItemId", "Quantity", "Quality", "Price" },
            new[] { "SKU-001", "1", "A", "10" },
            new[] { "SKU-002", "2", "B", "20" },
            new[] { "SKU-003", "3", "C", "30" },
            new[] { "SKU-004", "4", "D", "40" },
            new[] { "SKU-005", "5", "E", "50" }
        });

        var workbook = await _excelImportService.ImportAsync(workbookPath, new ImportSettings());
        var template = BuildTemplate(rowsPerPage: 2, mappingMode: ImportMappingMode.HeaderName);
        var builder = new ReportBuilder(_validationService);

        var report = builder.Build(template, workbook);

        Assert.Equal(3, report.Pages.Count);
        Assert.Equal(2, report.Pages[0].Rows.Count);
        Assert.Equal(2, report.Pages[1].Rows.Count);
        Assert.Single(report.Pages[2].Rows);
        Assert.Equal("SKU-005", report.Pages[2].Rows[0].Cells[0].Text);
    }

    [Fact]
    public async Task ReportBuilder_ShowsHeaderOnlyOnFirstPage_WhenConfigured()
    {
        var workbookPath = CreateWorkbook(new[]
        {
            new[] { "ItemId", "Quantity", "Quality", "Price" },
            new[] { "SKU-001", "1", "A", "10" },
            new[] { "SKU-002", "2", "B", "20" },
            new[] { "SKU-003", "3", "C", "30" },
            new[] { "SKU-004", "4", "D", "40" },
            new[] { "SKU-005", "5", "E", "50" }
        });

        var workbook = await _excelImportService.ImportAsync(workbookPath, new ImportSettings());
        var template = BuildTemplate(rowsPerPage: 2, mappingMode: ImportMappingMode.HeaderName, headerOnlyOnFirstPage: true);
        var builder = new ReportBuilder(_validationService);

        var report = builder.Build(template, workbook);

        Assert.NotEmpty(report.Pages[0].HeaderBlocks);
        Assert.Empty(report.Pages[1].HeaderBlocks);
        Assert.Empty(report.Pages[2].HeaderBlocks);
    }

    [Fact]
    public async Task ReportBuilder_InsertsSpacerRowsBetweenGroups()
    {
        var workbookPath = CreateWorkbook(new[]
        {
            new[] { "ItemId", "Quantity", "Quality", "Price" },
            new[] { "SKU-001", "1", "A", "10" },
            new[] { "SKU-002", "2", "B", "20" },
            new[] { "SKU-003", "3", "C", "30" },
            new[] { "SKU-004", "4", "D", "40" },
            new[] { "SKU-005", "5", "E", "50" },
            new[] { "SKU-006", "6", "F", "60" }
        });

        var workbook = await _excelImportService.ImportAsync(workbookPath, new ImportSettings());
        var template = BuildTemplate(rowsPerPage: 7, mappingMode: ImportMappingMode.HeaderName, groupEveryRows: 5);
        var builder = new ReportBuilder(_validationService);

        var report = builder.Build(template, workbook);

        Assert.Equal(7, report.Pages[0].Rows.Count);
        Assert.True(report.Pages[0].Rows[5].IsSpacer);
        Assert.Equal(10, report.DetailHeaderFontSize);
        Assert.Equal(12, report.DetailContentFontSize);
    }

    [Fact]
    public async Task ReportBuilder_ReducesRowsPerPageWhenFontSizesNeedMoreHeight()
    {
        var workbookPath = CreateWorkbook(new[]
        {
            new[] { "ItemId", "Quantity", "Quality", "Price" },
            new[] { "SKU-001", "1", "A", "10" },
            new[] { "SKU-002", "2", "B", "20" },
            new[] { "SKU-003", "3", "C", "30" },
            new[] { "SKU-004", "4", "D", "40" },
            new[] { "SKU-005", "5", "E", "50" },
            new[] { "SKU-006", "6", "F", "60" },
            new[] { "SKU-007", "7", "G", "70" },
            new[] { "SKU-008", "8", "H", "80" },
            new[] { "SKU-009", "9", "I", "90" },
            new[] { "SKU-010", "10", "J", "100" },
            new[] { "SKU-011", "11", "K", "110" },
            new[] { "SKU-012", "12", "L", "120" },
            new[] { "SKU-013", "13", "M", "130" },
            new[] { "SKU-014", "14", "N", "140" },
            new[] { "SKU-015", "15", "O", "150" },
            new[] { "SKU-016", "16", "P", "160" },
            new[] { "SKU-017", "17", "Q", "170" },
            new[] { "SKU-018", "18", "R", "180" },
            new[] { "SKU-019", "19", "S", "190" },
            new[] { "SKU-020", "20", "T", "200" },
            new[] { "SKU-021", "21", "U", "210" },
            new[] { "SKU-022", "22", "V", "220" },
            new[] { "SKU-023", "23", "W", "230" },
            new[] { "SKU-024", "24", "X", "240" }
        });

        var workbook = await _excelImportService.ImportAsync(workbookPath, new ImportSettings());
        var template = BuildTemplate(rowsPerPage: 30, mappingMode: ImportMappingMode.HeaderName, contentFontSize: 24, headerFontSize: 24);
        var builder = new ReportBuilder(_validationService);

        var report = builder.Build(template, workbook);

        Assert.True(report.PageSettings.RowsPerPage < 30);
        Assert.True(report.PageSettings.RowsPerPage >= 1);
        Assert.True(report.Pages.Count >= 2);
    }

    [Fact]
    public async Task ReportBuilder_DoesNotStartNextPageWithSpacerRow()
    {
        var workbookPath = CreateWorkbook(new[]
        {
            new[] { "ItemId", "Quantity", "Quality", "Price" },
            new[] { "SKU-001", "1", "A", "10" },
            new[] { "SKU-002", "2", "B", "20" },
            new[] { "SKU-003", "3", "C", "30" },
            new[] { "SKU-004", "4", "D", "40" },
            new[] { "SKU-005", "5", "E", "50" },
            new[] { "SKU-006", "6", "F", "60" },
            new[] { "SKU-007", "7", "G", "70" },
            new[] { "SKU-008", "8", "H", "80" },
            new[] { "SKU-009", "9", "I", "90" },
            new[] { "SKU-010", "10", "J", "100" }
        });

        var workbook = await _excelImportService.ImportAsync(workbookPath, new ImportSettings());
        var template = BuildTemplate(rowsPerPage: 5, mappingMode: ImportMappingMode.HeaderName, groupEveryRows: 5);
        var builder = new ReportBuilder(_validationService);

        var report = builder.Build(template, workbook);

        Assert.Equal(2, report.Pages.Count);
        Assert.All(report.Pages, page => Assert.DoesNotContain(page.Rows, row => row.IsSpacer));
    }

    [Fact]
    public async Task ReportBuilder_ResetsGroupCountOnEachNewPage()
    {
        var workbookPath = CreateWorkbook(new[]
        {
            new[] { "ItemId", "Quantity", "Quality", "Price" },
            new[] { "SKU-001", "1", "A", "10" },
            new[] { "SKU-002", "2", "B", "20" },
            new[] { "SKU-003", "3", "C", "30" },
            new[] { "SKU-004", "4", "D", "40" },
            new[] { "SKU-005", "5", "E", "50" },
            new[] { "SKU-006", "6", "F", "60" },
            new[] { "SKU-007", "7", "G", "70" },
            new[] { "SKU-008", "8", "H", "80" }
        });

        var workbook = await _excelImportService.ImportAsync(workbookPath, new ImportSettings());
        var template = BuildTemplate(rowsPerPage: 4, mappingMode: ImportMappingMode.HeaderName, groupEveryRows: 5);
        var builder = new ReportBuilder(_validationService);

        var report = builder.Build(template, workbook);

        Assert.Equal(2, report.Pages.Count);
        Assert.Equal(4, report.Pages[0].Rows.Count);
        Assert.Equal(4, report.Pages[1].Rows.Count);
        Assert.All(report.Pages[0].Rows, row => Assert.False(row.IsSpacer));
        Assert.All(report.Pages[1].Rows, row => Assert.False(row.IsSpacer));
        Assert.Equal("SKU-005", report.Pages[1].Rows[0].Cells[0].Text);
    }

    [Fact]
    public async Task PdfExport_CreatesNonEmptyFile()
    {
        var report = await CreateSampleReportAsync();
        var outputPath = CreateTempFile(".pdf");

        await _pdfExportService.ExportAsync(report, outputPath);

        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Exists);
        Assert.True(fileInfo.Length > 0);
    }

    [Fact]
    public async Task DocxExport_CreatesReadableWordDocument()
    {
        var report = await CreateSampleReportAsync();
        var outputPath = CreateTempFile(".docx");

        await _docxExportService.ExportAsync(report, outputPath);

        using var wordDocument = WordprocessingDocument.Open(outputPath, false);
        Assert.NotNull(wordDocument.MainDocumentPart);
        Assert.NotNull(wordDocument.MainDocumentPart!.Document);
        Assert.NotNull(wordDocument.MainDocumentPart.Document.Body);
    }

    [Fact]
    public async Task Validation_FailsWhenMappedColumnIsMissing()
    {
        var workbookPath = CreateWorkbook(new[]
        {
            new[] { "ItemId", "Quantity" },
            new[] { "SKU-001", "4" }
        });

        var workbook = await _excelImportService.ImportAsync(workbookPath, new ImportSettings());
        var template = BuildTemplate(rowsPerPage: 10, mappingMode: ImportMappingMode.HeaderName);
        template.DetailTable.Columns[3].Source = "Price";

        var result = _validationService.ValidateTemplate(template, workbook);

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, issue => issue.Field.Contains("detailTable.columns[3].source", StringComparison.Ordinal));
    }

    private async Task<PagedReport> CreateSampleReportAsync()
    {
        var workbookPath = CreateWorkbook(new[]
        {
            new[] { "ItemId", "Quantity", "Quality", "Price" },
            new[] { "SKU-001", "4", "High", "10.50" },
            new[] { "SKU-002", "2", "Mid", "8.25" },
            new[] { "SKU-003", "6", "Low", "12.80" }
        });

        var workbook = await _excelImportService.ImportAsync(workbookPath, new ImportSettings());
        var builder = new ReportBuilder(_validationService);
        return builder.Build(BuildTemplate(rowsPerPage: 2, mappingMode: ImportMappingMode.HeaderName), workbook);
    }

    private static ReportTemplate BuildTemplate(
        int rowsPerPage,
        ImportMappingMode mappingMode,
        bool headerOnlyOnFirstPage = false,
        int groupEveryRows = 0,
        double contentFontSize = 12,
        double headerFontSize = 10)
    {
        return new ReportTemplate
        {
            Name = "Test Report",
            PageSettings = new PageSettings
            {
                RowsPerPage = rowsPerPage,
                HeaderOnlyOnFirstPage = headerOnlyOnFirstPage
            },
            ImportSettings = new ImportSettings
            {
                HeaderRowEnabled = true,
                MappingMode = mappingMode,
                FirstSheetOnly = true
            },
            HeaderBlocks =
            {
                new ReportBlock
                {
                    Type = ReportBlockType.StaticText,
                    Text = "Inventory Report",
                    FontSize = 16,
                    IsBold = true
                }
            },
            DetailTable = new DetailTableDefinition
            {
                HeaderFontSize = headerFontSize,
                ContentFontSize = contentFontSize,
                GroupEveryRows = groupEveryRows,
                Columns =
                {
                    new DetailColumnDefinition { HeaderText = "Item Id", Source = mappingMode == ImportMappingMode.ColumnLetter ? "A" : "ItemId", WidthWeight = 1.5 },
                    new DetailColumnDefinition { HeaderText = "Quantity", Source = mappingMode == ImportMappingMode.ColumnLetter ? "B" : "Quantity", Alignment = ReportTextAlignment.Right },
                    new DetailColumnDefinition { HeaderText = "Quality", Source = mappingMode == ImportMappingMode.ColumnLetter ? "C" : "Quality" },
                    new DetailColumnDefinition { HeaderText = "Price", Source = mappingMode == ImportMappingMode.ColumnLetter ? "D" : "Price", Alignment = ReportTextAlignment.Right }
                }
            },
            FooterBlocks =
            {
                new ReportBlock
                {
                    Type = ReportBlockType.PageNumber,
                    Text = "Page {page} of {totalPages}",
                    Alignment = ReportTextAlignment.Right
                }
            }
        };
    }

    private static string CreateWorkbook(IReadOnlyList<string[]> rows)
    {
        var filePath = CreateTempFile(".xlsx");
        using var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet("Sheet1");

        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            for (var columnIndex = 0; columnIndex < rows[rowIndex].Length; columnIndex++)
            {
                worksheet.Cell(rowIndex + 1, columnIndex + 1).Value = rows[rowIndex][columnIndex];
            }
        }

        workbook.SaveAs(filePath);
        return filePath;
    }

    private static string CreateTempFile(string extension)
    {
        return Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}{extension}");
    }
}
