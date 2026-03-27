using ReportGenerator.Core.Abstractions;
using ReportGenerator.Core.Models;
using ReportGenerator.Core.Utilities;

namespace ReportGenerator.Core.Services;

public sealed class ReportBuilder : IReportBuilder
{
    private readonly ITemplateValidationService _templateValidationService;

    public ReportBuilder(ITemplateValidationService templateValidationService)
    {
        _templateValidationService = templateValidationService;
    }

    public PagedReport Build(ReportTemplate template, ImportedWorkbook workbook)
    {
        ArgumentNullException.ThrowIfNull(template);
        ArgumentNullException.ThrowIfNull(workbook);

        template = TemplateStorageService.NormalizeTemplate(template);
        var validationResult = _templateValidationService.ValidateTemplate(template, workbook);
        if (!validationResult.IsValid)
        {
            throw new ReportValidationException(validationResult);
        }

        var columns = template.DetailTable.Columns
            .Select(column => (Definition: column, Index: _templateValidationService.ResolveColumnIndex(workbook, template.ImportSettings, column.Source)))
            .ToArray();

        var effectiveRowsPerPage = CalculateEffectiveRowsPerPage(template);
        var pagedRows = BuildPagedRows(workbook, template, columns, effectiveRowsPerPage);
        var totalPages = Math.Max(1, pagedRows.Count);
        var pages = new List<ReportPage>(totalPages);

        for (var pageIndex = 0; pageIndex < totalPages; pageIndex++)
        {
            var pageNumber = pageIndex + 1;
            var headerBlocks = pageNumber == 1 || !template.PageSettings.HeaderOnlyOnFirstPage
                ? BuildBlocks(template.HeaderBlocks, workbook, template.ImportSettings, pageNumber, totalPages)
                : Array.Empty<ReportBlockContent>();
            var pageRows = pagedRows[pageIndex];

            pages.Add(ReportPage.Create(
                pageNumber,
                totalPages,
                headerBlocks,
                pageRows,
                BuildBlocks(template.FooterBlocks, workbook, template.ImportSettings, pageNumber, totalPages)));
        }

        return PagedReport.Create(
            template.Name,
            new PageSettings
            {
                PageSize = template.PageSettings.PageSize,
                Orientation = template.PageSettings.Orientation,
                MarginTopMm = template.PageSettings.MarginTopMm,
                MarginRightMm = template.PageSettings.MarginRightMm,
                MarginBottomMm = template.PageSettings.MarginBottomMm,
                MarginLeftMm = template.PageSettings.MarginLeftMm,
                RowsPerPage = effectiveRowsPerPage,
                HeaderOnlyOnFirstPage = template.PageSettings.HeaderOnlyOnFirstPage
            },
            workbook.FileName,
            workbook.SheetName,
            DateTimeOffset.UtcNow,
            columns.Select(column => new ReportColumnLayout
            {
                HeaderText = column.Definition.HeaderText,
                Source = column.Definition.Source,
                WidthWeight = column.Definition.WidthWeight,
                Alignment = column.Definition.Alignment
            }),
            pages,
            template.DetailTable.HeaderFontSize,
            template.DetailTable.ContentFontSize,
            workbook.Rows.Count);
    }

    private static IReadOnlyList<IReadOnlyList<ReportRow>> BuildPagedRows(
        ImportedWorkbook workbook,
        ReportTemplate template,
        IReadOnlyList<(DetailColumnDefinition Definition, int Index)> columns,
        int rowsPerPage)
    {
        var pages = new List<IReadOnlyList<ReportRow>>();
        var currentPageRows = new List<ReportRow>();
        var groupEveryRows = template.DetailTable.GroupEveryRows;
        var currentGroupItemCount = 0;

        for (var rowIndex = 0; rowIndex < workbook.Rows.Count; rowIndex++)
        {
            if (currentPageRows.Count == rowsPerPage)
            {
                pages.Add(currentPageRows.ToArray());
                currentPageRows = [];
                currentGroupItemCount = 0;
            }

            var workbookRow = workbook.Rows[rowIndex];
            currentPageRows.Add(ReportRow.Create(
                workbookRow.RowNumber,
                columns.Select(column => new ReportCell
                {
                    Text = workbookRow.GetCell(column.Index) ?? string.Empty,
                    Alignment = column.Definition.Alignment
                })));
            currentGroupItemCount++;

            var hasMoreRows = rowIndex < workbook.Rows.Count - 1;
            var isGroupBoundary = groupEveryRows > 0 && currentGroupItemCount == groupEveryRows;
            if (!hasMoreRows || !isGroupBoundary)
            {
                continue;
            }

            var remainingSlots = rowsPerPage - currentPageRows.Count;
            if (remainingSlots >= 2)
            {
                currentPageRows.Add(ReportRow.CreateSpacer());
                currentGroupItemCount = 0;
            }
            else if (currentPageRows.Count > 0)
            {
                pages.Add(currentPageRows.ToArray());
                currentPageRows = [];
                currentGroupItemCount = 0;
            }
        }

        if (currentPageRows.Count > 0 || pages.Count == 0)
        {
            pages.Add(currentPageRows.ToArray());
        }

        return pages;
    }

    private static int CalculateEffectiveRowsPerPage(ReportTemplate template)
    {
        var (_, pageHeightMm) = PageMeasurementHelper.GetPageDimensionsMillimeters(template.PageSettings);
        var pageHeightDip = PageMeasurementHelper.MillimetersToDip(pageHeightMm);
        var printableHeightDip = pageHeightDip
            - PageMeasurementHelper.MillimetersToDip(template.PageSettings.MarginTopMm)
            - PageMeasurementHelper.MillimetersToDip(template.PageSettings.MarginBottomMm);

        var headerHeightDip = CalculateBlocksHeightDip(template.HeaderBlocks);
        var footerHeightDip = CalculateBlocksHeightDip(template.FooterBlocks);
        var availableTableHeightDip = Math.Max(
            PageMeasurementHelper.MinimumTableHeightDip,
            printableHeightDip - headerHeightDip - footerHeightDip - (PageMeasurementHelper.SectionSpacingDip * 2));

        var headerRowHeightDip = PageMeasurementHelper.CalculateTableCellMinHeightDip(template.DetailTable.HeaderFontSize);
        var contentRowHeightDip = PageMeasurementHelper.CalculateTableCellMinHeightDip(template.DetailTable.ContentFontSize);
        var maxRowsThatFit = Math.Max(
            1,
            (int)Math.Floor(Math.Max(0, availableTableHeightDip - headerRowHeightDip) / contentRowHeightDip));

        return Math.Max(1, Math.Min(template.PageSettings.RowsPerPage, maxRowsThatFit));
    }

    private static double CalculateBlocksHeightDip(IEnumerable<ReportBlock> blocks)
    {
        return blocks.Sum(block => PageMeasurementHelper.CalculateTextLineHeightDip(block.FontSize));
    }

    private IEnumerable<ReportBlockContent> BuildBlocks(
        IEnumerable<ReportBlock> blocks,
        ImportedWorkbook workbook,
        ImportSettings importSettings,
        int pageNumber,
        int totalPages)
    {
        var firstRow = workbook.FirstRow;

        foreach (var block in blocks)
        {
            yield return new ReportBlockContent
            {
                Text = ResolveBlockText(block, workbook, importSettings, firstRow, pageNumber, totalPages),
                Alignment = block.Alignment,
                FontSize = block.FontSize,
                IsBold = block.IsBold
            };
        }
    }

    private string ResolveBlockText(
        ReportBlock block,
        ImportedWorkbook workbook,
        ImportSettings importSettings,
        ImportedDataRow? firstRow,
        int pageNumber,
        int totalPages)
    {
        return block.Type switch
        {
            ReportBlockType.StaticText => block.Text,
            ReportBlockType.PageNumber => ResolvePageNumberText(block.Text, pageNumber, totalPages),
            ReportBlockType.MappedField => ResolveMappedFieldText(block, workbook, importSettings, firstRow),
            _ => string.Empty
        };
    }

    private string ResolveMappedFieldText(
        ReportBlock block,
        ImportedWorkbook workbook,
        ImportSettings importSettings,
        ImportedDataRow? firstRow)
    {
        if (firstRow is null)
        {
            return block.Text;
        }

        var index = _templateValidationService.ResolveColumnIndex(workbook, importSettings, block.Source);
        var value = firstRow.GetCell(index) ?? string.Empty;
        return $"{block.Text}{value}";
    }

    private static string ResolvePageNumberText(string template, int pageNumber, int totalPages)
    {
        var format = string.IsNullOrWhiteSpace(template) ? "Page {page} of {totalPages}" : template;
        return format
            .Replace("{page}", pageNumber.ToString(), StringComparison.Ordinal)
            .Replace("{totalPages}", totalPages.ToString(), StringComparison.Ordinal);
    }
}
