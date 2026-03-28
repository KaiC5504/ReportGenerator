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

        var firstPageRowsPerPage = CalculateEffectiveRowsPerPage(template, pageNumber: 1);
        var subsequentPageRowsPerPage = CalculateEffectiveRowsPerPage(template, pageNumber: 2);
        var effectiveRowsPerPage = Math.Min(firstPageRowsPerPage, subsequentPageRowsPerPage);
        var pagedRows = BuildPagedRows(workbook, template, columns, firstPageRowsPerPage, subsequentPageRowsPerPage);
        var totalPages = Math.Max(1, pagedRows.Count);
        var pages = new List<ReportPage>(totalPages);

        for (var pageIndex = 0; pageIndex < totalPages; pageIndex++)
        {
            var pageNumber = pageIndex + 1;
            var headerBlocks = BuildHeaderBlocks(template, workbook, pageNumber, totalPages);
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
            template.DetailTable.ContentRowSpacing,
            workbook.Rows.Count);
    }

    private static IReadOnlyList<IReadOnlyList<ReportRow>> BuildPagedRows(
        ImportedWorkbook workbook,
        ReportTemplate template,
        IReadOnlyList<(DetailColumnDefinition Definition, int Index)> columns,
        int firstPageRowsPerPage,
        int subsequentPageRowsPerPage)
    {
        var pages = new List<IReadOnlyList<ReportRow>>();
        var currentPageRows = new List<ReportRow>();
        var groupEveryRows = template.DetailTable.GroupEveryRows;
        var groupSpacingRows = template.DetailTable.GroupSpacingRows;
        var currentGroupItemCount = 0;
        var currentPageHeightFactor = 0d;
        var currentPageNumber = 1;
        var currentPageRowsPerPage = GetRowsPerPageForPage(currentPageNumber, firstPageRowsPerPage, subsequentPageRowsPerPage);

        for (var rowIndex = 0; rowIndex < workbook.Rows.Count; rowIndex++)
        {
            if (currentPageHeightFactor + 1 > currentPageRowsPerPage)
            {
                pages.Add(currentPageRows.ToArray());
                currentPageRows = [];
                currentGroupItemCount = 0;
                currentPageHeightFactor = 0;
                currentPageNumber++;
                currentPageRowsPerPage = GetRowsPerPageForPage(currentPageNumber, firstPageRowsPerPage, subsequentPageRowsPerPage);
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
            currentPageHeightFactor += 1;

            var hasMoreRows = rowIndex < workbook.Rows.Count - 1;
            var isGroupBoundary = groupEveryRows > 0 && currentGroupItemCount == groupEveryRows;
            if (!hasMoreRows || !isGroupBoundary)
            {
                continue;
            }

            if (groupSpacingRows == 0)
            {
                currentGroupItemCount = 0;
            }
            else if (currentPageHeightFactor + groupSpacingRows + 1 <= currentPageRowsPerPage)
            {
                currentPageRows.Add(ReportRow.CreateSpacer(groupSpacingRows));
                currentPageHeightFactor += groupSpacingRows;
                currentGroupItemCount = 0;
            }
            else if (currentPageRows.Count > 0)
            {
                pages.Add(currentPageRows.ToArray());
                currentPageRows = [];
                currentGroupItemCount = 0;
                currentPageHeightFactor = 0;
                currentPageNumber++;
                currentPageRowsPerPage = GetRowsPerPageForPage(currentPageNumber, firstPageRowsPerPage, subsequentPageRowsPerPage);
            }
        }

        if (currentPageRows.Count > 0 || pages.Count == 0)
        {
            pages.Add(currentPageRows.ToArray());
        }

        return pages;
    }

    private static int CalculateEffectiveRowsPerPage(ReportTemplate template, int pageNumber)
    {
        var (_, pageHeightMm) = PageMeasurementHelper.GetPageDimensionsMillimeters(template.PageSettings);
        var pageHeightDip = PageMeasurementHelper.MillimetersToDip(pageHeightMm);
        var printableHeightDip = pageHeightDip
            - PageMeasurementHelper.MillimetersToDip(template.PageSettings.MarginTopMm)
            - PageMeasurementHelper.MillimetersToDip(template.PageSettings.MarginBottomMm);

        var headerHeightDip = CalculateHeaderBlocksHeightDip(GetVisibleHeaderBlocks(template, pageNumber));
        var footerHeightDip = CalculateFlatBlocksHeightDip(template.FooterBlocks);
        var availableTableHeightDip = Math.Max(
            PageMeasurementHelper.MinimumTableHeightDip,
            printableHeightDip - headerHeightDip - footerHeightDip - PageMeasurementHelper.SectionSpacingDip);

        var headerRowHeightDip = PageMeasurementHelper.CalculateTableCellMinHeightDip(template.DetailTable.HeaderFontSize);
        var contentRowHeightDip = PageMeasurementHelper.CalculateContentRowHeightDip(
            template.DetailTable.ContentFontSize,
            template.DetailTable.ContentRowSpacing);
        var maxRowsThatFit = Math.Max(
            1,
            (int)Math.Floor(Math.Max(0, availableTableHeightDip - headerRowHeightDip) / contentRowHeightDip));

        return Math.Max(1, Math.Min(template.PageSettings.RowsPerPage, maxRowsThatFit));
    }

    private static int GetRowsPerPageForPage(int pageNumber, int firstPageRowsPerPage, int subsequentPageRowsPerPage)
    {
        return pageNumber == 1 ? firstPageRowsPerPage : subsequentPageRowsPerPage;
    }

    private static double CalculateHeaderBlocksHeightDip(IEnumerable<ReportBlock> blocks)
    {
        return blocks
            .GroupBy(block => block.Row)
            .Sum(group => group.Max(block => PageMeasurementHelper.CalculateTextLineHeightDip(block.FontSize)));
    }

    private static double CalculateFlatBlocksHeightDip(IEnumerable<ReportBlock> blocks)
    {
        return blocks.Sum(block => PageMeasurementHelper.CalculateTextLineHeightDip(block.FontSize));
    }

    private IReadOnlyList<ReportBlockContent> BuildHeaderBlocks(
        ReportTemplate template,
        ImportedWorkbook workbook,
        int pageNumber,
        int totalPages)
    {
        return BuildBlocks(
                GetVisibleHeaderBlocks(template, pageNumber),
                workbook,
                template.ImportSettings,
                pageNumber,
                totalPages)
            .ToArray();
    }

    private static IReadOnlyList<ReportBlock> GetVisibleHeaderBlocks(ReportTemplate template, int pageNumber)
    {
        if (template.PageSettings.HeaderOnlyOnFirstPage && pageNumber > 1)
        {
            return Array.Empty<ReportBlock>();
        }

        return template.HeaderBlocks
            .GroupBy(block => block.Row)
            .OrderBy(group => group.Key)
            .Where(group => pageNumber == 1 || !group.First().OnlyOnFirstPage)
            .SelectMany(group => group.OrderBy(block => GetAlignmentOrder(block.Alignment)))
            .ToArray();
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
                Type = block.Type,
                Text = ResolveBlockText(block, workbook, importSettings, firstRow, pageNumber, totalPages),
                TemplateText = block.Text,
                Alignment = block.Alignment,
                Row = block.Row,
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

    private static int GetAlignmentOrder(ReportTextAlignment alignment)
    {
        return alignment switch
        {
            ReportTextAlignment.Left => 0,
            ReportTextAlignment.Center => 1,
            ReportTextAlignment.Right => 2,
            _ => 3
        };
    }
}
