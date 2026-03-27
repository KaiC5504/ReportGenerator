using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Threading;
using ReportGenerator.App.Models;
using ReportGenerator.Core.Models;
using ReportGenerator.Core.Utilities;

namespace ReportGenerator.App.Services;

public sealed class ReportPreviewService
{
    private const int PreviewYieldInterval = 8;

    public FixedDocument CreateDocument(PagedReport report)
    {
        ArgumentNullException.ThrowIfNull(report);

        var (pageWidthMm, pageHeightMm) = PageMeasurementHelper.GetPageDimensionsMillimeters(report.PageSettings);
        var pageSize = new Size(
            PageMeasurementHelper.MillimetersToDip(pageWidthMm),
            PageMeasurementHelper.MillimetersToDip(pageHeightMm));

        var document = new FixedDocument
        {
            DocumentPaginator =
            {
                PageSize = pageSize
            }
        };

        foreach (var page in report.Pages)
        {
            var fixedPage = BuildPageVisual(report, page, pageSize);
            var pageContent = new PageContent();
            ((IAddChild)pageContent).AddChild(fixedPage);
            document.Pages.Add(pageContent);
        }

        return document;
    }

    public async Task<FixedDocument> CreateDocumentAsync(
        PagedReport report,
        IProgress<PreviewGenerationProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(report);

        var (pageWidthMm, pageHeightMm) = PageMeasurementHelper.GetPageDimensionsMillimeters(report.PageSettings);
        var pageSize = new Size(
            PageMeasurementHelper.MillimetersToDip(pageWidthMm),
            PageMeasurementHelper.MillimetersToDip(pageHeightMm));

        var document = new FixedDocument
        {
            DocumentPaginator =
            {
                PageSize = pageSize
            }
        };

        await Dispatcher.Yield(DispatcherPriority.Background);

        for (var pageIndex = 0; pageIndex < report.Pages.Count; pageIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var fixedPage = BuildPageVisual(report, report.Pages[pageIndex], pageSize);
            var pageContent = new PageContent();
            ((IAddChild)pageContent).AddChild(fixedPage);
            document.Pages.Add(pageContent);

            progress?.Report(new PreviewGenerationProgress(pageIndex + 1, report.Pages.Count));

            if ((pageIndex + 1) % PreviewYieldInterval == 0)
            {
                await Dispatcher.Yield(DispatcherPriority.Background);
            }
        }

        return document;
    }

    private static FixedPage BuildPageVisual(PagedReport report, ReportPage reportPage, Size pageSize)
    {
        var fixedPage = new FixedPage
        {
            Width = pageSize.Width,
            Height = pageSize.Height,
            Background = Brushes.White
        };

        var marginLeft = PageMeasurementHelper.MillimetersToDip(report.PageSettings.MarginLeftMm);
        var marginTop = PageMeasurementHelper.MillimetersToDip(report.PageSettings.MarginTopMm);
        var marginRight = PageMeasurementHelper.MillimetersToDip(report.PageSettings.MarginRightMm);
        var marginBottom = PageMeasurementHelper.MillimetersToDip(report.PageSettings.MarginBottomMm);
        var contentWidth = pageSize.Width - marginLeft - marginRight;

        var headerHeight = CalculateBlocksHeight(reportPage.HeaderBlocks);
        var footerHeight = CalculateBlocksHeight(reportPage.FooterBlocks);
        var footerTop = pageSize.Height - marginBottom - footerHeight;
        var tableTop = marginTop + headerHeight + PageMeasurementHelper.SectionSpacingDip;
        var tableHeight = Math.Max(PageMeasurementHelper.MinimumTableHeightDip, footerTop - tableTop - PageMeasurementHelper.SectionSpacingDip);
        var (headerRowHeight, contentRowHeight) = CalculateTableRowHeights(report, tableHeight);
        var columnWidths = CalculateColumnWidths(report.Columns, contentWidth);

        AddBlocks(fixedPage, reportPage.HeaderBlocks, marginLeft, marginTop, contentWidth);
        AddTable(fixedPage, report, reportPage, marginLeft, tableTop, headerRowHeight, contentRowHeight, columnWidths);
        AddBlocks(fixedPage, reportPage.FooterBlocks, marginLeft, footerTop, contentWidth);

        fixedPage.Measure(pageSize);
        fixedPage.Arrange(new Rect(new Point(0, 0), pageSize));
        fixedPage.UpdateLayout();
        return fixedPage;
    }

    private static void AddBlocks(FixedPage fixedPage, IReadOnlyList<ReportBlockContent> blocks, double left, double top, double width)
    {
        var currentY = top;
        foreach (var block in blocks)
        {
            var lineHeight = PageMeasurementHelper.CalculateTextLineHeightDip(block.FontSize);
            var textBlock = new TextBlock
            {
                Text = block.Text,
                Width = width,
                Height = lineHeight,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = block.FontSize,
                FontWeight = block.IsBold ? FontWeights.SemiBold : FontWeights.Normal,
                TextAlignment = ToTextAlignment(block.Alignment),
                TextTrimming = TextTrimming.CharacterEllipsis
            };

            FixedPage.SetLeft(textBlock, left);
            FixedPage.SetTop(textBlock, currentY);
            fixedPage.Children.Add(textBlock);
            currentY += lineHeight;
        }
    }

    private static void AddTable(
        FixedPage fixedPage,
        PagedReport report,
        ReportPage page,
        double left,
        double top,
        double headerRowHeight,
        double contentRowHeight,
        IReadOnlyList<double> columnWidths)
    {
        var currentX = left;
        for (var columnIndex = 0; columnIndex < report.Columns.Count; columnIndex++)
        {
            var column = report.Columns[columnIndex];
            var headerCell = CreateCell(
                column.HeaderText,
                column.Alignment,
                columnWidths[columnIndex],
                headerRowHeight,
                true,
                report.DetailHeaderFontSize,
                report.DetailContentFontSize);
            FixedPage.SetLeft(headerCell, currentX);
            FixedPage.SetTop(headerCell, top);
            fixedPage.Children.Add(headerCell);
            currentX += columnWidths[columnIndex];
        }

        for (var rowIndex = 0; rowIndex < page.Rows.Count; rowIndex++)
        {
            var y = top + headerRowHeight + (rowIndex * contentRowHeight);
            var row = page.Rows[rowIndex];
            if (row.IsSpacer)
            {
                continue;
            }

            currentX = left;

            for (var columnIndex = 0; columnIndex < row.Cells.Count; columnIndex++)
            {
                var cell = row.Cells[columnIndex];
                var cellVisual = CreateCell(
                    cell.Text,
                    cell.Alignment,
                    columnWidths[columnIndex],
                    contentRowHeight,
                    false,
                    report.DetailHeaderFontSize,
                    report.DetailContentFontSize);
                FixedPage.SetLeft(cellVisual, currentX);
                FixedPage.SetTop(cellVisual, y);
                fixedPage.Children.Add(cellVisual);
                currentX += columnWidths[columnIndex];
            }
        }
    }

    private static Border CreateCell(
        string text,
        ReportTextAlignment alignment,
        double width,
        double height,
        bool isHeader,
        double headerFontSize,
        double contentFontSize)
    {
        return new Border
        {
            Width = width,
            Height = height,
            BorderBrush = new SolidColorBrush(Color.FromRgb(210, 214, 220)),
            BorderThickness = new Thickness(0.7),
            Background = isHeader ? new SolidColorBrush(Color.FromRgb(245, 247, 250)) : Brushes.White,
            Child = new TextBlock
            {
                Text = text,
                Margin = new Thickness(4, 2, 4, 2),
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = isHeader ? headerFontSize : contentFontSize,
                FontWeight = isHeader ? FontWeights.SemiBold : FontWeights.Normal,
                TextAlignment = ToTextAlignment(alignment),
                TextTrimming = TextTrimming.CharacterEllipsis,
                VerticalAlignment = VerticalAlignment.Center
            }
        };
    }

    private static double CalculateBlocksHeight(IEnumerable<ReportBlockContent> blocks)
    {
        return blocks.Sum(block => PageMeasurementHelper.CalculateTextLineHeightDip(block.FontSize));
    }

    private static (double HeaderRowHeight, double ContentRowHeight) CalculateTableRowHeights(PagedReport report, double tableHeight)
    {
        var rowCount = Math.Max(1, report.PageSettings.RowsPerPage);
        var headerMinHeight = PageMeasurementHelper.CalculateTableCellMinHeightDip(report.DetailHeaderFontSize);
        var contentMinHeight = PageMeasurementHelper.CalculateTableCellMinHeightDip(report.DetailContentFontSize);
        var totalMinHeight = headerMinHeight + (contentMinHeight * rowCount);
        var extraHeightPerRow = totalMinHeight >= tableHeight
            ? 0
            : (tableHeight - totalMinHeight) / (rowCount + 1);

        return (headerMinHeight + extraHeightPerRow, contentMinHeight + extraHeightPerRow);
    }

    private static IReadOnlyList<double> CalculateColumnWidths(IReadOnlyList<ReportColumnLayout> columns, double totalWidth)
    {
        var widths = new double[columns.Count];
        var totalWeight = columns.Sum(column => column.WidthWeight <= 0 ? 1 : column.WidthWeight);
        var runningWidth = 0d;

        for (var index = 0; index < columns.Count; index++)
        {
            widths[index] = index == columns.Count - 1
                ? totalWidth - runningWidth
                : totalWidth * ((columns[index].WidthWeight <= 0 ? 1 : columns[index].WidthWeight) / totalWeight);

            runningWidth += widths[index];
        }

        return widths;
    }

    private static TextAlignment ToTextAlignment(ReportTextAlignment alignment)
    {
        return alignment switch
        {
            ReportTextAlignment.Center => TextAlignment.Center,
            ReportTextAlignment.Right => TextAlignment.Right,
            _ => TextAlignment.Left
        };
    }
}
