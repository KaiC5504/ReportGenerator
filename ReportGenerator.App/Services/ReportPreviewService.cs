using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Markup;
using System.Windows.Media;
using ReportGenerator.Core.Models;
using ReportGenerator.Core.Utilities;

namespace ReportGenerator.App.Services;

public sealed class ReportPreviewService
{
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
        var tableTop = marginTop + headerHeight + 12;
        var tableHeight = Math.Max(36, footerTop - tableTop - 12);
        var rowHeight = tableHeight / Math.Max(1, report.PageSettings.RowsPerPage + 1);
        var columnWidths = CalculateColumnWidths(report.Columns, contentWidth);

        AddBlocks(fixedPage, reportPage.HeaderBlocks, marginLeft, marginTop, contentWidth);
        AddTable(fixedPage, report, reportPage, marginLeft, tableTop, rowHeight, columnWidths);
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
            var lineHeight = Math.Max(block.FontSize * 1.45, 14);
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
        double rowHeight,
        IReadOnlyList<double> columnWidths)
    {
        var currentX = left;
        for (var columnIndex = 0; columnIndex < report.Columns.Count; columnIndex++)
        {
            var column = report.Columns[columnIndex];
            var headerCell = CreateCell(column.HeaderText, column.Alignment, columnWidths[columnIndex], rowHeight, true);
            FixedPage.SetLeft(headerCell, currentX);
            FixedPage.SetTop(headerCell, top);
            fixedPage.Children.Add(headerCell);
            currentX += columnWidths[columnIndex];
        }

        for (var rowIndex = 0; rowIndex < page.Rows.Count; rowIndex++)
        {
            var y = top + ((rowIndex + 1) * rowHeight);
            var row = page.Rows[rowIndex];
            if (row.IsSpacer)
            {
                continue;
            }

            currentX = left;

            for (var columnIndex = 0; columnIndex < row.Cells.Count; columnIndex++)
            {
                var cell = row.Cells[columnIndex];
                var cellVisual = CreateCell(cell.Text, cell.Alignment, columnWidths[columnIndex], rowHeight, false, report.DetailContentFontSize);
                FixedPage.SetLeft(cellVisual, currentX);
                FixedPage.SetTop(cellVisual, y);
                fixedPage.Children.Add(cellVisual);
                currentX += columnWidths[columnIndex];
            }
        }
    }

    private static Border CreateCell(string text, ReportTextAlignment alignment, double width, double height, bool isHeader, double contentFontSize = 10)
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
                FontSize = isHeader ? 10 : contentFontSize,
                FontWeight = isHeader ? FontWeights.SemiBold : FontWeights.Normal,
                TextAlignment = ToTextAlignment(alignment),
                TextTrimming = TextTrimming.CharacterEllipsis,
                VerticalAlignment = VerticalAlignment.Center
            }
        };
    }

    private static double CalculateBlocksHeight(IEnumerable<ReportBlockContent> blocks)
    {
        return blocks.Sum(block => Math.Max(block.FontSize * 1.45, 14));
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
