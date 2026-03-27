using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using ReportGenerator.Core.Abstractions;
using ReportGenerator.Core.Models;
using ReportGenerator.Core.Utilities;

namespace ReportGenerator.Core.Services;

public sealed class PdfExportService : IPdfExportService
{
    static PdfExportService()
    {
        if (GlobalFontSettings.FontResolver is null)
        {
            GlobalFontSettings.FontResolver = new WindowsFontResolver();
        }
    }

    public Task ExportAsync(PagedReport report, string outputPath, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(report);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        return Task.Run(() => ExportInternal(report, outputPath, cancellationToken), cancellationToken);
    }

    private static void ExportInternal(PagedReport report, string outputPath, CancellationToken cancellationToken)
    {
        var directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        using var document = new PdfDocument();

        foreach (var page in report.Pages)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var pdfPage = document.AddPage();
            var (pageWidthMm, pageHeightMm) = PageMeasurementHelper.GetPageDimensionsMillimeters(report.PageSettings);
            pdfPage.Width = XUnit.FromPoint(PageMeasurementHelper.MillimetersToPoints(pageWidthMm));
            pdfPage.Height = XUnit.FromPoint(PageMeasurementHelper.MillimetersToPoints(pageHeightMm));

            using var graphics = XGraphics.FromPdfPage(pdfPage);
            DrawPage(graphics, pdfPage, report, page);
        }

        document.Save(outputPath);
    }

    private static void DrawPage(XGraphics graphics, PdfPage pdfPage, PagedReport report, ReportPage page)
    {
        var pageWidth = pdfPage.Width.Point;
        var pageHeight = pdfPage.Height.Point;
        var marginLeft = PageMeasurementHelper.MillimetersToPoints(report.PageSettings.MarginLeftMm);
        var marginTop = PageMeasurementHelper.MillimetersToPoints(report.PageSettings.MarginTopMm);
        var marginRight = PageMeasurementHelper.MillimetersToPoints(report.PageSettings.MarginRightMm);
        var marginBottom = PageMeasurementHelper.MillimetersToPoints(report.PageSettings.MarginBottomMm);
        var contentWidth = pageWidth - marginLeft - marginRight;

        var headerHeight = CalculateBlocksHeight(page.HeaderBlocks);
        var footerHeight = CalculateBlocksHeight(page.FooterBlocks);
        var footerTop = pageHeight - marginBottom - footerHeight;
        var sectionSpacing = PageMeasurementHelper.DipToPoints(PageMeasurementHelper.SectionSpacingDip);
        var tableTop = marginTop + headerHeight + sectionSpacing;
        var tableHeight = Math.Max(PageMeasurementHelper.DipToPoints(PageMeasurementHelper.MinimumTableHeightDip), footerTop - tableTop - sectionSpacing);
        var (headerRowHeight, contentRowHeight) = CalculateTableRowHeights(report, tableHeight);
        var columnWidths = CalculateColumnWidths(report.Columns, contentWidth);

        DrawBlocks(graphics, page.HeaderBlocks, marginLeft, marginTop, contentWidth);
        DrawTable(graphics, report, page, marginLeft, tableTop, headerRowHeight, contentRowHeight, columnWidths);
        DrawBlocks(graphics, page.FooterBlocks, marginLeft, footerTop, contentWidth);
    }

    private static void DrawBlocks(
        XGraphics graphics,
        IReadOnlyList<ReportBlockContent> blocks,
        double left,
        double top,
        double width)
    {
        var y = top;
        foreach (var block in blocks)
        {
            var font = CreateFont(block.FontSize, block.IsBold);
            var lineHeight = CalculateTextLineHeight(block.FontSize);
            var rect = new XRect(left, y, width, lineHeight);
            graphics.DrawString(block.Text, font, XBrushes.Black, rect, GetBlockStringFormat(block.Alignment));
            y += lineHeight;
        }
    }

    private static void DrawTable(
        XGraphics graphics,
        PagedReport report,
        ReportPage page,
        double left,
        double top,
        double headerRowHeight,
        double contentRowHeight,
        IReadOnlyList<double> columnWidths)
    {
        var headerBrush = new XSolidBrush(XColor.FromArgb(245, 247, 250));
        var borderPen = new XPen(XColor.FromArgb(210, 214, 220), 0.7);
        var textBrush = XBrushes.Black;
        var tableHeaderFont = CreateFont(report.DetailHeaderFontSize, true);
        var tableRowFont = CreateFont(report.DetailContentFontSize, false);

        var currentX = left;
        for (var columnIndex = 0; columnIndex < report.Columns.Count; columnIndex++)
        {
            var column = report.Columns[columnIndex];
            var cellRect = new XRect(currentX, top, columnWidths[columnIndex], headerRowHeight);
            graphics.DrawRectangle(headerBrush, cellRect);
            graphics.DrawRectangle(borderPen, cellRect);
            DrawCellText(graphics, column.HeaderText, tableHeaderFont, textBrush, cellRect, column.Alignment);
            currentX += columnWidths[columnIndex];
        }

        for (var rowIndex = 0; rowIndex < page.Rows.Count; rowIndex++)
        {
            var row = page.Rows[rowIndex];
            var y = top + headerRowHeight + (rowIndex * contentRowHeight);
            if (row.IsSpacer)
            {
                continue;
            }

            currentX = left;

            for (var columnIndex = 0; columnIndex < row.Cells.Count; columnIndex++)
            {
                var cell = row.Cells[columnIndex];
                var cellRect = new XRect(currentX, y, columnWidths[columnIndex], contentRowHeight);
                graphics.DrawRectangle(borderPen, cellRect);
                DrawCellText(graphics, cell.Text, tableRowFont, textBrush, cellRect, cell.Alignment);
                currentX += columnWidths[columnIndex];
            }
        }
    }

    private static void DrawCellText(
        XGraphics graphics,
        string text,
        XFont font,
        XBrush brush,
        XRect rect,
        ReportTextAlignment alignment)
    {
        var horizontalPadding = PageMeasurementHelper.DipToPoints(4);
        var verticalPadding = PageMeasurementHelper.DipToPoints(PageMeasurementHelper.TableCellVerticalPaddingDip);
        var paddedRect = new XRect(
            rect.X + horizontalPadding,
            rect.Y + verticalPadding,
            Math.Max(0, rect.Width - (horizontalPadding * 2)),
            Math.Max(0, rect.Height - (verticalPadding * 2)));
        graphics.DrawString(text ?? string.Empty, font, brush, paddedRect, GetCellStringFormat(alignment));
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

    private static double CalculateBlocksHeight(IEnumerable<ReportBlockContent> blocks)
    {
        return blocks.Sum(block => CalculateTextLineHeight(block.FontSize));
    }

    private static XFont CreateFont(double fontSize, bool isBold)
    {
        return new XFont("Segoe UI", PageMeasurementHelper.DipToPoints(fontSize), isBold ? XFontStyleEx.Bold : XFontStyleEx.Regular);
    }

    private static double CalculateTextLineHeight(double fontSize)
    {
        return PageMeasurementHelper.DipToPoints(PageMeasurementHelper.CalculateTextLineHeightDip(fontSize));
    }

    private static XStringFormat GetBlockStringFormat(ReportTextAlignment alignment)
    {
        return alignment switch
        {
            ReportTextAlignment.Center => XStringFormats.TopCenter,
            ReportTextAlignment.Right => XStringFormats.TopRight,
            _ => XStringFormats.TopLeft
        };
    }

    private static XStringFormat GetCellStringFormat(ReportTextAlignment alignment)
    {
        return alignment switch
        {
            ReportTextAlignment.Center => XStringFormats.Center,
            ReportTextAlignment.Right => XStringFormats.CenterRight,
            _ => XStringFormats.CenterLeft
        };
    }

    private static (double HeaderRowHeight, double ContentRowHeight) CalculateTableRowHeights(PagedReport report, double tableHeight)
    {
        var rowCount = Math.Max(1, report.PageSettings.RowsPerPage);
        var headerMinHeight = PageMeasurementHelper.DipToPoints(PageMeasurementHelper.CalculateTableCellMinHeightDip(report.DetailHeaderFontSize));
        var contentMinHeight = PageMeasurementHelper.DipToPoints(PageMeasurementHelper.CalculateTableCellMinHeightDip(report.DetailContentFontSize));
        var totalMinHeight = headerMinHeight + (contentMinHeight * rowCount);
        var extraHeightPerRow = totalMinHeight >= tableHeight
            ? 0
            : (tableHeight - totalMinHeight) / (rowCount + 1);

        return (headerMinHeight + extraHeightPerRow, contentMinHeight + extraHeightPerRow);
    }
}
