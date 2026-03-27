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
        var tableTop = marginTop + headerHeight + 12;
        var tableHeight = Math.Max(36, footerTop - tableTop - 12);
        var rowHeight = tableHeight / Math.Max(1, report.PageSettings.RowsPerPage + 1);
        var columnWidths = CalculateColumnWidths(report.Columns, contentWidth);

        DrawBlocks(graphics, page.HeaderBlocks, marginLeft, marginTop, contentWidth);
        DrawTable(graphics, report, page, marginLeft, tableTop, contentWidth, rowHeight, columnWidths);
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
            var lineHeight = Math.Max(block.FontSize * 1.45, 14);
            var rect = new XRect(left, y, width, lineHeight);
            graphics.DrawString(block.Text, font, XBrushes.Black, rect, GetStringFormat(block.Alignment));
            y += lineHeight;
        }
    }

    private static void DrawTable(
        XGraphics graphics,
        PagedReport report,
        ReportPage page,
        double left,
        double top,
        double width,
        double rowHeight,
        IReadOnlyList<double> columnWidths)
    {
        var headerBrush = new XSolidBrush(XColor.FromArgb(245, 247, 250));
        var borderPen = new XPen(XColor.FromArgb(210, 214, 220), 0.7);
        var textBrush = XBrushes.Black;
        var tableHeaderFont = CreateFont(10, true);
        var tableRowFont = CreateFont(report.DetailContentFontSize, false);

        var currentX = left;
        for (var columnIndex = 0; columnIndex < report.Columns.Count; columnIndex++)
        {
            var column = report.Columns[columnIndex];
            var cellRect = new XRect(currentX, top, columnWidths[columnIndex], rowHeight);
            graphics.DrawRectangle(headerBrush, cellRect);
            graphics.DrawRectangle(borderPen, cellRect);
            DrawCellText(graphics, column.HeaderText, tableHeaderFont, textBrush, cellRect, column.Alignment);
            currentX += columnWidths[columnIndex];
        }

        for (var rowIndex = 0; rowIndex < page.Rows.Count; rowIndex++)
        {
            var row = page.Rows[rowIndex];
            var y = top + ((rowIndex + 1) * rowHeight);
            if (row.IsSpacer)
            {
                continue;
            }

            currentX = left;

            for (var columnIndex = 0; columnIndex < row.Cells.Count; columnIndex++)
            {
                var cell = row.Cells[columnIndex];
                var cellRect = new XRect(currentX, y, columnWidths[columnIndex], rowHeight);
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
        var padding = 4d;
        var paddedRect = new XRect(rect.X + padding, rect.Y + 2, Math.Max(0, rect.Width - (padding * 2)), Math.Max(0, rect.Height - 4));
        graphics.DrawString(text ?? string.Empty, font, brush, paddedRect, GetStringFormat(alignment));
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
        return blocks.Sum(block => Math.Max(block.FontSize * 1.45, 14));
    }

    private static XFont CreateFont(double fontSize, bool isBold)
    {
        return new XFont("Segoe UI", fontSize, isBold ? XFontStyleEx.Bold : XFontStyleEx.Regular);
    }

    private static XStringFormat GetStringFormat(ReportTextAlignment alignment)
    {
        return alignment switch
        {
            ReportTextAlignment.Center => XStringFormats.TopCenter,
            ReportTextAlignment.Right => XStringFormats.TopRight,
            _ => XStringFormats.TopLeft
        };
    }
}
