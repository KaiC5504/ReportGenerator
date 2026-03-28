using System.Globalization;
using System.Windows;
using System.Windows.Media;
using ReportGenerator.Core.Models;
using ReportGenerator.Core.Utilities;

namespace ReportGenerator.App.Controls;

public sealed class PreviewPageVisual : FrameworkElement
{
    private const double CellHorizontalPaddingDip = 4;

    private static readonly FontFamily DefaultFontFamily = new("Segoe UI");
    private static readonly Brush TextBrush = Brushes.Black;
    private static readonly Brush PageBackgroundBrush = Brushes.White;
    private static readonly Pen RulePen = CreateFrozenPen(Colors.Black, 1);
    private static readonly Typeface NormalTypeface = new(DefaultFontFamily, FontStyles.Normal, FontWeights.Normal, FontStretches.Normal);
    private static readonly Typeface SemiBoldTypeface = new(DefaultFontFamily, FontStyles.Normal, FontWeights.SemiBold, FontStretches.Normal);

    private readonly PagedReport _report;
    private readonly ReportPage _reportPage;
    private readonly Size _pageSize;
    private readonly double _marginLeft;
    private readonly double _marginTop;
    private readonly double _contentWidth;
    private readonly double _tableTop;
    private readonly double _headerRowHeight;
    private readonly double _contentRowHeight;
    private readonly IReadOnlyList<double> _columnWidths;

    public PreviewPageVisual(PagedReport report, ReportPage reportPage, Size pageSize)
    {
        _report = report;
        _reportPage = reportPage;
        _pageSize = pageSize;

        Width = pageSize.Width;
        Height = pageSize.Height;
        Focusable = false;
        IsHitTestVisible = false;
        SnapsToDevicePixels = true;

        TextOptions.SetTextFormattingMode(this, TextFormattingMode.Display);

        _marginLeft = PageMeasurementHelper.MillimetersToDip(report.PageSettings.MarginLeftMm);
        _marginTop = PageMeasurementHelper.MillimetersToDip(report.PageSettings.MarginTopMm);

        var marginRight = PageMeasurementHelper.MillimetersToDip(report.PageSettings.MarginRightMm);
        _contentWidth = pageSize.Width - _marginLeft - marginRight;

        var headerHeight = CalculateHeaderHeight(reportPage.HeaderBlocks);
        _tableTop = _marginTop + headerHeight + PageMeasurementHelper.SectionSpacingDip;

        (_headerRowHeight, _contentRowHeight) = CalculateTableRowHeights(report);
        _columnWidths = CalculateColumnWidths(report.Columns, _contentWidth);
    }

    protected override Size MeasureOverride(Size availableSize)
    {
        return _pageSize;
    }

    protected override Size ArrangeOverride(Size finalSize)
    {
        return _pageSize;
    }

    protected override void OnRender(DrawingContext drawingContext)
    {
        base.OnRender(drawingContext);

        drawingContext.DrawRectangle(PageBackgroundBrush, null, new Rect(new Point(0, 0), _pageSize));

        DrawHeaderRows(drawingContext, _reportPage.HeaderBlocks, _marginLeft, _marginTop, _contentWidth);
        var footerTop = DrawTable(drawingContext);
        DrawBlocks(drawingContext, _reportPage.FooterBlocks, _marginLeft, footerTop, _contentWidth);
    }

    private void DrawHeaderRows(
        DrawingContext drawingContext,
        IReadOnlyList<ReportBlockContent> blocks,
        double left,
        double top,
        double width)
    {
        var currentY = top;
        foreach (var rowBlocks in blocks.GroupBy(block => block.Row).OrderBy(group => group.Key))
        {
            var rowHeight = rowBlocks.Max(block => PageMeasurementHelper.CalculateTextLineHeightDip(block.FontSize));
            foreach (var block in rowBlocks)
            {
                DrawText(
                    drawingContext,
                    block.Text,
                    block.Alignment,
                    block.FontSize,
                    block.IsBold ? SemiBoldTypeface : NormalTypeface,
                    new Rect(left, currentY, width, rowHeight));
            }

            currentY += rowHeight;
        }
    }

    private void DrawBlocks(
        DrawingContext drawingContext,
        IReadOnlyList<ReportBlockContent> blocks,
        double left,
        double top,
        double width)
    {
        var currentY = top;
        foreach (var block in blocks)
        {
            var lineHeight = PageMeasurementHelper.CalculateTextLineHeightDip(block.FontSize);
            DrawText(
                drawingContext,
                block.Text,
                block.Alignment,
                block.FontSize,
                block.IsBold ? SemiBoldTypeface : NormalTypeface,
                new Rect(left, currentY, width, lineHeight));

            currentY += lineHeight;
        }
    }

    private double DrawTable(DrawingContext drawingContext)
    {
        DrawHorizontalRule(drawingContext, _tableTop);

        var currentX = _marginLeft;
        for (var columnIndex = 0; columnIndex < _report.Columns.Count; columnIndex++)
        {
            var column = _report.Columns[columnIndex];
            DrawText(
                drawingContext,
                column.HeaderText,
                column.Alignment,
                _report.DetailHeaderFontSize,
                SemiBoldTypeface,
                new Rect(
                    currentX + CellHorizontalPaddingDip,
                    _tableTop + PageMeasurementHelper.TableCellVerticalPaddingDip,
                    Math.Max(0, _columnWidths[columnIndex] - (CellHorizontalPaddingDip * 2)),
                    Math.Max(0, _headerRowHeight - (PageMeasurementHelper.TableCellVerticalPaddingDip * 2))));
            currentX += _columnWidths[columnIndex];
        }

        var contentTop = _tableTop + _headerRowHeight;
        DrawHorizontalRule(drawingContext, contentTop);

        var currentY = contentTop;
        for (var rowIndex = 0; rowIndex < _reportPage.Rows.Count; rowIndex++)
        {
            var row = _reportPage.Rows[rowIndex];
            var rowHeight = _contentRowHeight * row.HeightFactor;
            if (row.IsSpacer)
            {
                currentY += rowHeight;
                continue;
            }

            currentX = _marginLeft;

            for (var columnIndex = 0; columnIndex < row.Cells.Count; columnIndex++)
            {
                var cell = row.Cells[columnIndex];
                DrawText(
                    drawingContext,
                    cell.Text,
                    cell.Alignment,
                    _report.DetailContentFontSize,
                    NormalTypeface,
                    new Rect(
                        currentX + CellHorizontalPaddingDip,
                        currentY + PageMeasurementHelper.ContentCellVerticalPaddingDip,
                        Math.Max(0, _columnWidths[columnIndex] - (CellHorizontalPaddingDip * 2)),
                        Math.Max(0, rowHeight - (PageMeasurementHelper.ContentCellVerticalPaddingDip * 2))));
                currentX += _columnWidths[columnIndex];
            }

            currentY += rowHeight;
        }

        DrawHorizontalRule(drawingContext, currentY);
        return currentY;
    }

    private void DrawText(
        DrawingContext drawingContext,
        string text,
        ReportTextAlignment alignment,
        double fontSize,
        Typeface typeface,
        Rect bounds)
    {
        if (bounds.Width <= 0 || bounds.Height <= 0 || string.IsNullOrEmpty(text))
        {
            return;
        }

        var formattedText = new FormattedText(
            text,
            CultureInfo.CurrentUICulture,
            FlowDirection.LeftToRight,
            typeface,
            fontSize,
            TextBrush,
            VisualTreeHelper.GetDpi(this).PixelsPerDip)
        {
            MaxTextWidth = bounds.Width,
            MaxTextHeight = bounds.Height,
            TextAlignment = ToTextAlignment(alignment),
            Trimming = TextTrimming.CharacterEllipsis
        };

        var y = bounds.Top + Math.Max(0, (bounds.Height - formattedText.Height) / 2);
        drawingContext.PushClip(new RectangleGeometry(bounds));
        drawingContext.DrawText(formattedText, new Point(bounds.Left, y));
        drawingContext.Pop();
    }

    private void DrawHorizontalRule(DrawingContext drawingContext, double y)
    {
        drawingContext.DrawLine(RulePen, new Point(_marginLeft, y), new Point(_marginLeft + _contentWidth, y));
    }

    private static double CalculateHeaderHeight(IEnumerable<ReportBlockContent> blocks)
    {
        return blocks
            .GroupBy(block => block.Row)
            .Sum(group => group.Max(block => PageMeasurementHelper.CalculateTextLineHeightDip(block.FontSize)));
    }

    private static (double HeaderRowHeight, double ContentRowHeight) CalculateTableRowHeights(PagedReport report)
    {
        var headerMinHeight = PageMeasurementHelper.CalculateTableCellMinHeightDip(report.DetailHeaderFontSize);
        var contentRowHeight = PageMeasurementHelper.CalculateContentRowHeightDip(
            report.DetailContentFontSize,
            report.DetailContentRowSpacing);

        return (headerMinHeight, contentRowHeight);
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

    private static Brush CreateFrozenBrush(Color color)
    {
        var brush = new SolidColorBrush(color);
        brush.Freeze();
        return brush;
    }

    private static Pen CreateFrozenPen(Color color, double thickness)
    {
        var pen = new Pen(CreateFrozenBrush(color), thickness);
        pen.Freeze();
        return pen;
    }
}
