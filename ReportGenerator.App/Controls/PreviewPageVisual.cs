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
    private static readonly Brush HeaderCellBackgroundBrush = CreateFrozenBrush(Color.FromRgb(245, 247, 250));
    private static readonly Pen CellBorderPen = CreateFrozenPen(Color.FromRgb(210, 214, 220), 0.7);
    private static readonly Typeface NormalTypeface = new(DefaultFontFamily, FontStyles.Normal, FontWeights.Normal, FontStretches.Normal);
    private static readonly Typeface SemiBoldTypeface = new(DefaultFontFamily, FontStyles.Normal, FontWeights.SemiBold, FontStretches.Normal);

    private readonly PagedReport _report;
    private readonly ReportPage _reportPage;
    private readonly Size _pageSize;
    private readonly double _marginLeft;
    private readonly double _marginTop;
    private readonly double _contentWidth;
    private readonly double _footerTop;
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
        var marginBottom = PageMeasurementHelper.MillimetersToDip(report.PageSettings.MarginBottomMm);
        _contentWidth = pageSize.Width - _marginLeft - marginRight;

        var headerHeight = CalculateBlocksHeight(reportPage.HeaderBlocks);
        var footerHeight = CalculateBlocksHeight(reportPage.FooterBlocks);
        _footerTop = pageSize.Height - marginBottom - footerHeight;
        _tableTop = _marginTop + headerHeight + PageMeasurementHelper.SectionSpacingDip;

        var tableHeight = Math.Max(
            PageMeasurementHelper.MinimumTableHeightDip,
            _footerTop - _tableTop - PageMeasurementHelper.SectionSpacingDip);
        (_headerRowHeight, _contentRowHeight) = CalculateTableRowHeights(report, tableHeight);
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

        DrawBlocks(drawingContext, _reportPage.HeaderBlocks, _marginLeft, _marginTop, _contentWidth);
        DrawTable(drawingContext);
        DrawBlocks(drawingContext, _reportPage.FooterBlocks, _marginLeft, _footerTop, _contentWidth);
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

    private void DrawTable(DrawingContext drawingContext)
    {
        var currentX = _marginLeft;
        for (var columnIndex = 0; columnIndex < _report.Columns.Count; columnIndex++)
        {
            var column = _report.Columns[columnIndex];
            DrawCell(
                drawingContext,
                column.HeaderText,
                column.Alignment,
                currentX,
                _tableTop,
                _columnWidths[columnIndex],
                _headerRowHeight,
                HeaderCellBackgroundBrush,
                SemiBoldTypeface,
                _report.DetailHeaderFontSize);
            currentX += _columnWidths[columnIndex];
        }

        for (var rowIndex = 0; rowIndex < _reportPage.Rows.Count; rowIndex++)
        {
            var row = _reportPage.Rows[rowIndex];
            if (row.IsSpacer)
            {
                continue;
            }

            var y = _tableTop + _headerRowHeight + (rowIndex * _contentRowHeight);
            currentX = _marginLeft;

            for (var columnIndex = 0; columnIndex < row.Cells.Count; columnIndex++)
            {
                var cell = row.Cells[columnIndex];
                DrawCell(
                    drawingContext,
                    cell.Text,
                    cell.Alignment,
                    currentX,
                    y,
                    _columnWidths[columnIndex],
                    _contentRowHeight,
                    PageBackgroundBrush,
                    NormalTypeface,
                    _report.DetailContentFontSize);
                currentX += _columnWidths[columnIndex];
            }
        }
    }

    private void DrawCell(
        DrawingContext drawingContext,
        string text,
        ReportTextAlignment alignment,
        double left,
        double top,
        double width,
        double height,
        Brush background,
        Typeface typeface,
        double fontSize)
    {
        var cellRect = new Rect(left, top, width, height);
        drawingContext.DrawRectangle(background, CellBorderPen, cellRect);

        DrawText(
            drawingContext,
            text,
            alignment,
            fontSize,
            typeface,
            new Rect(
                cellRect.Left + CellHorizontalPaddingDip,
                cellRect.Top + PageMeasurementHelper.TableCellVerticalPaddingDip,
                Math.Max(0, cellRect.Width - (CellHorizontalPaddingDip * 2)),
                Math.Max(0, cellRect.Height - (PageMeasurementHelper.TableCellVerticalPaddingDip * 2))));
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
