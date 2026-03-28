using ReportGenerator.Core.Models;

namespace ReportGenerator.Core.Utilities;

public static class PageMeasurementHelper
{
    public const double SectionSpacingDip = 6;
    public const double MinimumTableHeightDip = 36;
    public const double TableCellVerticalPaddingDip = 2;
    public const double ContentCellVerticalPaddingDip = 0;

    private const double MillimetersPerInch = 25.4;
    private const double PointsPerInch = 72;
    private const double DipPerInch = 96;
    private const double TwipsPerPoint = 20;

    public static (double WidthMm, double HeightMm) GetPageDimensionsMillimeters(PageSettings settings)
    {
        var (width, height) = settings.PageSize switch
        {
            ReportPageSize.A4 => (210d, 297d),
            ReportPageSize.Letter => (215.9d, 279.4d),
            _ => (210d, 297d)
        };

        return settings.Orientation == ReportOrientation.Landscape
            ? (height, width)
            : (width, height);
    }

    public static double MillimetersToPoints(double millimeters)
    {
        return (millimeters / MillimetersPerInch) * PointsPerInch;
    }

    public static double MillimetersToDip(double millimeters)
    {
        return (millimeters / MillimetersPerInch) * DipPerInch;
    }

    public static double DipToPoints(double dip)
    {
        return dip * (PointsPerInch / DipPerInch);
    }

    public static double CalculateTextLineHeightDip(double fontSize)
    {
        return Math.Max(fontSize * 1.45, 14);
    }

    public static double CalculateTableCellMinHeightDip(double fontSize)
    {
        return CalculateTextLineHeightDip(fontSize) + (TableCellVerticalPaddingDip * 2);
    }

    public static double CalculateContentRowHeightDip(double fontSize, double contentRowSpacing)
    {
        return CalculateTextLineHeightDip(fontSize) * (1 + Math.Max(0, contentRowSpacing));
    }

    public static int MillimetersToTwips(double millimeters)
    {
        return (int)Math.Round(MillimetersToPoints(millimeters) * TwipsPerPoint, MidpointRounding.AwayFromZero);
    }

    public static int DipToTwips(double dip)
    {
        return (int)Math.Round(DipToPoints(dip) * TwipsPerPoint, MidpointRounding.AwayFromZero);
    }
}
