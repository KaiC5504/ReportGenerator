using ReportGenerator.Core.Models;

namespace ReportGenerator.Core.Utilities;

public static class PageMeasurementHelper
{
    private const double MillimetersPerInch = 25.4;
    private const double PointsPerInch = 72;
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
        return (millimeters / MillimetersPerInch) * 96d;
    }

    public static int MillimetersToTwips(double millimeters)
    {
        return (int)Math.Round(MillimetersToPoints(millimeters) * TwipsPerPoint, MidpointRounding.AwayFromZero);
    }
}
