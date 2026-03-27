namespace ReportGenerator.Core.Models;

public sealed class PageSettings
{
    public ReportPageSize PageSize { get; set; } = ReportPageSize.A4;

    public ReportOrientation Orientation { get; set; } = ReportOrientation.Portrait;

    public double MarginTopMm { get; set; } = 18;

    public double MarginRightMm { get; set; } = 14;

    public double MarginBottomMm { get; set; } = 18;

    public double MarginLeftMm { get; set; } = 14;

    public int RowsPerPage { get; set; } = 30;

    public bool HeaderOnlyOnFirstPage { get; set; }
}
