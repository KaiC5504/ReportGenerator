namespace ReportGenerator.Core.Models;

public sealed class DetailColumnDefinition
{
    public string HeaderText { get; set; } = "Column";

    public string Source { get; set; } = string.Empty;

    public double WidthWeight { get; set; } = 1;

    public ReportTextAlignment Alignment { get; set; } = ReportTextAlignment.Left;
}
