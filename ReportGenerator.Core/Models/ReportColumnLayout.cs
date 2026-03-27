namespace ReportGenerator.Core.Models;

public sealed class ReportColumnLayout
{
    public required string HeaderText { get; init; }

    public required string Source { get; init; }

    public required double WidthWeight { get; init; }

    public required ReportTextAlignment Alignment { get; init; }
}
