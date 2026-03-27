namespace ReportGenerator.Core.Models;

public sealed class ReportCell
{
    public required string Text { get; init; }

    public required ReportTextAlignment Alignment { get; init; }
}
