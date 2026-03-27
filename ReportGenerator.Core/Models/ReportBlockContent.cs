namespace ReportGenerator.Core.Models;

public sealed class ReportBlockContent
{
    public required string Text { get; init; }

    public required ReportTextAlignment Alignment { get; init; }

    public required double FontSize { get; init; }

    public required bool IsBold { get; init; }
}
