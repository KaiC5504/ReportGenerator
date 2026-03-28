namespace ReportGenerator.Core.Models;

public sealed class ReportBlockContent
{
    public required ReportBlockType Type { get; init; }

    public required string Text { get; init; }

    public required string TemplateText { get; init; }

    public required ReportTextAlignment Alignment { get; init; }

    public required int Row { get; init; }

    public required double FontSize { get; init; }

    public required bool IsBold { get; init; }
}
