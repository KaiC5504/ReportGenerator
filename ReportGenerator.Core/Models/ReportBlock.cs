namespace ReportGenerator.Core.Models;

public sealed class ReportBlock
{
    public ReportBlockType Type { get; set; } = ReportBlockType.StaticText;

    public string Text { get; set; } = string.Empty;

    public string Source { get; set; } = string.Empty;

    public ReportTextAlignment Alignment { get; set; } = ReportTextAlignment.Left;

    public double FontSize { get; set; } = 11;

    public bool IsBold { get; set; }
}
