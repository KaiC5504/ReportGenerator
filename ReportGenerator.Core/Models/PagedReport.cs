using System.Collections.ObjectModel;

namespace ReportGenerator.Core.Models;

public sealed class PagedReport
{
    public required string TemplateName { get; init; }

    public required PageSettings PageSettings { get; init; }

    public required string SourceFileName { get; init; }

    public required string SheetName { get; init; }

    public required DateTimeOffset GeneratedAtUtc { get; init; }

    public required IReadOnlyList<ReportColumnLayout> Columns { get; init; }

    public required IReadOnlyList<ReportPage> Pages { get; init; }

    public required double DetailHeaderFontSize { get; init; }

    public required double DetailContentFontSize { get; init; }

    public required double DetailContentRowSpacing { get; init; }

    public required int TotalRows { get; init; }

    public static PagedReport Create(
        string templateName,
        PageSettings pageSettings,
        string sourceFileName,
        string sheetName,
        DateTimeOffset generatedAtUtc,
        IEnumerable<ReportColumnLayout> columns,
        IEnumerable<ReportPage> pages,
        double detailHeaderFontSize,
        double detailContentFontSize,
        double detailContentRowSpacing,
        int totalRows)
    {
        return new PagedReport
        {
            TemplateName = templateName,
            PageSettings = pageSettings,
            SourceFileName = sourceFileName,
            SheetName = sheetName,
            GeneratedAtUtc = generatedAtUtc,
            Columns = new ReadOnlyCollection<ReportColumnLayout>(columns.ToArray()),
            Pages = new ReadOnlyCollection<ReportPage>(pages.ToArray()),
            DetailHeaderFontSize = detailHeaderFontSize,
            DetailContentFontSize = detailContentFontSize,
            DetailContentRowSpacing = detailContentRowSpacing,
            TotalRows = totalRows
        };
    }
}
