using System.Collections.ObjectModel;

namespace ReportGenerator.Core.Models;

public sealed class ReportPage
{
    public required int PageNumber { get; init; }

    public required int TotalPages { get; init; }

    public required IReadOnlyList<ReportBlockContent> HeaderBlocks { get; init; }

    public required IReadOnlyList<ReportRow> Rows { get; init; }

    public required IReadOnlyList<ReportBlockContent> FooterBlocks { get; init; }

    public static ReportPage Create(
        int pageNumber,
        int totalPages,
        IEnumerable<ReportBlockContent> headerBlocks,
        IEnumerable<ReportRow> rows,
        IEnumerable<ReportBlockContent> footerBlocks)
    {
        return new ReportPage
        {
            PageNumber = pageNumber,
            TotalPages = totalPages,
            HeaderBlocks = new ReadOnlyCollection<ReportBlockContent>(headerBlocks.ToArray()),
            Rows = new ReadOnlyCollection<ReportRow>(rows.ToArray()),
            FooterBlocks = new ReadOnlyCollection<ReportBlockContent>(footerBlocks.ToArray())
        };
    }
}
