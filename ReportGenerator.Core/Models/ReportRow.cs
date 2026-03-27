using System.Collections.ObjectModel;

namespace ReportGenerator.Core.Models;

public sealed class ReportRow
{
    public required int SourceRowNumber { get; init; }

    public required IReadOnlyList<ReportCell> Cells { get; init; }

    public required bool IsSpacer { get; init; }

    public static ReportRow Create(int sourceRowNumber, IEnumerable<ReportCell> cells)
    {
        return new ReportRow
        {
            SourceRowNumber = sourceRowNumber,
            Cells = new ReadOnlyCollection<ReportCell>(cells.ToArray()),
            IsSpacer = false
        };
    }

    public static ReportRow CreateSpacer()
    {
        return new ReportRow
        {
            SourceRowNumber = -1,
            Cells = Array.Empty<ReportCell>(),
            IsSpacer = true
        };
    }
}
