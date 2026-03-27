namespace ReportGenerator.Core.Models;

public sealed class ImportedDataRow
{
    public ImportedDataRow(int rowNumber, IReadOnlyList<string?> cells)
    {
        RowNumber = rowNumber;
        Cells = cells;
    }

    public int RowNumber { get; }

    public IReadOnlyList<string?> Cells { get; }

    public string? GetCell(int zeroBasedColumnIndex)
    {
        return zeroBasedColumnIndex >= 0 && zeroBasedColumnIndex < Cells.Count
            ? Cells[zeroBasedColumnIndex]
            : null;
    }
}
