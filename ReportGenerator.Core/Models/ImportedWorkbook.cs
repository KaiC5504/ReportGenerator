using System.Collections.ObjectModel;
using System.IO;

namespace ReportGenerator.Core.Models;

public sealed class ImportedWorkbook
{
    public ImportedWorkbook(string sourcePath, string sheetName, IReadOnlyList<string> headers, IReadOnlyList<ImportedDataRow> rows)
    {
        SourcePath = sourcePath;
        SheetName = sheetName;
        Headers = new ReadOnlyCollection<string>(headers.ToArray());
        Rows = new ReadOnlyCollection<ImportedDataRow>(rows.ToArray());
    }

    public string SourcePath { get; }

    public string FileName => Path.GetFileName(SourcePath);

    public string SheetName { get; }

    public IReadOnlyList<string> Headers { get; }

    public IReadOnlyList<ImportedDataRow> Rows { get; }

    public int ColumnCount => Headers.Count;

    public ImportedDataRow? FirstRow => Rows.Count > 0 ? Rows[0] : null;
}
