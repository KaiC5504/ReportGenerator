using System.Globalization;
using System.Text;
using ExcelDataReader;
using ReportGenerator.Core.Abstractions;
using ReportGenerator.Core.Models;
using ReportGenerator.Core.Utilities;

namespace ReportGenerator.Core.Services;

public sealed class ExcelImportService : IExcelImportService
{
    static ExcelImportService()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    public Task<ImportedWorkbook> ImportAsync(string filePath, ImportSettings settings, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        ArgumentNullException.ThrowIfNull(settings);

        return Task.Run(() => ImportInternal(filePath, settings, cancellationToken), cancellationToken);
    }

    private static ImportedWorkbook ImportInternal(string filePath, ImportSettings settings, CancellationToken cancellationToken)
    {
        using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = ExcelReaderFactory.CreateReader(stream);

        var sheetName = reader.Name ?? "Sheet1";
        var rows = new List<ImportedDataRow>();
        string[]? headers = null;
        var sheetRowNumber = 0;

        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            sheetRowNumber++;

            var cells = ReadRow(reader);
            var isEntirelyBlank = cells.All(static value => string.IsNullOrWhiteSpace(value));

            if (settings.HeaderRowEnabled && headers is null)
            {
                headers = CreateHeadersFromRow(cells);
                continue;
            }

            if (isEntirelyBlank)
            {
                continue;
            }

            rows.Add(new ImportedDataRow(sheetRowNumber, cells));
        }

        headers ??= CreateGeneratedHeaders(rows);
        var normalizedRows = rows.Select(row => NormalizeRow(row, headers.Length)).ToArray();
        return new ImportedWorkbook(filePath, sheetName, headers, normalizedRows);
    }

    private static string?[] ReadRow(IExcelDataReader reader)
    {
        var cells = new string?[reader.FieldCount];

        for (var index = 0; index < reader.FieldCount; index++)
        {
            cells[index] = ConvertCellValue(reader.GetValue(index));
        }

        return cells;
    }

    private static string? ConvertCellValue(object? value)
    {
        return value switch
        {
            null => null,
            string text => text.Trim(),
            DateTime dateTime when dateTime.TimeOfDay == TimeSpan.Zero => dateTime.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
            DateTime dateTime => dateTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
            bool boolean => boolean ? "TRUE" : "FALSE",
            IFormattable formattable => formattable.ToString(null, CultureInfo.InvariantCulture),
            _ => value.ToString()?.Trim()
        };
    }

    private static string[] CreateHeadersFromRow(IReadOnlyList<string?> cells)
    {
        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var headers = new string[cells.Count];

        for (var index = 0; index < cells.Count; index++)
        {
            var baseName = string.IsNullOrWhiteSpace(cells[index])
                ? ColumnReferenceHelper.ToGeneratedHeader(index)
                : cells[index]!.Trim();

            var candidate = baseName;
            var suffix = 2;
            while (!usedNames.Add(candidate))
            {
                candidate = $"{baseName}_{suffix++}";
            }

            headers[index] = candidate;
        }

        return headers;
    }

    private static string[] CreateGeneratedHeaders(IReadOnlyCollection<ImportedDataRow> rows)
    {
        var columnCount = rows.Count == 0 ? 0 : rows.Max(row => row.Cells.Count);
        return Enumerable.Range(0, columnCount)
            .Select(ColumnReferenceHelper.ToGeneratedHeader)
            .ToArray();
    }

    private static ImportedDataRow NormalizeRow(ImportedDataRow row, int expectedWidth)
    {
        if (row.Cells.Count == expectedWidth)
        {
            return row;
        }

        var normalized = new string?[expectedWidth];
        for (var index = 0; index < expectedWidth; index++)
        {
            normalized[index] = index < row.Cells.Count ? row.Cells[index] : null;
        }

        return new ImportedDataRow(row.RowNumber, normalized);
    }
}
