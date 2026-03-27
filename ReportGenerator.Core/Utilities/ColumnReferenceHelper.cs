using System.Text;

namespace ReportGenerator.Core.Utilities;

public static class ColumnReferenceHelper
{
    public static string ToColumnLetter(int zeroBasedIndex)
    {
        if (zeroBasedIndex < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(zeroBasedIndex));
        }

        var index = zeroBasedIndex + 1;
        var builder = new StringBuilder();

        while (index > 0)
        {
            index--;
            builder.Insert(0, (char)('A' + (index % 26)));
            index /= 26;
        }

        return builder.ToString();
    }

    public static bool TryParseColumnLetter(string? input, out int zeroBasedIndex)
    {
        zeroBasedIndex = -1;

        if (string.IsNullOrWhiteSpace(input))
        {
            return false;
        }

        var normalized = input.Trim().ToUpperInvariant();
        if (normalized.Any(c => !char.IsAsciiLetter(c)))
        {
            return false;
        }

        var result = 0;
        foreach (var character in normalized)
        {
            result = (result * 26) + (character - 'A' + 1);
        }

        zeroBasedIndex = result - 1;
        return true;
    }

    public static string ToGeneratedHeader(int zeroBasedIndex)
    {
        return $"Column {ToColumnLetter(zeroBasedIndex)}";
    }
}
