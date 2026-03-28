using ReportGenerator.Core.Abstractions;
using ReportGenerator.Core.Models;
using ReportGenerator.Core.Utilities;

namespace ReportGenerator.Core.Services;

public sealed class TemplateValidationService : ITemplateValidationService
{
    public ReportValidationResult ValidateTemplate(ReportTemplate template, ImportedWorkbook workbook)
    {
        ArgumentNullException.ThrowIfNull(template);
        ArgumentNullException.ThrowIfNull(workbook);

        var result = new ReportValidationResult();

        if (template.PageSettings.RowsPerPage <= 0)
        {
            result.AddIssue("pageSettings.rowsPerPage", "Rows per page must be at least 1.");
        }

        if (template.DetailTable.Columns.Count == 0)
        {
            result.AddIssue("detailTable.columns", "At least one detail column is required.");
        }

        if (template.DetailTable.ContentFontSize < 8 || template.DetailTable.ContentFontSize > 24)
        {
            result.AddIssue("detailTable.contentFontSize", "Content font size must be between 8 and 24.");
        }

        if (template.DetailTable.ContentRowSpacing < 0)
        {
            result.AddIssue("detailTable.contentRowSpacing", "Content spacing must be 0 or greater.");
        }
        else if (!HasAtMostTwoDecimalPlaces(template.DetailTable.ContentRowSpacing))
        {
            result.AddIssue("detailTable.contentRowSpacing", "Content spacing must use at most 2 decimal places.");
        }

        if (template.DetailTable.HeaderFontSize < 8 || template.DetailTable.HeaderFontSize > 24)
        {
            result.AddIssue("detailTable.headerFontSize", "Header font size must be between 8 and 24.");
        }

        if (template.DetailTable.GroupEveryRows < 0)
        {
            result.AddIssue("detailTable.groupEveryRows", "Group by must be 0 or greater.");
        }

        if (template.DetailTable.GroupSpacingRows < 0)
        {
            result.AddIssue("detailTable.groupSpacingRows", "Group spacing must be 0 or greater.");
        }
        else if (!HasAtMostTwoDecimalPlaces(template.DetailTable.GroupSpacingRows))
        {
            result.AddIssue("detailTable.groupSpacingRows", "Group spacing must use at most 2 decimal places.");
        }

        for (var index = 0; index < template.DetailTable.Columns.Count; index++)
        {
            var column = template.DetailTable.Columns[index];
            if (string.IsNullOrWhiteSpace(column.HeaderText))
            {
                result.AddIssue($"detailTable.columns[{index}].headerText", "Column header text is required.");
            }

            if (string.IsNullOrWhiteSpace(column.Source))
            {
                result.AddIssue($"detailTable.columns[{index}].source", "Column source is required.");
            }
            else if (!TryResolveColumnIndex(workbook, template.ImportSettings, column.Source, out _))
            {
                result.AddIssue(
                    $"detailTable.columns[{index}].source",
                    $"Column source '{column.Source}' does not exist in the selected worksheet.");
            }
        }

        ValidateBlocks(result, workbook, template.ImportSettings, template.HeaderBlocks, "headerBlocks", validateRowLayout: true);
        ValidateBlocks(result, workbook, template.ImportSettings, template.FooterBlocks, "footerBlocks", validateRowLayout: false);

        return result;
    }

    public int ResolveColumnIndex(ImportedWorkbook workbook, ImportSettings settings, string source)
    {
        if (!TryResolveColumnIndex(workbook, settings, source, out var index))
        {
            throw new InvalidOperationException($"Unable to resolve source '{source}' using mapping mode '{settings.MappingMode}'.");
        }

        return index;
    }

    private static void ValidateBlocks(
        ReportValidationResult result,
        ImportedWorkbook workbook,
        ImportSettings settings,
        IReadOnlyList<ReportBlock> blocks,
        string fieldName,
        bool validateRowLayout)
    {
        var firstPageVisibilityByRow = new Dictionary<int, bool>();
        var alignmentsByRow = new Dictionary<int, HashSet<ReportTextAlignment>>();

        for (var index = 0; index < blocks.Count; index++)
        {
            var block = blocks[index];
            if (block.Row < 1)
            {
                result.AddIssue($"{fieldName}[{index}].row", "Row must be 1 or greater.");
            }

            if (validateRowLayout && block.Row >= 1)
            {
                if (!firstPageVisibilityByRow.TryAdd(block.Row, block.OnlyOnFirstPage)
                    && firstPageVisibilityByRow[block.Row] != block.OnlyOnFirstPage)
                {
                    result.AddIssue(
                        $"{fieldName}[{index}].onlyOnFirstPage",
                        "All blocks in the same row must share the same Only On First Page setting.");
                }

                if (!alignmentsByRow.TryGetValue(block.Row, out var alignments))
                {
                    alignments = [];
                    alignmentsByRow[block.Row] = alignments;
                }

                if (!alignments.Add(block.Alignment))
                {
                    result.AddIssue(
                        $"{fieldName}[{index}].alignment",
                        "Each header row can contain at most one Left, Center, and Right block.");
                }
            }

            if (block.Type == ReportBlockType.MappedField && string.IsNullOrWhiteSpace(block.Source))
            {
                result.AddIssue($"{fieldName}[{index}].source", "Mapped field blocks require a source.");
                continue;
            }

            if (block.Type == ReportBlockType.MappedField
                && !TryResolveColumnIndex(workbook, settings, block.Source, out _))
            {
                result.AddIssue(
                    $"{fieldName}[{index}].source",
                    $"Mapped field source '{block.Source}' does not exist in the selected worksheet.");
            }
        }
    }

    private static bool TryResolveColumnIndex(ImportedWorkbook workbook, ImportSettings settings, string? source, out int index)
    {
        index = -1;
        if (string.IsNullOrWhiteSpace(source))
        {
            return false;
        }

        return settings.MappingMode switch
        {
            ImportMappingMode.HeaderName => ResolveByHeaderName(workbook, source, out index),
            ImportMappingMode.ColumnLetter => ResolveByColumnLetter(workbook, source, out index),
            _ => false
        };
    }

    private static bool HasAtMostTwoDecimalPlaces(double value)
    {
        return Math.Abs(value - Math.Round(value, 2, MidpointRounding.AwayFromZero)) < 0.000001d;
    }

    private static bool ResolveByHeaderName(ImportedWorkbook workbook, string source, out int index)
    {
        index = workbook.Headers
            .Select((header, headerIndex) => new { header, headerIndex })
            .FirstOrDefault(item => string.Equals(item.header, source.Trim(), StringComparison.OrdinalIgnoreCase))
            ?.headerIndex ?? -1;

        return index >= 0;
    }

    private static bool ResolveByColumnLetter(ImportedWorkbook workbook, string source, out int index)
    {
        if (!ColumnReferenceHelper.TryParseColumnLetter(source, out index))
        {
            return false;
        }

        return index < workbook.ColumnCount;
    }
}
