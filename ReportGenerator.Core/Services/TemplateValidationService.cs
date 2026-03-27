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

        if (template.DetailTable.GroupEveryRows < 0)
        {
            result.AddIssue("detailTable.groupEveryRows", "Group by must be 0 or greater.");
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

        ValidateBlocks(result, workbook, template.ImportSettings, template.HeaderBlocks, "headerBlocks");
        ValidateBlocks(result, workbook, template.ImportSettings, template.FooterBlocks, "footerBlocks");

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
        string fieldName)
    {
        for (var index = 0; index < blocks.Count; index++)
        {
            var block = blocks[index];
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
