using ReportGenerator.Core.Models;

namespace ReportGenerator.Core.Abstractions;

public interface ITemplateValidationService
{
    ReportValidationResult ValidateTemplate(ReportTemplate template, ImportedWorkbook workbook);

    int ResolveColumnIndex(ImportedWorkbook workbook, ImportSettings settings, string source);
}
