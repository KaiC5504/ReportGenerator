using ReportGenerator.Core.Models;

namespace ReportGenerator.Core.Abstractions;

public interface IExcelImportService
{
    Task<ImportedWorkbook> ImportAsync(string filePath, ImportSettings settings, CancellationToken cancellationToken = default);
}
