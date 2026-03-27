using ReportGenerator.Core.Models;

namespace ReportGenerator.Core.Abstractions;

public interface IDocxExportService
{
    Task ExportAsync(PagedReport report, string outputPath, CancellationToken cancellationToken = default);
}
