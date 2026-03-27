using ReportGenerator.Core.Models;

namespace ReportGenerator.Core.Abstractions;

public interface IPdfExportService
{
    Task ExportAsync(PagedReport report, string outputPath, CancellationToken cancellationToken = default);
}
