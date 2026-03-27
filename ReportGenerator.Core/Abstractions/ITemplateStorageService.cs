using ReportGenerator.Core.Models;

namespace ReportGenerator.Core.Abstractions;

public interface ITemplateStorageService
{
    ReportTemplate CreateDefaultTemplate();

    Task SaveAsync(ReportTemplate template, string filePath, CancellationToken cancellationToken = default);

    Task<ReportTemplate> LoadAsync(string filePath, CancellationToken cancellationToken = default);
}
