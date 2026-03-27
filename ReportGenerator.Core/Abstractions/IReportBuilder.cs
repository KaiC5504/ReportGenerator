using ReportGenerator.Core.Models;

namespace ReportGenerator.Core.Abstractions;

public interface IReportBuilder
{
    PagedReport Build(ReportTemplate template, ImportedWorkbook workbook);
}
