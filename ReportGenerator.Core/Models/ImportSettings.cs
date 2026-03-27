namespace ReportGenerator.Core.Models;

public sealed class ImportSettings
{
    public ImportMappingMode MappingMode { get; set; } = ImportMappingMode.HeaderName;

    public bool HeaderRowEnabled { get; set; } = true;

    public bool FirstSheetOnly { get; set; } = true;
}
