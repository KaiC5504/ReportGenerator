using ReportGenerator.Core.Models;

namespace ReportGenerator.Core.Services;

public sealed class ReportValidationException : Exception
{
    public ReportValidationException(ReportValidationResult validationResult)
        : base("The report template is not valid for the selected workbook.")
    {
        ValidationResult = validationResult;
    }

    public ReportValidationResult ValidationResult { get; }
}
