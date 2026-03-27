using System.Collections.ObjectModel;

namespace ReportGenerator.Core.Models;

public sealed class ReportValidationResult
{
    private readonly List<ReportValidationIssue> _issues = [];

    public bool IsValid => _issues.Count == 0;

    public IReadOnlyList<ReportValidationIssue> Issues => new ReadOnlyCollection<ReportValidationIssue>(_issues);

    public void AddIssue(string field, string message)
    {
        _issues.Add(new ReportValidationIssue(field, message));
    }
}
