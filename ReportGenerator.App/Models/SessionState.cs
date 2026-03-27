using ReportGenerator.Core.Models;

namespace ReportGenerator.App.Models;

public sealed class SessionState
{
    public string CurrentTemplatePath { get; set; } = "Unsaved template";

    public string? ExcelFilePath { get; set; }

    public bool IsEasyMode { get; set; }

    public ReportTemplate Template { get; set; } = new();
}
