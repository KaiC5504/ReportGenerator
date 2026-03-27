namespace ReportGenerator.App.Models;

public readonly record struct PreviewGenerationProgress(int CurrentPage, int TotalPages)
{
    public double Percentage => TotalPages <= 0
        ? 0
        : (CurrentPage * 100d) / TotalPages;
}
