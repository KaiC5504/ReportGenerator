using Velopack;

namespace ReportGenerator.App.Models;

public sealed class UpdateCheckResult
{
    public bool IsInstalledApp { get; init; }

    public bool IsUpdateAvailable { get; init; }

    public bool IsTimedOut { get; init; }

    public bool IsError { get; init; }

    public string Message { get; init; } = string.Empty;

    public string AvailableVersion { get; init; } = string.Empty;

    public string ReleaseNotes { get; init; } = string.Empty;

    public UpdateInfo? UpdateInfo { get; init; }
}
