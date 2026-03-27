namespace ReportGenerator.App.Models;

public sealed class UpdateSettings
{
    public string FeedUrl { get; set; } = "https://updates.kaic5504.com/reportGenerator";

    public int StartupDelaySeconds { get; set; } = 2;

    public int RequestTimeoutSeconds { get; set; } = 5;

    public string CurrentChannel { get; set; } = "win-x64-stable";

    public DateTimeOffset? LastCheckUtc { get; set; }
}
