using ReportGenerator.App.Models;
using Velopack;
using Velopack.Sources;

namespace ReportGenerator.App.Services;

public sealed class UpdateService
{
    private readonly FileLogService _logService;
    private readonly UpdateSettingsStore _settingsStore;
    private UpdateSettings? _cachedSettings;

    public UpdateService(FileLogService logService, UpdateSettingsStore settingsStore)
    {
        _logService = logService;
        _settingsStore = settingsStore;
    }

    public async Task<UpdateSettings> GetSettingsAsync(CancellationToken cancellationToken = default)
    {
        _cachedSettings ??= await _settingsStore.LoadAsync(cancellationToken);
        return _cachedSettings;
    }

    public async Task<UpdateCheckResult> CheckForUpdatesAsync(bool userInitiated, CancellationToken cancellationToken = default)
    {
        try
        {
            var settings = await GetSettingsAsync(cancellationToken);
            var manager = CreateUpdateManager(settings);

            if (!manager.IsInstalled)
            {
                return new UpdateCheckResult
                {
                    IsInstalledApp = false,
                    Message = userInitiated
                        ? "Updates are available only for the installed application package."
                        : "Auto update check skipped because the app is not installed yet."
                };
            }

            var timeoutTask = Task.Delay(TimeSpan.FromSeconds(settings.RequestTimeoutSeconds), cancellationToken);
            var updateTask = manager.CheckForUpdatesAsync();
            var completedTask = await Task.WhenAny(updateTask, timeoutTask);

            if (completedTask != updateTask)
            {
                return new UpdateCheckResult
                {
                    IsInstalledApp = true,
                    IsTimedOut = true,
                    Message = userInitiated
                        ? "The update server did not respond in time."
                        : "Automatic update check timed out."
                };
            }

            var updateInfo = await updateTask;
            settings.LastCheckUtc = DateTimeOffset.UtcNow;
            await _settingsStore.SaveAsync(settings, cancellationToken);

            if (updateInfo is null)
            {
                return new UpdateCheckResult
                {
                    IsInstalledApp = true,
                    Message = "You are already on the latest version."
                };
            }

            return new UpdateCheckResult
            {
                IsInstalledApp = true,
                IsUpdateAvailable = true,
                AvailableVersion = updateInfo.TargetFullRelease.Version.ToString(),
                ReleaseNotes = updateInfo.TargetFullRelease.NotesMarkdown ?? string.Empty,
                UpdateInfo = updateInfo,
                Message = $"Version {updateInfo.TargetFullRelease.Version} is available."
            };
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception)
        {
            _logService.LogError("Update check failed.", exception);
            return new UpdateCheckResult
            {
                IsError = true,
                Message = userInitiated
                    ? "Unable to reach the update server right now."
                    : "Automatic update check skipped because the server is unavailable."
            };
        }
    }

    public async Task DownloadAndApplyUpdateAsync(
        UpdateCheckResult updateResult,
        Action<int>? progressCallback,
        CancellationToken cancellationToken = default)
    {
        if (updateResult.UpdateInfo is null)
        {
            throw new InvalidOperationException("No update payload is available to install.");
        }

        var settings = await GetSettingsAsync(cancellationToken);
        var manager = CreateUpdateManager(settings);
        await manager.DownloadUpdatesAsync(updateResult.UpdateInfo, progressCallback, cancellationToken);
        _logService.LogInfo($"Downloaded update {updateResult.AvailableVersion}.");
        await manager.WaitExitThenApplyUpdatesAsync(updateResult.UpdateInfo.TargetFullRelease, false, true, []);
    }

    private static UpdateManager CreateUpdateManager(UpdateSettings settings)
    {
        var timeoutMinutes = Math.Max(0.05, settings.RequestTimeoutSeconds / 60d);
        var source = new SimpleWebSource(settings.FeedUrl, null, timeoutMinutes);
        return new UpdateManager(source, new UpdateOptions
        {
            ExplicitChannel = settings.CurrentChannel
        });
    }
}
