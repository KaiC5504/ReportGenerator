using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using ReportGenerator.App.Models;
using ReportGenerator.Core.Services;

namespace ReportGenerator.App.Services;

public sealed class SessionStateStore
{
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
        Converters =
        {
            new JsonStringEnumConverter(JsonNamingPolicy.CamelCase)
        }
    };

    private readonly string _sessionStatePath;

    public SessionStateStore(string? sessionStatePath = null)
    {
        _sessionStatePath = sessionStatePath ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ReportGenerator",
            "config",
            "session-state.json");
    }

    public SessionState? Load()
    {
        if (!File.Exists(_sessionStatePath))
        {
            return null;
        }

        using var stream = File.OpenRead(_sessionStatePath);
        var sessionState = JsonSerializer.Deserialize<SessionState>(stream, SerializerOptions);
        if (sessionState is null)
        {
            return null;
        }

        sessionState.Template = TemplateStorageService.NormalizeTemplate(sessionState.Template);
        sessionState.CurrentTemplatePath = string.IsNullOrWhiteSpace(sessionState.CurrentTemplatePath)
            ? "Unsaved template"
            : sessionState.CurrentTemplatePath;
        sessionState.ExcelFilePath = string.IsNullOrWhiteSpace(sessionState.ExcelFilePath)
            ? null
            : sessionState.ExcelFilePath;

        return sessionState;
    }

    public void Save(SessionState sessionState)
    {
        ArgumentNullException.ThrowIfNull(sessionState);
        ArgumentNullException.ThrowIfNull(sessionState.Template);

        var directory = Path.GetDirectoryName(_sessionStatePath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var normalizedState = new SessionState
        {
            CurrentTemplatePath = string.IsNullOrWhiteSpace(sessionState.CurrentTemplatePath)
                ? "Unsaved template"
                : sessionState.CurrentTemplatePath,
            ExcelFilePath = string.IsNullOrWhiteSpace(sessionState.ExcelFilePath)
                ? null
                : sessionState.ExcelFilePath,
            IsEasyMode = sessionState.IsEasyMode,
            Template = TemplateStorageService.NormalizeTemplate(sessionState.Template)
        };

        using var stream = File.Create(_sessionStatePath);
        JsonSerializer.Serialize(stream, normalizedState, SerializerOptions);
    }
}
