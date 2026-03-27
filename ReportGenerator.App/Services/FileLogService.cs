using System.IO;
using System.Text;

namespace ReportGenerator.App.Services;

public sealed class FileLogService
{
    private readonly string _logDirectoryPath;

    public FileLogService()
    {
        _logDirectoryPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ReportGenerator",
            "logs");
    }

    public string LogFilePath => Path.Combine(_logDirectoryPath, "app.log");

    public void LogInfo(string message)
    {
        WriteEntry("INFO", message, null);
    }

    public void LogError(string message, Exception? exception = null)
    {
        WriteEntry("ERROR", message, exception);
    }

    private void WriteEntry(string level, string message, Exception? exception)
    {
        Directory.CreateDirectory(_logDirectoryPath);

        var builder = new StringBuilder()
            .Append('[')
            .Append(DateTimeOffset.Now.ToString("yyyy-MM-dd HH:mm:ss"))
            .Append("] [")
            .Append(level)
            .Append("] ")
            .AppendLine(message);

        if (exception is not null)
        {
            builder.AppendLine(exception.ToString());
        }

        File.AppendAllText(LogFilePath, builder.ToString());
    }
}
