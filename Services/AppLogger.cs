using Serilog;
using Serilog.Core;

namespace Cafe24ShipmentManager.Services;

public class AppLogger
{
    private readonly Logger _fileLogger;
    public event Action<string>? OnLog;

    public AppLogger(string logDirectory)
    {
        Directory.CreateDirectory(logDirectory);
        _fileLogger = new LoggerConfiguration()
            .WriteTo.File(
                Path.Combine(logDirectory, "app-.log"),
                rollingInterval: RollingInterval.Day,
                outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
            .CreateLogger();
    }

    public void Info(string message)
    {
        _fileLogger.Information(message);
        OnLog?.Invoke($"[INFO] {DateTime.Now:HH:mm:ss} {message}");
    }

    public void Warn(string message)
    {
        _fileLogger.Warning(message);
        OnLog?.Invoke($"[WARN] {DateTime.Now:HH:mm:ss} {message}");
    }

    public void Error(string message, Exception? ex = null)
    {
        _fileLogger.Error(ex, message);
        var logMsg = $"[ERROR] {DateTime.Now:HH:mm:ss} {message}";
        if (ex != null) logMsg += $"\n  → {ex.Message}\n  → {ex.StackTrace}";
        OnLog?.Invoke(logMsg);
    }
}
