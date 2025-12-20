using System.Collections.Concurrent;
using Topo.Services;

// A singleton service to hold the log queue.
public class LogQueueService
{
    public ConcurrentQueue<string> Queue { get; } = new();
}

// Topo Logger, which produces messages for the queue.
public class TopoLogger : ILogger
{
    private readonly string _categoryName;
    private readonly LogQueueService _logQueue;
    private readonly StorageService _storageService;

    public TopoLogger(string categoryName, LogQueueService logQueue, StorageService storageService)
    {
        _categoryName = categoryName;
        _logQueue = logQueue;
        _storageService = storageService;
    }


    public void Log<TState>(LogLevel logLevel, EventId eventId, TState state, Exception? exception, Func<TState, Exception?, string> formatter)
    {
        string? message;
#if DEBUG
        message = $"[{logLevel}] {_categoryName}: {formatter(state, exception)}";
        Console.WriteLine(message);
#endif

        if (!IsEnabled(logLevel))
        {
            return;
        }

        if (_storageService.LogToQueue)
        {
            // Formats the message and adds it to the queue. A fast and synchronous operation.
            message = $"[{logLevel}] {_categoryName}: {formatter(state, exception)}";
            _logQueue.Queue.Enqueue(message);
        }
    }

    public bool IsEnabled(LogLevel logLevel) => logLevel != LogLevel.None;

    public IDisposable BeginScope<TState>(TState state) => default!;
}

// Topo Provider, which creates instances of the Logger.
public class TopoLoggerProvider : ILoggerProvider
{
    private readonly LogQueueService _logQueue;
    private readonly StorageService _storageService;

    public TopoLoggerProvider(LogQueueService logQueue, StorageService storageService)
    {
        _logQueue = logQueue;
        _storageService = storageService;
    }

    public ILogger CreateLogger(string categoryName)
    {
        return new TopoLogger(categoryName, _logQueue, _storageService);
    }

    public void Dispose() { }
}
