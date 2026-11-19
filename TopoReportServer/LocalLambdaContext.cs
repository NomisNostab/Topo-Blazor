using Amazon.Lambda.Core;

namespace TopoReportServer;

/// <summary>
/// The context AWS hands to a Lambda function. The report function never reads it,
/// so this only has to exist; it fills in plausible values in case that changes.
/// </summary>
public class LocalLambdaContext : ILambdaContext
{
    public string AwsRequestId { get; } = Guid.NewGuid().ToString();
    public IClientContext ClientContext => null!;
    public string FunctionName => "TopoReportGenerator (local)";
    public string FunctionVersion => "local";
    public ICognitoIdentity Identity => null!;
    public string InvokedFunctionArn => "arn:aws:lambda:local:000000000000:function:TopoReportGenerator";
    public ILambdaLogger Logger { get; } = new ConsoleLambdaLogger();
    public string LogGroupName => "local";
    public string LogStreamName => "local";
    public int MemoryLimitInMB => 1024;
    public TimeSpan RemainingTime => TimeSpan.FromMinutes(5);

    private sealed class ConsoleLambdaLogger : ILambdaLogger
    {
        public void Log(string message) => Console.Write(message);
        public void LogLine(string message) => Console.WriteLine(message);
    }
}
