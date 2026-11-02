namespace Office365CleanupTool.Services
{
    public class ScriptExecutionResult
    {
        public int ExitCode { get; set; }

        public string StandardOutput { get; set; } = string.Empty;

        public string StandardError { get; set; } = string.Empty;

        public bool TimedOut { get; set; }

        public bool WasCanceled { get; set; }

        public bool IsStillRunning { get; set; }

        public int ProcessId { get; set; }

        public bool IsSuccess => ExitCode == 0 && !TimedOut && !WasCanceled && !IsStillRunning;
    }
}
