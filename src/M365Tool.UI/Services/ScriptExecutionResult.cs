namespace Office365CleanupTool.Services
{
    public class ScriptExecutionResult
    {
        public int ExitCode { get; set; }

        public string StandardOutput { get; set; } = string.Empty;

        public string StandardError { get; set; } = string.Empty;

        public bool IsSuccess => ExitCode == 0;
    }
}
