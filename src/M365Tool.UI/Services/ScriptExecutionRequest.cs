using System;
using System.Diagnostics;
using System.Threading;

namespace Office365CleanupTool.Services
{
    public class ScriptExecutionRequest
    {
        public string FileName { get; set; } = string.Empty;

        public string Arguments { get; set; } = string.Empty;

        public CancellationToken CancellationToken { get; set; }

        public TimeSpan Timeout { get; set; } = TimeSpan.FromMinutes(30);

        public bool KillOnCancellation { get; set; }

        public string? WorkingDirectory { get; set; }

        public bool UseShellExecute { get; set; }

        public bool CreateNoWindow { get; set; } = true;

        public bool RedirectStandardOutput { get; set; } = true;

        public bool RedirectStandardError { get; set; } = true;

        public ProcessWindowStyle WindowStyle { get; set; } = ProcessWindowStyle.Normal;

        public Action<string>? OnOutputDataReceived { get; set; }

        public Action<string>? OnErrorDataReceived { get; set; }
    }
}
