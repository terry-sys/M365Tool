using System;
using System.Diagnostics;
using System.Text;
using System.Threading.Tasks;

namespace Office365CleanupTool.Services
{
    public class ScriptRunner : IScriptRunner
    {
        public async Task<ScriptExecutionResult> RunAsync(ScriptExecutionRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.FileName))
            {
                throw new ArgumentException("FileName cannot be empty.", nameof(request));
            }

            var result = new ScriptExecutionResult();
            var standardOutput = new StringBuilder();
            var standardError = new StringBuilder();

            using (var process = new Process())
            {
                process.StartInfo.FileName = request.FileName;
                process.StartInfo.Arguments = request.Arguments;
                process.StartInfo.UseShellExecute = request.UseShellExecute;
                process.StartInfo.CreateNoWindow = request.CreateNoWindow;
                process.StartInfo.WindowStyle = request.WindowStyle;

                if (!string.IsNullOrWhiteSpace(request.WorkingDirectory))
                {
                    process.StartInfo.WorkingDirectory = request.WorkingDirectory;
                }

                var canRedirect = !request.UseShellExecute;
                process.StartInfo.RedirectStandardOutput = canRedirect && request.RedirectStandardOutput;
                process.StartInfo.RedirectStandardError = canRedirect && request.RedirectStandardError;

                if (process.StartInfo.RedirectStandardOutput)
                {
                    process.OutputDataReceived += (_, e) =>
                    {
                        if (string.IsNullOrEmpty(e.Data))
                        {
                            return;
                        }

                        standardOutput.AppendLine(e.Data);
                        request.OnOutputDataReceived?.Invoke(e.Data);
                    };
                }

                if (process.StartInfo.RedirectStandardError)
                {
                    process.ErrorDataReceived += (_, e) =>
                    {
                        if (string.IsNullOrEmpty(e.Data))
                        {
                            return;
                        }

                        standardError.AppendLine(e.Data);
                        request.OnErrorDataReceived?.Invoke(e.Data);
                    };
                }

                process.Start();

                if (process.StartInfo.RedirectStandardOutput)
                {
                    process.BeginOutputReadLine();
                }

                if (process.StartInfo.RedirectStandardError)
                {
                    process.BeginErrorReadLine();
                }

                await process.WaitForExitAsync();
                result.ExitCode = process.ExitCode;
            }

            result.StandardOutput = standardOutput.ToString();
            result.StandardError = standardError.ToString();
            return result;
        }

        public void StartDetached(ScriptExecutionRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.FileName))
            {
                throw new ArgumentException("FileName cannot be empty.", nameof(request));
            }

            var startInfo = new ProcessStartInfo
            {
                FileName = request.FileName,
                Arguments = request.Arguments,
                UseShellExecute = request.UseShellExecute,
                CreateNoWindow = request.CreateNoWindow,
                WindowStyle = request.WindowStyle
            };

            if (!string.IsNullOrWhiteSpace(request.WorkingDirectory))
            {
                startInfo.WorkingDirectory = request.WorkingDirectory;
            }

            Process.Start(startInfo);
        }
    }
}
