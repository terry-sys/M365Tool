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

            request.CancellationToken.ThrowIfCancellationRequested();
            using var timeout = new System.Threading.CancellationTokenSource(request.Timeout);
            using var cancellation = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(
                request.CancellationToken, timeout.Token);
            var result = new ScriptExecutionResult();
            var output = new StringBuilder();
            var error = new StringBuilder();
            bool observing = true;
            bool detached = false;
            var process = new Process();
            try
            {
                process.StartInfo = new ProcessStartInfo
                {
                    FileName = request.FileName,
                    Arguments = request.Arguments,
                    UseShellExecute = request.UseShellExecute,
                    CreateNoWindow = request.CreateNoWindow,
                    WindowStyle = request.WindowStyle,
                    RedirectStandardOutput = !request.UseShellExecute && request.RedirectStandardOutput,
                    RedirectStandardError = !request.UseShellExecute && request.RedirectStandardError
                };
                if (!string.IsNullOrWhiteSpace(request.WorkingDirectory))
                {
                    process.StartInfo.WorkingDirectory = request.WorkingDirectory;
                }

                process.OutputDataReceived += (_, e) =>
                {
                    if (e.Data == null) return;
                    lock (output)
                    {
                        if (!System.Threading.Volatile.Read(ref observing)) return;
                        output.AppendLine(e.Data);
                        request.OnOutputDataReceived?.Invoke(e.Data);
                    }
                };
                process.ErrorDataReceived += (_, e) =>
                {
                    if (e.Data == null) return;
                    lock (error)
                    {
                        if (!System.Threading.Volatile.Read(ref observing)) return;
                        error.AppendLine(e.Data);
                        request.OnErrorDataReceived?.Invoke(e.Data);
                    }
                };
                request.CancellationToken.ThrowIfCancellationRequested();
                process.Start();
                result.ProcessId = process.Id;
                if (process.StartInfo.RedirectStandardOutput) process.BeginOutputReadLine();
                if (process.StartInfo.RedirectStandardError) process.BeginErrorReadLine();
                try
                {
                    await process.WaitForExitAsync(cancellation.Token);
                    result.ExitCode = process.ExitCode;
                }
                catch (OperationCanceledException)
                {
                    result.WasCanceled = request.CancellationToken.IsCancellationRequested;
                    result.TimedOut = !result.WasCanceled;
                    if (request.KillOnCancellation && !process.HasExited)
                    {
                        process.Kill(entireProcessTree: true);
                        await process.WaitForExitAsync();
                    }
                    result.IsStillRunning = !process.HasExited;
                    result.ExitCode = -1;
                    // Keep draining redirected pipes after monitoring stops, so a maintenance
                    // process can finish without broken pipes or callbacks to a closed page.
                    System.Threading.Volatile.Write(ref observing, false);
                    detached = true;
                    Task completion = DisposeAfterExitAsync(process);
                    MaintenanceCoordinator.Shared.TrackBackgroundTask(completion);
                }
                lock (output) result.StandardOutput = output.ToString();
                lock (error) result.StandardError = error.ToString();
                if (result.TimedOut || result.WasCanceled)
                {
                    result.StandardError += Environment.NewLine +
                        (result.TimedOut ? "Monitoring timed out." : "Monitoring was canceled.") +
                        (result.IsStillRunning ? $" Process {result.ProcessId} is still running." : string.Empty);
                }
                return result;
            }
            finally
            {
                if (!detached) process.Dispose();
            }
        }

        private static async Task DisposeAfterExitAsync(Process process)
        {
            try { await process.WaitForExitAsync(); }
            catch (InvalidOperationException) { }
            finally { process.Dispose(); }
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
