using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Office365CleanupTool.Services
{
    public class OutlookDiagnosticsService
    {
        private const string SaraDownloadUrl = "https://aka.ms/SaRA_EnterpriseVersionFiles";
        private static readonly string[] SaraExecutableCandidates =
        {
            "SaRAcmd.exe",
            "GetHelpCmd.exe"
        };

        private static readonly TimeSpan ScenarioTimeout = TimeSpan.FromMinutes(15);
        private static readonly TimeSpan InteractiveTimeout = TimeSpan.FromMinutes(30);
        private static readonly TimeSpan CapabilityTimeout = TimeSpan.FromSeconds(20);
        private static readonly TimeSpan HarvestWaitTimeout = TimeSpan.FromSeconds(45);
        private static readonly TimeSpan HarvestPollInterval = TimeSpan.FromSeconds(3);

        public Task<OutlookDiagnosticResult> RunOutlookScanAsync(bool offlineScan, UiLanguage language = UiLanguage.Chinese)
        {
            string scenarioName = offlineScan
                ? T(language, "Outlook 离线扫描", "Outlook Offline Scan")
                : T(language, "Outlook 全量扫描", "Outlook Full Scan");
            const string argsTemplate = "-S ExpertExperienceAdminTask -AcceptEula -LogFolder \"{LOG}\"";
            return RunScenarioAsync("OutlookScan", scenarioName, argsTemplate, offlineScan, language);
        }

        public Task<OutlookDiagnosticResult> RunCalendarScanAsync(UiLanguage language = UiLanguage.Chinese)
        {
            const string argsTemplate = "-S OutlookCalendarCheckTask -AcceptEula -LogFolder \"{LOG}\"";
            return RunScenarioAsync(
                "OutlookCalendarScan",
                T(language, "Outlook 日历扫描", "Outlook Calendar Scan"),
                argsTemplate,
                false,
                language);
        }

        private async Task<OutlookDiagnosticResult> RunScenarioAsync(
            string scenarioKey,
            string scenarioDisplayName,
            string argumentsTemplate,
            bool requestOfflineScan,
            UiLanguage language)
        {
            ToolSessionPaths session = ToolPathService.CreateSession("OutlookTools", scenarioKey);
            string saraZipPath = Path.Combine(session.DownloadsDirectory, "SaRA.zip");
            string extractRoot = Path.Combine(session.ExtractedDirectory, "SaRA");
            string logRoot = session.LogsDirectory;
            DateTime startedAtUtc = DateTime.UtcNow;

            using (var client = new HttpClient())
            using (var response = await client.GetAsync(SaraDownloadUrl, HttpCompletionOption.ResponseHeadersRead))
            {
                response.EnsureSuccessStatusCode();
                await using var file = new FileStream(saraZipPath, FileMode.Create, FileAccess.Write, FileShare.None);
                await response.Content.CopyToAsync(file);
            }

            ZipFile.ExtractToDirectory(saraZipPath, extractRoot, true);

            string executablePath = FindSaraExecutable(extractRoot);
            GetHelpCapabilities caps = await QueryCapabilitiesAsync(executablePath);
            string arguments = BuildArguments(argumentsTemplate, logRoot, requestOfflineScan, caps);

            // Current GetHelpCmd builds in this environment may require a real console handle.
            // If -OfflineScan is not supported, run Outlook scan in interactive console mode directly.
            if (scenarioKey == "OutlookScan" && !caps.SupportsOfflineScanSwitch)
            {
                return await RunOutlookScanInteractiveAsync(
                    scenarioDisplayName,
                    executablePath,
                    arguments,
                    logRoot,
                    extractRoot,
                    caps,
                    requestOfflineScan,
                    startedAtUtc,
                    language);
            }

            var startInfo = new ProcessStartInfo
            {
                FileName = executablePath,
                Arguments = arguments,
                WorkingDirectory = Path.GetDirectoryName(executablePath),
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using Process process = Process.Start(startInfo)
                ?? throw new InvalidOperationException(T(language, "无法启动 Outlook 诊断命令。", "Failed to start Outlook diagnostics command."));

            Task feedTask = TryFeedContinueAsync(process, scenarioKey);
            Task<string> outputTask = process.StandardOutput.ReadToEndAsync();
            Task<string> errorTask = process.StandardError.ReadToEndAsync();
            Task waitTask = process.WaitForExitAsync();
            Task timeoutTask = Task.Delay(ScenarioTimeout);

            Task completed = await Task.WhenAny(waitTask, timeoutTask);
            if (completed != waitTask)
            {
                TryKill(process);
                string timeoutOutput = (await outputTask).Trim();
                string timeoutError = (await errorTask).Trim();
                return CreateTimeoutResult(
                    scenarioDisplayName,
                    logRoot,
                    extractRoot,
                    timeoutOutput,
                    timeoutError,
                    caps,
                    requestOfflineScan,
                    language);
            }

            await feedTask;
            string output = (await outputTask).Trim();
            string error = (await errorTask).Trim();
            int harvestedCount = await HarvestGeneratedLogsAsync(logRoot, startedAtUtc);
            if (harvestedCount > 0)
            {
                output = AppendAutomationNote(output, T(language, $"[M365Tool] 已自动归集日志文件：{harvestedCount} 个。", $"[M365Tool] Harvested log files automatically: {harvestedCount}."));
            }
            else
            {
                output = AppendAutomationNote(output, T(language, "[M365Tool] 当前会话目录未检测到新日志，已尝试回收常见路径。", "[M365Tool] No new logs detected in current session; attempted common-path harvest."));
            }

            string displayLogDirectory = ResolveDisplayLogDirectory(logRoot, startedAtUtc);

            int? detectedCode = scenarioKey == "OutlookScan"
                ? (DetectOutlookScanCode(output, error, process.ExitCode)
                    ?? DetectOutlookScanCodeFromHarvestedLogs(logRoot))
                : null;

            var result = new OutlookDiagnosticResult
            {
                ScenarioDisplayName = scenarioDisplayName,
                ExitCode = process.ExitCode,
                DetectedScenarioCode = detectedCode,
                StandardOutput = output,
                StandardError = error,
                LogDirectory = displayLogDirectory,
                RawLogRoot = extractRoot,
                SupportsOfflineSwitch = caps.SupportsOfflineScanSwitch,
                SupportsHideProgressSwitch = caps.SupportsHideProgressSwitch,
                RequestedOfflineScan = requestOfflineScan,
                HasArgumentError = ContainsInvalidArguments(output, error),
                HasInteractionError = ContainsInteractionFailure(output, error),
                UsedInteractiveConsole = false
            };

            result.Success = DetermineSuccess(result, scenarioKey);
            result.Message = BuildSummary(result, scenarioKey, language);
            return result;
        }

        public void OpenFolder(string path, UiLanguage language = UiLanguage.Chinese)
        {
            if (string.IsNullOrWhiteSpace(path) || !Directory.Exists(path))
            {
                throw new DirectoryNotFoundException(T(language, "目录不存在。", "Directory does not exist."));
            }

            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"\"{path}\"",
                UseShellExecute = true
            });
        }

        private async Task<OutlookDiagnosticResult> RunOutlookScanInteractiveAsync(
            string scenarioDisplayName,
            string executablePath,
            string arguments,
            string logRoot,
            string extractRoot,
            GetHelpCapabilities caps,
            bool requestOfflineScan,
            DateTime startedAtUtc,
            UiLanguage language)
        {
            InteractiveExecutionResult interactive = await RunInteractiveConsoleAsync(executablePath, arguments);
            if (interactive.TimedOut)
            {
                return new OutlookDiagnosticResult
                {
                    ScenarioDisplayName = scenarioDisplayName,
                    ExitCode = -3,
                    Success = false,
                    StandardOutput = interactive.Note,
                    StandardError = string.Empty,
                    LogDirectory = logRoot,
                    RawLogRoot = extractRoot,
                    SupportsOfflineSwitch = caps.SupportsOfflineScanSwitch,
                    SupportsHideProgressSwitch = caps.SupportsHideProgressSwitch,
                    RequestedOfflineScan = requestOfflineScan,
                    UsedInteractiveConsole = true,
                    Message = string.Join(
                        Environment.NewLine,
                        T(language, "执行结果", "Result"),
                        T(language, "Outlook 扫描交互控制台超时，已中止。", "Outlook scan interactive console timed out and was aborted."),
                        string.Empty,
                        T(language, "场景信息", "Scenario Info"),
                        $"• {T(language, "场景", "Scenario")}: {scenarioDisplayName}",
                        $"• {T(language, "超时时间", "Timeout")}: {InteractiveTimeout.TotalMinutes:0} {T(language, "分钟", "minutes")}",
                        $"• {T(language, "报告目录", "Report directory")}: {logRoot}",
                        $"• {T(language, "原始日志目录", "Raw log directory")}: {extractRoot}",
                        string.Empty,
                        T(language, "建议", "Suggestion"),
                        $"• {T(language, "请重试并在弹出的命令窗口中按提示输入 y。", "Retry and type y in the prompted command window.")}",
                        $"• {T(language, "若再次超时，请把截图发我继续定位。", "If timeout happens again, share screenshots for further analysis.")}")
                };
            }

            int harvestedCount = await HarvestGeneratedLogsAsync(logRoot, startedAtUtc);
            int? detectedCode = DetectOutlookScanCodeFromHarvestedLogs(logRoot)
                ?? DetectOutlookScanCode(string.Empty, string.Empty, interactive.ExitCode);

            string interactiveNote = interactive.Note;
            if (harvestedCount > 0)
            {
                interactiveNote = AppendAutomationNote(interactiveNote, T(language, $"[M365Tool] 已自动归集日志文件：{harvestedCount} 个。", $"[M365Tool] Harvested log files automatically: {harvestedCount}."));
            }
            else
            {
                interactiveNote = AppendAutomationNote(interactiveNote, T(language, "[M365Tool] 未发现新的诊断日志文件。", "[M365Tool] No new diagnostics log was found."));
            }

            var result = new OutlookDiagnosticResult
            {
                ScenarioDisplayName = scenarioDisplayName,
                ExitCode = interactive.ExitCode,
                DetectedScenarioCode = detectedCode,
                StandardOutput = interactiveNote,
                StandardError = string.Empty,
                LogDirectory = ResolveDisplayLogDirectory(logRoot, startedAtUtc),
                RawLogRoot = extractRoot,
                SupportsOfflineSwitch = caps.SupportsOfflineScanSwitch,
                SupportsHideProgressSwitch = caps.SupportsHideProgressSwitch,
                RequestedOfflineScan = requestOfflineScan,
                HasArgumentError = false,
                HasInteractionError = false,
                UsedInteractiveConsole = true
            };

            result.Success = DetermineSuccess(result, "OutlookScan");
            result.Message = BuildSummary(result, "OutlookScan", language);
            return result;
        }

        private static string AppendAutomationNote(string existingText, string note)
        {
            if (string.IsNullOrWhiteSpace(existingText))
            {
                return note;
            }

            return existingText + Environment.NewLine + Environment.NewLine + note;
        }

        private static async Task<InteractiveExecutionResult> RunInteractiveConsoleAsync(string executablePath, string arguments)
        {
            string workDir = Path.GetDirectoryName(executablePath) ?? Environment.CurrentDirectory;
            string scriptPath = Path.Combine(Path.GetTempPath(), $"M365Tool_GetHelp_{Guid.NewGuid():N}.cmd");
            string exitCodePath = Path.Combine(Path.GetTempPath(), $"M365Tool_GetHelp_Exit_{Guid.NewGuid():N}.txt");

            string scriptContent = string.Join(
                Environment.NewLine,
                "@echo off",
                "chcp 65001 >nul",
                "echo.",
                "echo [M365Tool] Outlook scan requires interactive confirmation on this GetHelpCmd build.",
                "echo [M365Tool] When prompted with \"Do you want to continue? (y/n):\", type y and press Enter.",
                "echo.",
                $"cd /d \"{workDir}\"",
                $"\"{Path.GetFileName(executablePath)}\" {arguments}",
                "set EXIT_CODE=%ERRORLEVEL%",
                $"(echo %EXIT_CODE%) > \"{exitCodePath}\"",
                "echo.",
                "echo [M365Tool] GetHelpCmd has exited. You can close this window.",
                "exit /b %EXIT_CODE%");

            File.WriteAllText(scriptPath, scriptContent, Encoding.ASCII);

            try
            {
                using Process process = Process.Start(new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = $"/c \"\"{scriptPath}\"\"",
                    UseShellExecute = true,
                    WorkingDirectory = workDir,
                    WindowStyle = ProcessWindowStyle.Normal
                }) ?? throw new InvalidOperationException("无法启动交互控制台。");

                Task waitTask = process.WaitForExitAsync();
                Task completed = await Task.WhenAny(waitTask, Task.Delay(InteractiveTimeout));
                if (completed != waitTask)
                {
                    TryKill(process);
                    return new InteractiveExecutionResult
                    {
                        ExitCode = -3,
                        TimedOut = true,
                        Note = "Interactive console timed out."
                    };
                }

                int exitCode = process.ExitCode;
                if (File.Exists(exitCodePath))
                {
                    string text = (File.ReadAllText(exitCodePath) ?? string.Empty).Trim();
                    if (int.TryParse(text, out int parsed))
                    {
                        exitCode = parsed;
                    }
                }

                return new InteractiveExecutionResult
                {
                    ExitCode = exitCode,
                    TimedOut = false,
                    Note = "Executed in interactive console mode."
                };
            }
            finally
            {
                TryDeleteFile(scriptPath);
                TryDeleteFile(exitCodePath);
            }
        }

        private static void TryDeleteFile(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
                // Ignore cleanup failures.
            }
        }

        private static async Task<int> HarvestGeneratedLogsAsync(string logRoot, DateTime startedAtUtc)
        {
            DateTime lowerBoundUtc = startedAtUtc.AddMinutes(-2);
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

            int copied = HarvestGeneratedLogsOnce(logRoot, lowerBoundUtc);

            if (copied == 0)
            {
                DateTime deadlineUtc = DateTime.UtcNow.Add(HarvestWaitTimeout);
                while (DateTime.UtcNow < deadlineUtc && copied == 0)
                {
                    await Task.Delay(HarvestPollInterval);
                    copied += HarvestGeneratedLogsOnce(logRoot, lowerBoundUtc);
                }
            }

            if (copied == 0)
            {
                // Fallback: copy latest files without strict time filter, to avoid empty folder UX.
                copied += CopyLatestFiles(
                    Path.Combine(localAppData, "GetHelp"),
                    Path.Combine(logRoot, "GetHelp_Fallback"),
                    startedAtUtc.AddHours(-6),
                    30);
                copied += CopyLatestFiles(
                    Path.Combine(localAppData, "Temp", "GetHelp"),
                    Path.Combine(logRoot, "TempGetHelp_Fallback"),
                    startedAtUtc.AddHours(-6),
                    30);
                copied += CopyLatestFiles(
                    Path.Combine(localAppData, "DAFs"),
                    Path.Combine(logRoot, "DAFs_Fallback"),
                    startedAtUtc.AddHours(-6),
                    20);
                copied += CopyLatestFiles(
                    Path.Combine(localAppData, "DiagSvc"),
                    Path.Combine(logRoot, "DiagSvc_Fallback"),
                    startedAtUtc.AddHours(-6),
                    20);
                copied += CopyLatestFiles(
                    Path.Combine(localAppData, "Temp", "Diagnostics", "OUTLOOK"),
                    Path.Combine(logRoot, "TempDiagnosticsOutlook_Fallback"),
                    startedAtUtc.AddHours(-6),
                    20);
            }

            EnsureLogFolderHasContent(logRoot);
            return copied;
        }

        private static int HarvestGeneratedLogsOnce(string logRoot, DateTime lowerBoundUtc)
        {
            int copied = 0;
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

            copied += CopyRecentFiles(
                Path.Combine(localAppData, "GetHelp"),
                Path.Combine(logRoot, "GetHelp"),
                lowerBoundUtc);

            copied += CopyRecentFiles(
                Path.Combine(localAppData, "Temp", "GetHelp"),
                Path.Combine(logRoot, "TempGetHelp"),
                lowerBoundUtc);

            copied += CopyRecentFiles(
                Path.Combine(localAppData, "DAFs"),
                Path.Combine(logRoot, "DAFs"),
                lowerBoundUtc);

            copied += CopyRecentFiles(
                Path.Combine(localAppData, "DiagSvc"),
                Path.Combine(logRoot, "DiagSvc"),
                lowerBoundUtc);

            copied += CopyRecentFiles(
                Path.Combine(localAppData, "SaRALogs", "Log"),
                Path.Combine(logRoot, "SaRALogs_Log"),
                lowerBoundUtc);

            copied += CopyRecentFiles(
                Path.Combine(localAppData, "SaRALogs", "UploadLogs"),
                Path.Combine(logRoot, "SaRALogs_UploadLogs"),
                lowerBoundUtc);

            copied += CopyRecentFiles(
                Path.Combine(localAppData, "SaraResults"),
                Path.Combine(logRoot, "SaraResults"),
                lowerBoundUtc);

            copied += CopyRecentFiles(
                Path.Combine(localAppData, "Temp", "Diagnostics", "OUTLOOK"),
                Path.Combine(logRoot, "TempDiagnosticsOutlook"),
                lowerBoundUtc);

            copied += HarvestFromResultXmlReportPath(
                Path.Combine(localAppData, "DiagSvc", "result.xml"),
                logRoot,
                lowerBoundUtc);

            copied += HarvestFromResultXmlReportPath(
                Path.Combine(logRoot, "DiagSvc", "result.xml"),
                logRoot,
                lowerBoundUtc);

            return copied;
        }

        private static int CopyRecentFiles(string sourceRoot, string targetRoot, DateTime lowerBoundUtc)
        {
            if (!Directory.Exists(sourceRoot))
            {
                return 0;
            }

            int copied = 0;
            int scanned = 0;
            const int maxScan = 2000;
            const int maxCopy = 400;

            IEnumerable<string> files;
            try
            {
                files = Directory.EnumerateFiles(sourceRoot, "*", SearchOption.AllDirectories);
            }
            catch
            {
                return 0;
            }

            foreach (string sourceFile in files)
            {
                if (scanned++ > maxScan || copied >= maxCopy)
                {
                    break;
                }

                FileInfo info;
                try
                {
                    info = new FileInfo(sourceFile);
                }
                catch
                {
                    continue;
                }

                if (info.LastWriteTimeUtc < lowerBoundUtc)
                {
                    continue;
                }

                if (!IsInterestingLogFile(info))
                {
                    continue;
                }

                string relativePath;
                try
                {
                    relativePath = Path.GetRelativePath(sourceRoot, sourceFile);
                }
                catch
                {
                    relativePath = info.Name;
                }

                string targetPath = Path.Combine(targetRoot, relativePath);
                string? targetDir = Path.GetDirectoryName(targetPath);
                if (!string.IsNullOrWhiteSpace(targetDir))
                {
                    Directory.CreateDirectory(targetDir);
                }

                targetPath = GetUniqueDestinationPath(targetPath);
                try
                {
                    File.Copy(sourceFile, targetPath, false);
                    copied++;
                }
                catch
                {
                    // Ignore copy failures and continue harvesting.
                }
            }

            return copied;
        }

        private static int CopyLatestFiles(string sourceRoot, string targetRoot, DateTime lowerBoundUtc, int take)
        {
            if (!Directory.Exists(sourceRoot))
            {
                return 0;
            }

            var candidates = new List<FileInfo>();
            try
            {
                foreach (string path in Directory.EnumerateFiles(sourceRoot, "*", SearchOption.AllDirectories))
                {
                    try
                    {
                        var info = new FileInfo(path);
                        if (info.LastWriteTimeUtc < lowerBoundUtc || !IsInterestingLogFile(info))
                        {
                            continue;
                        }

                        candidates.Add(info);
                    }
                    catch
                    {
                        // ignore invalid candidate
                    }
                }
            }
            catch
            {
                return 0;
            }

            int copied = 0;
            foreach (FileInfo info in candidates.OrderByDescending(item => item.LastWriteTimeUtc).Take(Math.Max(1, take)))
            {
                string relativePath;
                try
                {
                    relativePath = Path.GetRelativePath(sourceRoot, info.FullName);
                }
                catch
                {
                    relativePath = info.Name;
                }

                string targetPath = Path.Combine(targetRoot, relativePath);
                string? targetDir = Path.GetDirectoryName(targetPath);
                if (!string.IsNullOrWhiteSpace(targetDir))
                {
                    Directory.CreateDirectory(targetDir);
                }

                targetPath = GetUniqueDestinationPath(targetPath);
                try
                {
                    File.Copy(info.FullName, targetPath, false);
                    copied++;
                }
                catch
                {
                    // Ignore fallback copy failures.
                }
            }

            return copied;
        }

        private static int HarvestFromResultXmlReportPath(string resultXmlPath, string logRoot, DateTime lowerBoundUtc)
        {
            string? reportPath = TryReadReportPathFromResultXml(resultXmlPath);
            if (string.IsNullOrWhiteSpace(reportPath))
            {
                return 0;
            }

            string targetRoot = Path.Combine(logRoot, "ReportFile");
            int copied = 0;

            if (Directory.Exists(reportPath))
            {
                copied += CopyRecentFiles(reportPath, targetRoot, lowerBoundUtc);
                if (copied == 0)
                {
                    copied += CopyLatestFiles(reportPath, targetRoot, DateTime.UtcNow.AddDays(-30), 80);
                }

                return copied;
            }

            if (File.Exists(reportPath))
            {
                try
                {
                    Directory.CreateDirectory(targetRoot);
                    string targetPath = Path.Combine(targetRoot, Path.GetFileName(reportPath));
                    targetPath = GetUniqueDestinationPath(targetPath);
                    File.Copy(reportPath, targetPath, false);
                    return 1;
                }
                catch
                {
                    return 0;
                }
            }

            return 0;
        }

        private static string? TryReadReportPathFromResultXml(string resultXmlPath)
        {
            if (!File.Exists(resultXmlPath))
            {
                return null;
            }

            try
            {
                string xmlText = File.ReadAllText(resultXmlPath);
                Match m = Regex.Match(
                    xmlText,
                    "<Property\\s+Name\\s*=\\s*\"ReportFile\"\\s*>(?<path>[^<]+)</Property>",
                    RegexOptions.IgnoreCase);
                if (!m.Success)
                {
                    return null;
                }

                string path = m.Groups["path"].Value.Trim();
                return string.IsNullOrWhiteSpace(path) ? null : path;
            }
            catch
            {
                return null;
            }
        }

        private static bool IsInterestingLogFile(FileInfo info)
        {
            string ext = info.Extension.ToLowerInvariant();
            if (ext != ".log" &&
                ext != ".txt" &&
                ext != ".xml" &&
                ext != ".json" &&
                ext != ".html" &&
                ext != ".htm" &&
                ext != ".csv" &&
                ext != ".zip" &&
                ext != ".cab")
            {
                return false;
            }

            const long maxSize = 50L * 1024 * 1024;
            return info.Length <= maxSize;
        }

        private static string GetUniqueDestinationPath(string destinationPath)
        {
            if (!File.Exists(destinationPath))
            {
                return destinationPath;
            }

            string directory = Path.GetDirectoryName(destinationPath) ?? string.Empty;
            string name = Path.GetFileNameWithoutExtension(destinationPath);
            string ext = Path.GetExtension(destinationPath);
            string candidate = destinationPath;
            int i = 1;

            while (File.Exists(candidate))
            {
                candidate = Path.Combine(directory, $"{name}_{i}{ext}");
                i++;
            }

            return candidate;
        }

        private static void EnsureLogFolderHasContent(string logRoot)
        {
            bool hasEntries = DirectoryHasAnyVisibleEntry(logRoot);

            if (hasEntries)
            {
                return;
            }

            string readmePath = Path.Combine(logRoot, "README.txt");
            string content = string.Join(
                Environment.NewLine,
                "M365Tool did not find generated logs in expected locations.",
                "Try running the scenario again and keep the interactive console open until it exits.",
                "Then reopen this folder.");

            try
            {
                File.WriteAllText(readmePath, content, Encoding.UTF8);
            }
            catch
            {
                // Ignore write failures.
            }
        }

        private static int? DetectOutlookScanCodeFromHarvestedLogs(string logRoot)
        {
            if (!Directory.Exists(logRoot))
            {
                return null;
            }

            var candidateFiles = new List<FileInfo>();
            IEnumerable<string> paths;
            try
            {
                paths = Directory.EnumerateFiles(logRoot, "*", SearchOption.AllDirectories);
            }
            catch
            {
                return null;
            }

            foreach (string path in paths)
            {
                FileInfo info;
                try
                {
                    info = new FileInfo(path);
                }
                catch
                {
                    continue;
                }

                string ext = info.Extension.ToLowerInvariant();
                if (ext != ".log" && ext != ".txt" && ext != ".xml" && ext != ".json" && ext != ".html" && ext != ".htm")
                {
                    continue;
                }

                if (info.Length > 5L * 1024 * 1024)
                {
                    continue;
                }

                candidateFiles.Add(info);
            }

            candidateFiles.Sort((a, b) => b.LastWriteTimeUtc.CompareTo(a.LastWriteTimeUtc));
            int inspected = 0;

            foreach (FileInfo info in candidateFiles)
            {
                if (inspected++ >= 120)
                {
                    break;
                }

                string text;
                try
                {
                    text = File.ReadAllText(info.FullName);
                }
                catch
                {
                    continue;
                }

                int? code = DetectOutlookScanCode(text, string.Empty, 0);
                if (code.HasValue)
                {
                    return code;
                }
            }

            return null;
        }

        private static string ResolveDisplayLogDirectory(string logRoot, DateTime startedAtUtc)
        {
            string reportFileTarget = Path.Combine(logRoot, "ReportFile");
            if (DirectoryHasAnyVisibleEntry(reportFileTarget))
            {
                return reportFileTarget;
            }

            if (DirectoryHasAnyVisibleEntry(logRoot))
            {
                return logRoot;
            }

            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            DateTime lowerBoundUtc = startedAtUtc.AddHours(-6);
            string[] sourceRoots =
            {
                Path.Combine(localAppData, "GetHelp"),
                Path.Combine(localAppData, "Temp", "GetHelp"),
                Path.Combine(localAppData, "DAFs"),
                Path.Combine(localAppData, "DiagSvc"),
                Path.Combine(localAppData, "SaRALogs", "Log"),
                Path.Combine(localAppData, "SaRALogs", "UploadLogs"),
                Path.Combine(localAppData, "SaraResults"),
                Path.Combine(localAppData, "Temp", "Diagnostics", "OUTLOOK")
            };

            foreach (string root in sourceRoots)
            {
                if (HasRecentInterestingFile(root, lowerBoundUtc))
                {
                    return root;
                }
            }

            return logRoot;
        }

        private static bool DirectoryHasAnyVisibleEntry(string path)
        {
            if (!Directory.Exists(path))
            {
                return false;
            }

            try
            {
                foreach (string entry in Directory.EnumerateFileSystemEntries(path))
                {
                    var attrs = File.GetAttributes(entry);
                    bool isHidden = (attrs & FileAttributes.Hidden) == FileAttributes.Hidden
                        || (attrs & FileAttributes.System) == FileAttributes.System;
                    if (!isHidden)
                    {
                        return true;
                    }
                }
            }
            catch
            {
                return false;
            }

            return false;
        }

        private static bool HasRecentInterestingFile(string path, DateTime lowerBoundUtc)
        {
            if (!Directory.Exists(path))
            {
                return false;
            }

            try
            {
                foreach (string file in Directory.EnumerateFiles(path, "*", SearchOption.AllDirectories))
                {
                    try
                    {
                        var info = new FileInfo(file);
                        if (info.LastWriteTimeUtc >= lowerBoundUtc && IsInterestingLogFile(info))
                        {
                            return true;
                        }
                    }
                    catch
                    {
                        // ignore file-level exceptions
                    }
                }
            }
            catch
            {
                return false;
            }

            return false;
        }

        private static async Task<GetHelpCapabilities> QueryCapabilitiesAsync(string executablePath)
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = executablePath,
                Arguments = "-?",
                WorkingDirectory = Path.GetDirectoryName(executablePath),
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            try
            {
                using Process process = Process.Start(startInfo)
                    ?? throw new InvalidOperationException("无法读取 GetHelpCmd 功能列表。");

                Task<string> outputTask = process.StandardOutput.ReadToEndAsync();
                Task<string> errorTask = process.StandardError.ReadToEndAsync();
                Task waitTask = process.WaitForExitAsync();
                Task completed = await Task.WhenAny(waitTask, Task.Delay(CapabilityTimeout));

                if (completed != waitTask)
                {
                    TryKill(process);
                    return GetHelpCapabilities.None;
                }

                string helpText = (await outputTask) + Environment.NewLine + (await errorTask);
                return new GetHelpCapabilities
                {
                    SupportsOfflineScanSwitch = helpText.IndexOf("-OfflineScan", StringComparison.OrdinalIgnoreCase) >= 0,
                    SupportsHideProgressSwitch = helpText.IndexOf("-HideProgress", StringComparison.OrdinalIgnoreCase) >= 0
                };
            }
            catch
            {
                return GetHelpCapabilities.None;
            }
        }

        private static string BuildArguments(
            string argumentsTemplate,
            string logRoot,
            bool requestOfflineScan,
            GetHelpCapabilities caps)
        {
            string args = argumentsTemplate.Replace("{LOG}", logRoot, StringComparison.Ordinal);

            if (requestOfflineScan && caps.SupportsOfflineScanSwitch)
            {
                args += " -OfflineScan";
            }

            if (caps.SupportsHideProgressSwitch)
            {
                args += " -HideProgress";
            }

            return args;
        }

        private static async Task TryFeedContinueAsync(Process process, string scenarioKey)
        {
            if (!string.Equals(scenarioKey, "OutlookScan", StringComparison.Ordinal))
            {
                return;
            }

            try
            {
                await Task.Delay(1200);
                if (process.HasExited)
                {
                    return;
                }

                await process.StandardInput.WriteLineAsync("y");
                await process.StandardInput.FlushAsync();
                process.StandardInput.Close();
            }
            catch
            {
                // Ignore. Some builds don't accept redirected stdin.
            }
        }

        private static OutlookDiagnosticResult CreateTimeoutResult(
            string scenarioDisplayName,
            string logRoot,
            string extractRoot,
            string output,
            string error,
            GetHelpCapabilities caps,
            bool requestOfflineScan,
            UiLanguage language)
        {
            return new OutlookDiagnosticResult
            {
                ScenarioDisplayName = scenarioDisplayName,
                ExitCode = -2,
                Success = false,
                StandardOutput = output,
                StandardError = error,
                LogDirectory = logRoot,
                RawLogRoot = extractRoot,
                SupportsOfflineSwitch = caps.SupportsOfflineScanSwitch,
                SupportsHideProgressSwitch = caps.SupportsHideProgressSwitch,
                RequestedOfflineScan = requestOfflineScan,
                UsedInteractiveConsole = false,
                Message = BuildTimeoutSummary(scenarioDisplayName, logRoot, extractRoot, output, error, language)
            };
        }

        private static void TryKill(Process process)
        {
            try
            {
                if (!process.HasExited)
                {
                    process.Kill(true);
                }
            }
            catch
            {
                // Ignore kill failures.
            }
        }

        private static string FindSaraExecutable(string extractRoot)
        {
            foreach (string executableName in SaraExecutableCandidates)
            {
                string[] files = Directory.GetFiles(extractRoot, executableName, SearchOption.AllDirectories);
                if (files.Length > 0)
                {
                    return files[0];
                }
            }

            throw new FileNotFoundException(
                $"未在解压目录中找到 {string.Join(" 或 ", SaraExecutableCandidates)}。",
                extractRoot);
        }

        private static bool DetermineSuccess(OutlookDiagnosticResult result, string scenarioKey)
        {
            if (result.HasArgumentError || result.HasInteractionError)
            {
                return false;
            }

            if (scenarioKey == "OutlookScan")
            {
                int? code = result.DetectedScenarioCode;
                return code == 1 || code == 2;
            }

            if (scenarioKey == "OutlookCalendarScan")
            {
                return result.ExitCode == 0 || result.ExitCode == 43;
            }

            return result.ExitCode == 0;
        }

        private static int? DetectOutlookScanCode(string output, string error, int exitCode)
        {
            string all = (output ?? string.Empty) + Environment.NewLine + (error ?? string.Empty);

            Match m = Regex.Match(all, @"(?m)\b(0?[1-5])\s*:");
            if (m.Success && int.TryParse(m.Groups[1].Value, out int parsed))
            {
                return parsed;
            }

            if (all.IndexOf("An Offline scan was performed", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return 1;
            }

            if (all.IndexOf("A Full scan was performed", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return 2;
            }

            if (exitCode >= 1 && exitCode <= 5)
            {
                return exitCode;
            }

            return null;
        }

        private static bool ContainsInvalidArguments(string output, string error)
        {
            string all = (output ?? string.Empty) + Environment.NewLine + (error ?? string.Empty);
            return all.IndexOf("Invalid command line arguments", StringComparison.OrdinalIgnoreCase) >= 0
                || all.IndexOf("Usage: GetHelpCmd.exe", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool ContainsInteractionFailure(string output, string error)
        {
            string all = (output ?? string.Empty) + Environment.NewLine + (error ?? string.Empty);
            bool hasCatastrophic = all.IndexOf("A catastrophic error occurred", StringComparison.OrdinalIgnoreCase) >= 0;
            bool hasInvalidHandle = all.IndexOf("invalid handle", StringComparison.OrdinalIgnoreCase) >= 0
                || all.IndexOf("句柄无效", StringComparison.OrdinalIgnoreCase) >= 0;
            return hasCatastrophic || hasInvalidHandle;
        }

        private static string BuildTimeoutSummary(
            string scenarioDisplayName,
            string logRoot,
            string extractRoot,
            string output,
            string error,
            UiLanguage language)
        {
            var lines = new List<string>
            {
                T(language, "执行结果", "Result"),
                T(language, "Outlook 扫描超时并已中止。", "Outlook scan timed out and was aborted."),
                string.Empty,
                T(language, "场景信息", "Scenario Info"),
                $"• {T(language, "场景", "Scenario")}: {scenarioDisplayName}",
                $"• {T(language, "超时时间", "Timeout")}: {ScenarioTimeout.TotalMinutes:0} {T(language, "分钟", "minutes")}",
                $"• {T(language, "报告目录", "Report directory")}: {logRoot}",
                $"• {T(language, "原始日志目录", "Raw log directory")}: {extractRoot}"
            };

            AppendStream(lines, T(language, "标准输出", "Standard Output"), output);
            AppendStream(lines, T(language, "标准错误", "Standard Error"), error);
            return string.Join(Environment.NewLine, lines);
        }

        private static string BuildSummary(OutlookDiagnosticResult result, string scenarioKey, UiLanguage language)
        {
            if (result.HasArgumentError)
            {
                var lines = new List<string>
                {
                    T(language, "执行结果", "Result"),
                    T(language, "Outlook 扫描参数不兼容，场景未执行成功。", "Outlook scan parameters are incompatible and the scenario did not run."),
                    string.Empty,
                    T(language, "场景信息", "Scenario Info"),
                    $"• {T(language, "场景", "Scenario")}: {result.ScenarioDisplayName}",
                    $"• {T(language, "进程返回码", "Process exit code")}: {result.ExitCode}",
                    $"• {T(language, "报告目录", "Report directory")}: {result.LogDirectory}",
                    $"• {T(language, "原始日志目录", "Raw log directory")}: {result.RawLogRoot}",
                    string.Empty,
                    T(language, "建议", "Suggestion"),
                    $"• {T(language, "请确认下载的是官方最新 Command line Get Help。", "Ensure the downloaded Command line Get Help package is the latest official one.")}",
                    $"• {T(language, "如需我继续定位，请把标准输出/标准错误发给我。", "If you want further analysis, share standard output/error.")}"
                };

                AppendStream(lines, T(language, "标准输出", "Standard Output"), result.StandardOutput);
                AppendStream(lines, T(language, "标准错误", "Standard Error"), result.StandardError);
                return string.Join(Environment.NewLine, lines);
            }

            if (result.HasInteractionError)
            {
                var lines = new List<string>
                {
                    T(language, "执行结果", "Result"),
                    T(language, "Outlook 扫描触发了命令行交互，但当前会话返回了句柄错误，导致场景未完成。", "Outlook scan triggered CLI interaction, but the session returned an invalid-handle error."),
                    string.Empty,
                    T(language, "场景信息", "Scenario Info"),
                    $"• {T(language, "场景", "Scenario")}: {result.ScenarioDisplayName}",
                    $"• {T(language, "进程返回码", "Process exit code")}: {result.ExitCode}",
                    $"• {T(language, "报告目录", "Report directory")}: {result.LogDirectory}",
                    $"• {T(language, "原始日志目录", "Raw log directory")}: {result.RawLogRoot}",
                    string.Empty,
                    T(language, "说明", "Notes"),
                    $"• {T(language, "官方文档确实说明：Outlook 未运行时会自动执行离线扫描（01）。", "Official docs state that offline scan (01) should run automatically when Outlook is not running.")}",
                    $"• {T(language, "但当前下载到的 GetHelpCmd 在此会话里出现交互句柄异常（句柄无效），没有真正完成扫描。", "However, this GetHelpCmd build returned an invalid-handle interaction error and did not complete.")}"
                };

                if (result.RequestedOfflineScan && !result.SupportsOfflineSwitch)
                {
                    lines.Add($"• {T(language, "当前 GetHelpCmd 帮助信息未公开 -OfflineScan 开关，已按默认扫描命令执行。", "Current GetHelpCmd help does not expose -OfflineScan. Executed with default scan command.")}");
                }

                AppendStream(lines, T(language, "标准输出", "Standard Output"), result.StandardOutput);
                AppendStream(lines, T(language, "标准错误", "Standard Error"), result.StandardError);
                return string.Join(Environment.NewLine, lines);
            }

            if (scenarioKey == "OutlookCalendarScan")
            {
                return BuildCalendarSummary(result, language);
            }

            return BuildOutlookScanSummary(result, language);
        }

        private static string BuildOutlookScanSummary(OutlookDiagnosticResult result, UiLanguage language)
        {
            int code = result.DetectedScenarioCode ?? 0;
            string outcome = code switch
            {
                1 => T(language, "已执行 Outlook 离线扫描（01），请在日志目录查看报告。", "Outlook offline scan (01) completed. Check report directory."),
                2 => T(language, "已执行 Outlook 全量扫描（02），请在日志目录查看报告。", "Outlook full scan (02) completed. Check report directory."),
                4 => T(language, "返回 04：当前权限组合不符合场景要求（通常是命令窗口提升但 Outlook 未提升）。", "Returned 04: privilege combination does not meet scenario requirements."),
                5 => T(language, "返回 05：扫描未完成，建议关闭 Outlook 后重试。", "Returned 05: scan did not complete. Close Outlook and retry."),
                _ => T(language, "Outlook 扫描已执行，请查看日志目录。", "Outlook scan executed. Check log directory.")
            };

            var lines = new List<string>
            {
                T(language, "执行结果", "Result"),
                outcome,
                string.Empty,
                T(language, "场景信息", "Scenario Info"),
                $"• {T(language, "场景", "Scenario")}: {result.ScenarioDisplayName}",
                $"• {T(language, "进程返回码", "Process exit code")}: {result.ExitCode}",
                $"• {T(language, "场景结果码", "Scenario result code")}: {(result.DetectedScenarioCode.HasValue ? result.DetectedScenarioCode.Value.ToString("00") : T(language, "未识别", "Unknown"))}",
                $"• {T(language, "报告目录", "Report directory")}: {result.LogDirectory}",
                $"• {T(language, "原始日志目录", "Raw log directory")}: {result.RawLogRoot}"
            };

            if (result.UsedInteractiveConsole)
            {
                lines.Add($"• {T(language, "运行方式：交互控制台模式（按官方提示输入 y 后继续）。", "Run mode: interactive console (type y when prompted).")}");
            }

            if (result.RequestedOfflineScan && !result.SupportsOfflineSwitch)
            {
                lines.Add($"• {T(language, "备注：当前 GetHelpCmd 帮助信息未公开 -OfflineScan 开关，已按默认逻辑执行。", "Note: current GetHelpCmd help does not expose -OfflineScan; executed with default behavior.")}");
            }

            AppendStream(lines, T(language, "标准输出", "Standard Output"), result.StandardOutput);
            AppendStream(lines, T(language, "标准错误", "Standard Error"), result.StandardError);
            return string.Join(Environment.NewLine, lines);
        }

        private static string BuildCalendarSummary(OutlookDiagnosticResult result, UiLanguage language)
        {
            string outcome = result.ExitCode switch
            {
                0 => T(language, "Outlook 日历扫描已成功完成，请打开日志目录查看 CalCheck 结果。", "Outlook calendar scan completed. Open log folder for CalCheck result."),
                43 => T(language, "已完成 Outlook 日历扫描，并生成 CalCheck 结果。", "Outlook calendar scan completed and CalCheck results were generated."),
                40 => T(language, "未检测到已安装的 Outlook。", "No installed Outlook was detected."),
                41 => T(language, "未找到可用的 Outlook 配置文件。", "No valid Outlook profile was found."),
                42 => T(language, "Outlook 日历扫描执行失败。", "Outlook calendar scan failed."),
                44 => T(language, "指定的 Outlook 配置文件与当前运行配置文件不匹配。", "Specified Outlook profile does not match current running profile."),
                45 => T(language, "系统中未检测到可用的 Outlook 安装。", "No usable Outlook installation was detected."),
                _ => T(language, "Outlook 日历扫描未成功完成，请查看输出详情。", "Outlook calendar scan did not complete successfully. See details.")
            };

            var lines = new List<string>
            {
                T(language, "执行结果", "Result"),
                outcome,
                string.Empty,
                T(language, "场景信息", "Scenario Info"),
                $"• {T(language, "场景", "Scenario")}: {result.ScenarioDisplayName}",
                $"• {T(language, "进程返回码", "Process exit code")}: {result.ExitCode}",
                $"• {T(language, "报告目录", "Report directory")}: {result.LogDirectory}",
                $"• {T(language, "原始日志目录", "Raw log directory")}: {result.RawLogRoot}"
            };

            AppendStream(lines, T(language, "标准输出", "Standard Output"), result.StandardOutput);
            AppendStream(lines, T(language, "标准错误", "Standard Error"), result.StandardError);
            return string.Join(Environment.NewLine, lines);
        }

        private static void AppendStream(List<string> lines, string title, string content)
        {
            if (string.IsNullOrWhiteSpace(content))
            {
                return;
            }

            lines.Add(string.Empty);
            lines.Add(title);
            lines.Add(content);
        }

        private static string T(UiLanguage language, string zh, string en)
        {
            return LocalizationService.Localize(language, zh, en);
        }
    }

    public sealed class OutlookDiagnosticResult
    {
        public bool Success { get; set; }

        public bool HasArgumentError { get; set; }

        public bool HasInteractionError { get; set; }

        public bool SupportsOfflineSwitch { get; set; }

        public bool SupportsHideProgressSwitch { get; set; }

        public bool RequestedOfflineScan { get; set; }

        public bool UsedInteractiveConsole { get; set; }

        public string ScenarioDisplayName { get; set; } = string.Empty;

        public int ExitCode { get; set; }

        public int? DetectedScenarioCode { get; set; }

        public string StandardOutput { get; set; } = string.Empty;

        public string StandardError { get; set; } = string.Empty;

        public string LogDirectory { get; set; } = string.Empty;

        public string RawLogRoot { get; set; } = string.Empty;

        public string Message { get; set; } = string.Empty;
    }

    public sealed class GetHelpCapabilities
    {
        public static GetHelpCapabilities None { get; } = new();

        public bool SupportsOfflineScanSwitch { get; init; }

        public bool SupportsHideProgressSwitch { get; init; }
    }

    internal sealed class InteractiveExecutionResult
    {
        public int ExitCode { get; set; }

        public bool TimedOut { get; set; }

        public string Note { get; set; } = string.Empty;
    }
}
