using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;
using Office365CleanupTool.Models;

namespace Office365CleanupTool.Services
{
    public enum OfficeUninstallStatus
    {
        Processing,
        Success,
        Error
    }

    public class OfficeUninstallProgressEventArgs : EventArgs
    {
        public OfficeUninstallProgressEventArgs(string message, OfficeUninstallStatus status)
        {
            Message = message;
            Status = status;
        }

        public string Message { get; }

        public OfficeUninstallStatus Status { get; }
    }

    public class InstalledOfficeInfo
    {
        public string DisplayName { get; set; } = string.Empty;

        public string Version { get; set; } = string.Empty;

        public override string ToString() =>
            string.IsNullOrWhiteSpace(Version) ? DisplayName : $"{DisplayName} ({Version})";
    }

    public class OfficeUninstallService
    {
        private static readonly string[] SaraExecutableCandidates =
        {
            "SaRAcmd.exe",
            "GetHelpCmd.exe"
        };

        private static readonly string[] OfficeProcessNames =
        {
            "lync", "winword", "excel", "msaccess", "mstore", "infopath", "setlang",
            "msouc", "ois", "onenote", "outlook", "powerpnt", "mspub", "groove",
            "visio", "winproj", "graph", "teams"
        };

        private const string SaraDownloadUrl = "https://aka.ms/SaRA_EnterpriseVersionFiles";
        private static readonly TimeSpan BackgroundUninstallTimeout = TimeSpan.FromMinutes(30);
        private static readonly TimeSpan BackgroundUninstallPollInterval = TimeSpan.FromSeconds(15);
        private readonly IScriptRunner _scriptRunner;
        private readonly OfficeInventoryService _officeInventoryService;

        public OfficeUninstallService(IScriptRunner scriptRunner)
        {
            _scriptRunner = scriptRunner;
            _officeInventoryService = new OfficeInventoryService();
        }

        public event EventHandler<OfficeUninstallProgressEventArgs>? ProgressChanged;

        public IReadOnlyList<OfficeUninstallTargetOption> GetSupportedTargets(UiLanguage language = UiLanguage.Chinese)
        {
            return new[]
            {
                new OfficeUninstallTargetOption
                {
                    DisplayName = T(language, "自动检测已安装的 Office", "Auto-detect installed Office"),
                    Description = T(language, "让微软官方 SaRA 自动检测当前设备上的 Office 版本并尝试卸载。若设备上同时存在多个版本，SaRA 会返回多版本结果码。", "Use official SaRA to auto-detect installed Office versions and attempt uninstall. If multiple versions exist, SaRA returns multi-version result codes."),
                    SaraOfficeVersion = null
                },
                new OfficeUninstallTargetOption
                {
                    DisplayName = "Microsoft 365 Apps",
                    Description = T(language, "卸载订阅版 Office，例如 Microsoft 365 Apps for enterprise / business。", "Uninstall subscription Office, such as Microsoft 365 Apps for enterprise / business."),
                    SaraOfficeVersion = "M365"
                },
                new OfficeUninstallTargetOption
                {
                    DisplayName = "Office 2021",
                    Description = T(language, "仅卸载 Office 2021。", "Uninstall Office 2021 only."),
                    SaraOfficeVersion = "2021"
                },
                new OfficeUninstallTargetOption
                {
                    DisplayName = "Office 2019",
                    Description = T(language, "仅卸载 Office 2019。", "Uninstall Office 2019 only."),
                    SaraOfficeVersion = "2019"
                },
                new OfficeUninstallTargetOption
                {
                    DisplayName = "Office 2016",
                    Description = T(language, "仅卸载 Office 2016。", "Uninstall Office 2016 only."),
                    SaraOfficeVersion = "2016"
                },
                new OfficeUninstallTargetOption
                {
                    DisplayName = "Office 2013",
                    Description = T(language, "仅卸载 Office 2013。", "Uninstall Office 2013 only."),
                    SaraOfficeVersion = "2013"
                },
                new OfficeUninstallTargetOption
                {
                    DisplayName = "Office 2010",
                    Description = T(language, "仅卸载 Office 2010。", "Uninstall Office 2010 only."),
                    SaraOfficeVersion = "2010"
                },
                new OfficeUninstallTargetOption
                {
                    DisplayName = "Office 2007",
                    Description = T(language, "仅卸载 Office 2007。", "Uninstall Office 2007 only."),
                    SaraOfficeVersion = "2007"
                },
                new OfficeUninstallTargetOption
                {
                    DisplayName = T(language, "所有 Office 版本", "All Office Versions"),
                    Description = T(language, "调用 SaRA 的 All 模式，尝试移除设备上的所有 Office 版本。", "Use SaRA All mode to remove all Office versions from this device."),
                    SaraOfficeVersion = "All"
                }
            };
        }

        public IReadOnlyList<InstalledOfficeInfo> DetectInstalledOffice()
        {
            return _officeInventoryService
                .DetectInstalledProducts()
                .Select(item => new InstalledOfficeInfo
                {
                    DisplayName = item.DisplayName,
                    Version = item.Version
                })
                .ToList();
        }

        public async Task<OfficeUninstallResult> RunAsync(OfficeUninstallRequest request, UiLanguage language = UiLanguage.Chinese)
        {
            if (request.Target == null)
            {
                throw new ArgumentException(T(language, "必须指定卸载目标。", "Uninstall target is required."), nameof(request));
            }

            var detectedOffice = DetectInstalledOffice();
            string detectedOfficeSummary = detectedOffice.Count == 0
                ? T(language, "未检测到已安装的 Office。", "No installed Office detected.")
                : string.Join(language == UiLanguage.Chinese ? "；" : "; ", detectedOffice.Select(item => item.ToString()));

            ToolSessionPaths session = CreateSessionPaths();
            string downloadZipPath = Path.Combine(session.DownloadsDirectory, "SaRA.zip");
            string extractRoot = Path.Combine(session.ExtractedDirectory, "SaRA");
            string logRoot = session.LogsDirectory;
            string executionLogPath = Path.Combine(logRoot, "execution.log");
            Directory.CreateDirectory(logRoot);

            OnProgressChanged(T(language, "正在准备新版 Office 卸载引擎...", "Preparing latest Office uninstall engine..."), OfficeUninstallStatus.Processing);
            string processCleanupSummary = await CloseOfficeProcessesAsync(language);

            await DownloadSaraAsync(downloadZipPath, language);

            OnProgressChanged(T(language, "正在解压微软官方 SaRA 工具...", "Extracting official Microsoft SaRA package..."), OfficeUninstallStatus.Processing);
            ZipFile.ExtractToDirectory(downloadZipPath, extractRoot, true);

            string saraCmdPath = FindSaraCmdPath(extractRoot, language);
            string arguments = BuildArguments(request.Target);

            OnProgressChanged(T(language, "正在启动 SaRA 卸载场景...", "Launching SaRA uninstall scenario..."), OfficeUninstallStatus.Processing);
            var executionResult = await _scriptRunner.RunAsync(new ScriptExecutionRequest
            {
                FileName = saraCmdPath,
                Arguments = arguments,
                WorkingDirectory = Path.GetDirectoryName(saraCmdPath),
                OnOutputDataReceived = data => OnProgressChanged($"{T(language, "SaRA 输出", "SaRA Output")}: {data}", OfficeUninstallStatus.Processing),
                OnErrorDataReceived = data => OnProgressChanged($"{T(language, "SaRA 错误", "SaRA Error")}: {data}", OfficeUninstallStatus.Processing)
            });

            bool backgroundUninstallReported = IndicatesBackgroundUninstall(executionResult);
            IReadOnlyList<InstalledOfficeInfo> postCheckOffice = backgroundUninstallReported
                ? await WaitForBackgroundUninstallAsync(language)
                : DetectInstalledOffice();
            string postCheckOfficeSummary = postCheckOffice.Count == 0
                ? T(language, "未检测到已安装的 Office。", "No installed Office detected.")
                : string.Join(language == UiLanguage.Chinese ? "；" : "; ", postCheckOffice.Select(item => item.ToString()));
            bool stillHasOfficeAfterUninstall = postCheckOffice.Count > 0;
            bool requiresAttention = executionResult.ExitCode != 0 || (backgroundUninstallReported && stillHasOfficeAfterUninstall);
            string summary = BuildSummary(executionResult.ExitCode, backgroundUninstallReported, stillHasOfficeAfterUninstall, language);
            string nextStepAdvice = BuildNextStepAdvice(
                executionResult.ExitCode,
                request.Target,
                detectedOfficeSummary,
                postCheckOfficeSummary,
                backgroundUninstallReported,
                stillHasOfficeAfterUninstall,
                language);
            string saraArchivePath = ArchiveSaraLogs(extractRoot, logRoot);

            var result = new OfficeUninstallResult
            {
                ExitCode = executionResult.ExitCode,
                Summary = summary,
                Details = BuildDetails(
                    request.Target,
                    executionResult.ExitCode,
                    saraArchivePath,
                    processCleanupSummary,
                    detectedOfficeSummary,
                    postCheckOfficeSummary,
                    backgroundUninstallReported,
                    stillHasOfficeAfterUninstall,
                    language),
                NextStepAdvice = nextStepAdvice,
                DetectedOfficeSummary = detectedOfficeSummary,
                PostCheckOfficeSummary = postCheckOfficeSummary,
                LogDirectory = logRoot,
                StandardOutput = executionResult.StandardOutput,
                StandardError = executionResult.StandardError,
                ProcessCleanupSummary = processCleanupSummary,
                BackgroundUninstallReported = backgroundUninstallReported,
                RequiresAttention = requiresAttention
            };

            File.WriteAllText(
                executionLogPath,
                BuildExecutionLog(request.Target, arguments, result, saraArchivePath),
                Encoding.UTF8);

            OnProgressChanged(result.Summary, result.IsSuccess ? OfficeUninstallStatus.Success : OfficeUninstallStatus.Error);
            return result;
        }

        private static ToolSessionPaths CreateSessionPaths()
        {
            return ToolPathService.CreateSession("OfficeUninstall");
        }

        private async Task<string> CloseOfficeProcessesAsync(UiLanguage language)
        {
            return await Task.Run(() =>
            {
                var closedProcesses = new List<string>();
                var skippedProcesses = new List<string>();

                OnProgressChanged(T(language, "正在检查并结束 Office 相关进程...", "Checking and closing Office related processes..."), OfficeUninstallStatus.Processing);

                foreach (string processName in OfficeProcessNames)
                {
                    foreach (var process in Process.GetProcessesByName(processName))
                    {
                        try
                        {
                            string displayName = $"{process.ProcessName}({process.Id})";

                            if (!process.HasExited)
                            {
                                OnProgressChanged($"{T(language, "正在结束进程", "Closing process")}: {displayName}", OfficeUninstallStatus.Processing);
                                process.Kill(entireProcessTree: true);
                                process.WaitForExit(5000);
                            }

                            closedProcesses.Add(displayName);
                        }
                        catch
                        {
                            skippedProcesses.Add($"{process.ProcessName}({process.Id})");
                        }
                        finally
                        {
                            process.Dispose();
                        }
                    }
                }

                if (closedProcesses.Count == 0 && skippedProcesses.Count == 0)
                {
                    OnProgressChanged(T(language, "未发现正在运行的 Office 相关进程。", "No running Office related processes were found."), OfficeUninstallStatus.Processing);
                    return T(language, "未发现正在运行的 Office 相关进程。", "No running Office related processes were found.");
                }

                string closedSummary = closedProcesses.Count == 0
                    ? T(language, "未主动结束任何进程", "No process was actively terminated")
                    : $"{T(language, "已结束进程", "Closed processes")}: {string.Join(", ", closedProcesses)}";

                if (skippedProcesses.Count == 0)
                {
                    OnProgressChanged(closedSummary, OfficeUninstallStatus.Processing);
                    return closedSummary;
                }

                string skippedSummary = $"{(language == UiLanguage.Chinese ? "；" : "; ")}{T(language, "未能结束", "Failed to close")}: {string.Join(", ", skippedProcesses)}";
                OnProgressChanged(closedSummary + skippedSummary, OfficeUninstallStatus.Processing);
                return closedSummary + skippedSummary;
            });
        }

        private async Task DownloadSaraAsync(string zipPath, UiLanguage language)
        {
            OnProgressChanged(T(language, "正在下载微软官方 SaRA 企业版...", "Downloading official Microsoft SaRA enterprise package..."), OfficeUninstallStatus.Processing);

            using (var client = new HttpClient())
            using (var response = await client.GetAsync(SaraDownloadUrl, HttpCompletionOption.ResponseHeadersRead))
            {
                response.EnsureSuccessStatusCode();
                await using (var fileStream = new FileStream(zipPath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    await response.Content.CopyToAsync(fileStream);
                }
            }
        }

        private static string FindSaraCmdPath(string extractRoot, UiLanguage language)
        {
            foreach (string executableName in SaraExecutableCandidates)
            {
                string[] candidates = Directory.GetFiles(extractRoot, executableName, SearchOption.AllDirectories);
                if (candidates.Length > 0)
                {
                    return candidates[0];
                }
            }

            string[] discoveredFiles = Directory
                .GetFiles(extractRoot, "*", SearchOption.AllDirectories)
                .Select(path => Path.GetRelativePath(extractRoot, path))
                .Take(30)
                .ToArray();

            string discoveredSummary = discoveredFiles.Length == 0
                ? T(language, "解压目录为空。", "Extract directory is empty.")
                : T(language, "已发现文件", "Discovered files") + ": " + string.Join(", ", discoveredFiles);

            throw new FileNotFoundException(
                T(
                    language,
                    $"未在解压目录中找到 {string.Join(" 或 ", SaraExecutableCandidates)}。{discoveredSummary}",
                    $"Could not find {string.Join(" or ", SaraExecutableCandidates)} in extracted directory. {discoveredSummary}"),
                extractRoot);
        }

        private static string BuildArguments(OfficeUninstallTargetOption target)
        {
            if (target.UsesDetectedInstalledVersion)
            {
                return "-S OfficeScrubScenario -AcceptEula";
            }

            return $"-S OfficeScrubScenario -AcceptEula -OfficeVersion {target.SaraOfficeVersion}";
        }

        private static string ArchiveSaraLogs(string extractRoot, string logRoot)
        {
            string archivePath = Path.Combine(logRoot, "SaRA_Files.zip");
            if (File.Exists(archivePath))
            {
                File.Delete(archivePath);
            }

            ZipFile.CreateFromDirectory(extractRoot, archivePath);
            return archivePath;
        }

        private static string BuildDetails(
            OfficeUninstallTargetOption target,
            int exitCode,
            string saraArchivePath,
            string processCleanupSummary,
            string detectedOfficeSummary,
            string postCheckOfficeSummary,
            bool backgroundUninstallReported,
            bool stillHasOfficeAfterUninstall,
            UiLanguage language)
        {
            string targetText = target.UsesDetectedInstalledVersion
                ? T(language, "自动检测", "Auto-detect")
                : target.DisplayName;
            return
                $"{T(language, "卸载目标", "Uninstall target")}: {targetText}\r\n" +
                $"{T(language, "结果码", "Exit code")}: {exitCode}\r\n" +
                $"{T(language, "说明", "Summary")}: {BuildSummary(exitCode, backgroundUninstallReported, stillHasOfficeAfterUninstall, language)}\r\n" +
                $"{T(language, "卸载前检测", "Before uninstall")}: {detectedOfficeSummary}\r\n" +
                $"{T(language, "卸载后复查", "Post-check")}: {postCheckOfficeSummary}\r\n" +
                $"{T(language, "进程处理", "Process cleanup")}: {processCleanupSummary}\r\n" +
                $"{T(language, "日志压缩包", "Log archive")}: {saraArchivePath}";
        }

        private static string BuildExecutionLog(
            OfficeUninstallTargetOption target,
            string arguments,
            OfficeUninstallResult result,
            string saraArchivePath)
        {
            var builder = new StringBuilder();
            builder.AppendLine($"Time: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            builder.AppendLine($"Target: {target.DisplayName}");
            builder.AppendLine($"Arguments: {arguments}");
            builder.AppendLine($"ExitCode: {result.ExitCode}");
            builder.AppendLine($"Summary: {result.Summary}");
            builder.AppendLine($"DetectedOffice: {result.DetectedOfficeSummary}");
            builder.AppendLine($"PostCheckOffice: {result.PostCheckOfficeSummary}");
            builder.AppendLine($"BackgroundUninstallReported: {result.BackgroundUninstallReported}");
            builder.AppendLine($"RequiresAttention: {result.RequiresAttention}");
            builder.AppendLine($"NextStepAdvice: {result.NextStepAdvice}");
            builder.AppendLine($"ProcessCleanup: {result.ProcessCleanupSummary}");
            builder.AppendLine($"SaraArchive: {saraArchivePath}");
            builder.AppendLine();
            builder.AppendLine("[StandardOutput]");
            builder.AppendLine(result.StandardOutput);
            builder.AppendLine();
            builder.AppendLine("[StandardError]");
            builder.AppendLine(result.StandardError);
            return builder.ToString();
        }

        private void OnProgressChanged(string message, OfficeUninstallStatus status)
        {
            ProgressChanged?.Invoke(this, new OfficeUninstallProgressEventArgs(message, status));
        }

        private async Task<IReadOnlyList<InstalledOfficeInfo>> WaitForBackgroundUninstallAsync(UiLanguage language)
        {
            OnProgressChanged(
                T(
                    language,
                    "GetHelp 已将 Office 卸载转入后台执行，正在等待卸载完成并复查控制面板残留项...",
                    "GetHelp started Office uninstall in the background. Waiting for completion and checking remaining entries..."),
                OfficeUninstallStatus.Processing);

            DateTime deadline = DateTime.Now.Add(BackgroundUninstallTimeout);
            IReadOnlyList<InstalledOfficeInfo> current = DetectInstalledOffice();

            while (current.Count > 0 && DateTime.Now < deadline)
            {
                OnProgressChanged(
                    T(
                        language,
                        $"后台卸载仍在进行或等待系统刷新，当前仍检测到 {current.Count} 项 Office 组件。请不要关闭可能出现的卸载窗口。",
                        $"Background uninstall is still running or waiting for system refresh. {current.Count} Office entries are still detected. Do not close any uninstall window that appears."),
                    OfficeUninstallStatus.Processing);

                await Task.Delay(BackgroundUninstallPollInterval);
                current = DetectInstalledOffice();
            }

            return current;
        }

        private static bool IndicatesBackgroundUninstall(ScriptExecutionResult result)
        {
            string combinedOutput = $"{result.StandardOutput}\n{result.StandardError}";
            return combinedOutput.Contains("Uninstall office is running in the background", StringComparison.OrdinalIgnoreCase)
                || combinedOutput.Contains("Office uninstall is running in the background", StringComparison.OrdinalIgnoreCase)
                || combinedOutput.Contains("running in the background", StringComparison.OrdinalIgnoreCase);
        }

        private static string BuildSummary(
            int exitCode,
            bool backgroundUninstallReported,
            bool stillHasOfficeAfterUninstall,
            UiLanguage language)
        {
            if (exitCode == 0 && backgroundUninstallReported && stillHasOfficeAfterUninstall)
            {
                return T(
                    language,
                    "卸载已由 GetHelp 转入后台执行，但复查时仍检测到 Office 组件，不能判定为已完成。",
                    "GetHelp started uninstall in the background, but Office entries are still detected during post-check. Completion cannot be confirmed.");
            }

            if (exitCode == 0 && backgroundUninstallReported)
            {
                return T(
                    language,
                    "后台卸载已完成，复查未检测到 Office 组件，建议重启计算机完成剩余清理。",
                    "Background uninstall completed. No Office entries were detected during post-check. Restart is recommended.");
            }

            return SaraResultInterpreter.GetSummary(exitCode, language);
        }

        private static string BuildNextStepAdvice(
            int exitCode,
            OfficeUninstallTargetOption target,
            string detectedOfficeSummary,
            string postCheckOfficeSummary,
            bool backgroundUninstallReported,
            bool stillHasOfficeAfterUninstall,
            UiLanguage language)
        {
            if (exitCode == 0 && backgroundUninstallReported && stillHasOfficeAfterUninstall)
            {
                return T(
                    language,
                    $"请先等待后台卸载窗口结束或重启计算机，然后重新打开控制面板检查残留项。当前复查仍检测到：{postCheckOfficeSummary}",
                    $"Wait for the background uninstall window to finish or restart the computer, then check Control Panel again. Still detected: {postCheckOfficeSummary}");
            }

            return SaraResultInterpreter.GetNextStepAdvice(exitCode, target, detectedOfficeSummary, language);
        }

        private static string T(UiLanguage language, string zh, string en)
        {
            return LocalizationService.Localize(language, zh, en);
        }
    }
}
