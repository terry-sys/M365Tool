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

            string saraArchivePath = ArchiveSaraLogs(extractRoot, logRoot);

            var result = new OfficeUninstallResult
            {
                ExitCode = executionResult.ExitCode,
                Summary = SaraResultInterpreter.GetSummary(executionResult.ExitCode, language),
                Details = BuildDetails(request.Target, executionResult.ExitCode, saraArchivePath, processCleanupSummary, detectedOfficeSummary, language),
                NextStepAdvice = SaraResultInterpreter.GetNextStepAdvice(executionResult.ExitCode, request.Target, detectedOfficeSummary, language),
                DetectedOfficeSummary = detectedOfficeSummary,
                LogDirectory = logRoot,
                StandardOutput = executionResult.StandardOutput,
                StandardError = executionResult.StandardError,
                ProcessCleanupSummary = processCleanupSummary
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
            UiLanguage language)
        {
            string targetText = target.UsesDetectedInstalledVersion
                ? T(language, "自动检测", "Auto-detect")
                : target.DisplayName;
            return
                $"{T(language, "卸载目标", "Uninstall target")}: {targetText}\r\n" +
                $"{T(language, "结果码", "Exit code")}: {exitCode}\r\n" +
                $"{T(language, "说明", "Summary")}: {SaraResultInterpreter.GetSummary(exitCode, language)}\r\n" +
                $"{T(language, "当前检测", "Detected")}: {detectedOfficeSummary}\r\n" +
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

        private static string T(UiLanguage language, string zh, string en)
        {
            return LocalizationService.Localize(language, zh, en);
        }
    }
}
