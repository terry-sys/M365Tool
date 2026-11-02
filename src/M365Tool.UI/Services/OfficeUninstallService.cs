using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Xml.Linq;
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

        public string ProductId { get; set; } = string.Empty;

        public string OfficeVersion { get; set; } = string.Empty;

        public string Source { get; set; } = string.Empty;

        public override string ToString() =>
            string.IsNullOrWhiteSpace(Version) ? DisplayName : $"{DisplayName} ({Version})";
    }

    public class OfficeUninstallService
    {
        private static readonly string[] SaraExecutableCandidates =
        {
            "GetHelpCmd.exe",
            "SaRAcmd.exe"
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

        public string LastLogDirectory { get; private set; } = string.Empty;

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
                    DisplayName = "Office 2024",
                    Description = T(language, "仅卸载 Office 2024。", "Uninstall Office 2024 only."),
                    SaraOfficeVersion = "2024"
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

        public IReadOnlyList<OfficeUninstallTargetOption> GetAvailableTargets(
            OfficeUninstallMode mode, IReadOnlyList<InstalledOfficeInfo> installedOffice,
            UiLanguage language = UiLanguage.Chinese)
        {
            var supported = GetSupportedTargets(language);
            if (mode == OfficeUninstallMode.DeepCleanup) return supported;

            var c2rProducts = GetPostCheckScope(installedOffice, OfficeUninstallMode.Standard)
                .Where(item => !string.IsNullOrWhiteSpace(item.ProductId)).ToList();
            if (c2rProducts.Count == 0) return Array.Empty<OfficeUninstallTargetOption>();

            // An explicit All option also covers detected C2R products with an unknown version.
            // The removal builder still limits this option to the detected C2R product IDs.
            return supported.Where(option => !option.UsesDetectedInstalledVersion
                && (option.SaraOfficeVersion == "All"
                    || c2rProducts.Any(product => IsUninstallTarget(product, option)))).ToList();
        }

        public virtual IReadOnlyList<InstalledOfficeInfo> DetectInstalledOffice()
        {
            return _officeInventoryService
                .DetectInstalledProducts()
                .Select(item => new InstalledOfficeInfo
                {
                    DisplayName = item.DisplayName,
                    Version = item.Version,
                    ProductId = item.ProductId,
                    OfficeVersion = item.OfficeVersion,
                    Source = item.Source
                })
                .ToList();
        }

        public async Task<OfficeUninstallResult> RunAsync(OfficeUninstallRequest request, UiLanguage language = UiLanguage.Chinese,
            CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (request.Target == null)
            {
                throw new ArgumentException(T(language, "必须指定卸载目标。", "Uninstall target is required."), nameof(request));
            }

            using var operation = MaintenanceCoordinator.Shared.Begin(language);
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

            LastLogDirectory = logRoot;
            using var sessionLog = new StreamWriter(executionLogPath, false, Encoding.UTF8) { AutoFlush = true };
            void RecordProgress(object? sender, OfficeUninstallProgressEventArgs e)
            {
                lock (sessionLog) sessionLog.WriteLine($"[{DateTime.Now:HH:mm:ss}] {e.Message}");
            }
            ProgressChanged += RecordProgress;
            try
            {
                string saraCmdPath;
                string arguments;
                if (request.Mode == OfficeUninstallMode.Standard)
                {
                    XDocument removal = BuildStandardRemovalXml(request.Target, detectedOffice);
                    saraCmdPath = await new OdtToolService(_scriptRunner).PrepareAsync(extractRoot, cancellationToken,
                        message => OnProgressChanged(message, OfficeUninstallStatus.Processing), language);
                    string xmlPath = Path.Combine(logRoot, "OfficeRemove.xml");
                    removal.Save(xmlPath);
                    arguments = $"/configure \"{xmlPath}\"";
                }
                else
                {
                    await DownloadSaraAsync(downloadZipPath, language, cancellationToken);
                    OnProgressChanged(T(language, "正在解压 Get Help...", "Extracting Get Help..."), OfficeUninstallStatus.Processing);
                    ZipFile.ExtractToDirectory(downloadZipPath, extractRoot, true);
                    saraCmdPath = FindSaraCmdPath(extractRoot, language);
                    OnProgressChanged(T(language, "正在校验微软数字签名...", "Verifying Microsoft signature..."), OfficeUninstallStatus.Processing);
                    await MicrosoftToolVerifier.VerifyAsync(saraCmdPath, _scriptRunner, cancellationToken);
                    arguments = BuildArguments(request.Target);
                }

                OnProgressChanged(T(language, $"卸载引擎版本：{FileVersionInfo.GetVersionInfo(saraCmdPath).FileVersion}", $"Uninstall engine version: {FileVersionInfo.GetVersionInfo(saraCmdPath).FileVersion}"), OfficeUninstallStatus.Processing);
                cancellationToken.ThrowIfCancellationRequested();
                string processCleanupSummary = await CloseOfficeProcessesAsync(language, request.ForceCloseApplications, cancellationToken);
                cancellationToken.ThrowIfCancellationRequested();
                OnProgressChanged(T(language, "正在执行 Office 卸载，请按需处理安装程序窗口...", "Uninstalling Office. Respond to any installer windows as needed..."), OfficeUninstallStatus.Processing);
                var executionResult = await _scriptRunner.RunAsync(new ScriptExecutionRequest
                {
                    FileName = saraCmdPath,
                    Arguments = arguments,
                    CancellationToken = cancellationToken,
                    Timeout = TimeSpan.FromMinutes(30),
                    WorkingDirectory = Path.GetDirectoryName(saraCmdPath),
                    OnOutputDataReceived = data => OnProgressChanged($"{T(language, "引擎输出", "Engine Output")}: {data}", OfficeUninstallStatus.Processing),
                    OnErrorDataReceived = data => OnProgressChanged($"{T(language, "引擎错误", "Engine Error")}: {data}", OfficeUninstallStatus.Processing)
                });

                bool backgroundUninstallReported = IndicatesBackgroundUninstall(executionResult);
                bool engineSucceeded = executionResult.IsSuccess
                    || (request.Mode == OfficeUninstallMode.Standard && executionResult.ExitCode == 3010 && !executionResult.IsStillRunning);
                bool monitoringStopped = executionResult.TimedOut || executionResult.WasCanceled;
                IReadOnlyList<InstalledOfficeInfo> postCheckOffice;
                if (backgroundUninstallReported && engineSucceeded)
                {
                    Task<IReadOnlyList<InstalledOfficeInfo>> completion = WaitForBackgroundUninstallAsync(request.Target, request.Mode, language, cancellationToken);
                    try { postCheckOffice = await completion.WaitAsync(cancellationToken); }
                    catch (OperationCanceledException)
                    {
                        MaintenanceCoordinator.Shared.TrackBackgroundTask(completion);
                        monitoringStopped = true;
                        postCheckOffice = DetectInstalledOffice();
                    }
                }
                else postCheckOffice = DetectInstalledOffice();
                string postCheckOfficeSummary = postCheckOffice.Count == 0
                    ? T(language, "未检测到已安装的 Office。", "No installed Office detected.")
                    : string.Join(language == UiLanguage.Chinese ? "；" : "; ", postCheckOffice.Select(item => item.ToString()));
                bool stillHasOfficeAfterUninstall = HasRemainingTargetOrUnknownVersion(GetPostCheckScope(postCheckOffice, request.Mode), request.Target);
                bool requiresAttention = monitoringStopped || !engineSucceeded || stillHasOfficeAfterUninstall;
                int interpretedExitCode = engineSucceeded ? 0 : executionResult.ExitCode;
                string summary = monitoringStopped
                    ? T(language, $"已停止等待。已启动的卸载可能仍在后台运行（进程 {executionResult.ProcessId}），请勿重复启动。", $"Monitoring stopped. Uninstall may still be running (process {executionResult.ProcessId}). Do not start another uninstall.")
                    : BuildSummary(interpretedExitCode, backgroundUninstallReported, stillHasOfficeAfterUninstall, language);
                string nextStepAdvice = BuildNextStepAdvice(
                    interpretedExitCode,
                    request.Target,
                    detectedOfficeSummary,
                    postCheckOfficeSummary,
                    backgroundUninstallReported,
                    stillHasOfficeAfterUninstall,
                    language);
                string saraArchivePath = executionResult.IsStillRunning || cancellationToken.IsCancellationRequested
                    ? extractRoot : ArchiveSaraLogs(extractRoot, logRoot);

                var result = new OfficeUninstallResult
                {
                    ExitCode = executionResult.ExitCode,
                    Summary = summary,
                    Details = BuildDetails(
                        request.Target,
                        executionResult.ExitCode,
                        summary,
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
                    BackgroundUninstallReported = backgroundUninstallReported || executionResult.IsStillRunning,
                    RequiresAttention = requiresAttention
                };

                result.Details = $"{request.Mode}\r\n{summary}\r\n" + result.Details;

                lock (sessionLog) sessionLog.WriteLine(BuildExecutionLog(request.Target, arguments, result, saraArchivePath));

                OnProgressChanged(result.Summary, result.IsSuccess ? OfficeUninstallStatus.Success : OfficeUninstallStatus.Error);
                return result;
            }
            catch (Exception ex)
            {
                lock (sessionLog) sessionLog.WriteLine(ex);
                throw;
            }
            finally { ProgressChanged -= RecordProgress; }

        }

        private static ToolSessionPaths CreateSessionPaths()
        {
            return ToolPathService.CreateSession("OfficeUninstall");
        }

        private async Task<string> CloseOfficeProcessesAsync(UiLanguage language, bool forceClose, CancellationToken cancellationToken)
        {
            return await Task.Run(() =>
            {
                var closedProcesses = new List<string>();
                var skippedProcesses = new List<string>();

                OnProgressChanged(T(language, "正在检查并结束 Office 相关进程...", "Checking and closing Office related processes..."), OfficeUninstallStatus.Processing);

                foreach (string processName in OfficeProcessNames)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    foreach (var process in Process.GetProcessesByName(processName))
                    {
                        try
                        {
                            string displayName = $"{process.ProcessName}({process.Id})";

                            if (!process.HasExited)
                            {
                                OnProgressChanged($"{T(language, "正在结束进程", "Closing process")}: {displayName}", OfficeUninstallStatus.Processing);
                                process.CloseMainWindow();
                                for (int attempt = 0; attempt < 10 && !process.HasExited; attempt++)
                                {
                                    cancellationToken.ThrowIfCancellationRequested();
                                    process.WaitForExit(500);
                                }
                                if (!process.HasExited && forceClose)
                                {
                                    process.Kill(entireProcessTree: true);
                                    process.WaitForExit(5000);
                                }
                                if (!process.HasExited)
                                {
                                    skippedProcesses.Add(displayName);
                                    continue;
                                }
                            }

                            closedProcesses.Add(displayName);
                        }
                        catch (OperationCanceledException) { throw; }
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
                throw new InvalidOperationException(closedSummary + skippedSummary + T(language, "。请保存文档并手动关闭这些应用后重试。", ". Save your documents and close these apps before retrying."));
            });
        }

        private async Task DownloadSaraAsync(string zipPath, UiLanguage language, CancellationToken cancellationToken)
        {
            OnProgressChanged(T(language, "正在下载微软官方 SaRA 企业版...", "Downloading official Microsoft SaRA enterprise package..."), OfficeUninstallStatus.Processing);

            await ToolDownloadService.DownloadAsync(SaraDownloadUrl, zipPath, cancellationToken);
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
            string archivePath = Path.Combine(logRoot, "Engine_Files.zip");
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
            string summary,
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
                $"{T(language, "说明", "Summary")}: {summary}\r\n" +
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

        private async Task<IReadOnlyList<InstalledOfficeInfo>> WaitForBackgroundUninstallAsync(
            OfficeUninstallTargetOption target, OfficeUninstallMode mode, UiLanguage language, CancellationToken cancellationToken)
        {
            OnProgressChanged(
                T(
                    language,
                    "GetHelp 已将 Office 卸载转入后台执行，正在等待卸载完成并复查控制面板残留项...",
                    "GetHelp started Office uninstall in the background. Waiting for completion and checking remaining entries..."),
                OfficeUninstallStatus.Processing);

            DateTime deadline = DateTime.Now.Add(BackgroundUninstallTimeout);
            IReadOnlyList<InstalledOfficeInfo> current = DetectInstalledOffice();

            while (GetPostCheckScope(current, mode).Any(item => IsUninstallTarget(item, target)) && DateTime.Now < deadline)
            {
                int remainingTargets = GetPostCheckScope(current, mode).Count(item => IsUninstallTarget(item, target));
                if (!cancellationToken.IsCancellationRequested) OnProgressChanged(
                    T(
                        language,
                        $"后台卸载仍在进行或等待系统刷新，当前仍检测到 {remainingTargets} 项卸载目标。请不要关闭可能出现的卸载窗口。",
                        $"Background uninstall is still running or waiting for system refresh. {remainingTargets} target entries are still detected. Do not close any uninstall window that appears."),
                    OfficeUninstallStatus.Processing);

                await Task.Delay(BackgroundUninstallPollInterval);
                current = DetectInstalledOffice();
            }

            return current;
        }

        internal static IReadOnlyList<InstalledOfficeInfo> GetPostCheckScope(
            IReadOnlyList<InstalledOfficeInfo> products, OfficeUninstallMode mode)
        {
            return mode == OfficeUninstallMode.Standard
                ? products.Where(item => item.Source == "ClickToRun").ToList()
                : products;
        }

        internal static XDocument BuildStandardRemovalXml(
            OfficeUninstallTargetOption target, IReadOnlyList<InstalledOfficeInfo> products)
        {
            var targets = GetPostCheckScope(products, OfficeUninstallMode.Standard)
                .Where(item => IsUninstallTarget(item, target) && !string.IsNullOrWhiteSpace(item.ProductId)).ToList();
            if (targets.Count == 0)
                throw new InvalidOperationException("未检测到匹配的 C2R 产品；MSI 或损坏的安装请使用深度清理。 / No matching C2R product found. Use deep cleanup for MSI or damaged installations.");
            if (target.UsesDetectedInstalledVersion && (targets.Any(item => string.IsNullOrWhiteSpace(item.OfficeVersion))
                || targets.Select(item => item.OfficeVersion).Distinct(StringComparer.OrdinalIgnoreCase).Count() > 1))
                throw new InvalidOperationException("检测到多个或无法识别的 C2R 版本，请明确选择版本或所有版本。 / Multiple or unidentified C2R versions found. Select a specific version or All.");

            var remove = new XElement("Remove");
            foreach (string id in targets.Select(item => item.ProductId).Distinct(StringComparer.OrdinalIgnoreCase))
                remove.Add(new XElement("Product", new XAttribute("ID", id)));

            return new XDocument(new XElement("Configuration", remove,
                new XElement("Display", new XAttribute("Level", "Full"), new XAttribute("AcceptEULA", "TRUE"))));
        }

        internal static bool IsUninstallTarget(InstalledOfficeInfo product, OfficeUninstallTargetOption target)
        {
            return target.UsesDetectedInstalledVersion
                || string.Equals(target.SaraOfficeVersion, "All", StringComparison.OrdinalIgnoreCase)
                || string.Equals(product.OfficeVersion, target.SaraOfficeVersion, StringComparison.OrdinalIgnoreCase);
        }

        internal static bool HasRemainingTargetOrUnknownVersion(
            IReadOnlyList<InstalledOfficeInfo> products, OfficeUninstallTargetOption target)
        {
            return products.Any(item => IsUninstallTarget(item, target) || string.IsNullOrWhiteSpace(item.OfficeVersion));
        }

        private static bool IndicatesBackgroundUninstall(ScriptExecutionResult result)
        {
            string combinedOutput = $"{result.StandardOutput}\n{result.StandardError}";
            return combinedOutput.Contains("Uninstall office is running in the background", StringComparison.OrdinalIgnoreCase)
                || combinedOutput.Contains("Office uninstall is running in the background", StringComparison.OrdinalIgnoreCase)
                || combinedOutput.Contains("running in the background", StringComparison.OrdinalIgnoreCase);
        }

        internal static string BuildSummary(
            int exitCode,
            bool backgroundUninstallReported,
            bool stillHasOfficeAfterUninstall,
            UiLanguage language)
        {
            if (exitCode == 0 && stillHasOfficeAfterUninstall)
            {
                return T(
                    language,
                    "复查仍检测到卸载目标或版本不明的 Office 条目，不能判定为已完成。建议重启后重新检查。",
                    "Target Office products or entries with an unknown version remain. Completion cannot be confirmed. Restart and check again.");
            }

            if (exitCode == 0 && backgroundUninstallReported)
            {
                return T(
                    language,
                    "后台卸载已完成，复查未再检测到卸载目标，建议重启计算机完成剩余清理。",
                    "Background uninstall completed. No target Office entries remain. Restart is recommended.");
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
            if (exitCode == 0 && stillHasOfficeAfterUninstall)
            {
                return T(
                    language,
                    $"请先确认卸载窗口已结束，重启计算机后复查卸载目标及版本不明的条目。当前复查结果：{postCheckOfficeSummary}",
                    $"Ensure the uninstall window has finished, then restart and recheck the target and unidentified entries. Current post-check: {postCheckOfficeSummary}");
            }

            return SaraResultInterpreter.GetNextStepAdvice(exitCode, target, detectedOfficeSummary, language);
        }

        private static string T(UiLanguage language, string zh, string en)
        {
            return LocalizationService.Localize(language, zh, en);
        }
    }
}
