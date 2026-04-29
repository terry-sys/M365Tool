using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using Office365CleanupTool.Models;

namespace Office365CleanupTool.Services
{
    public class InstallationService
    {
        private const string OdtDetailsUrl = "https://www.microsoft.com/en-us/download/details.aspx?id=49117";
        private const string OdtInstallerFileName = "ODTSetup.exe";
        private const string OdtSetupExecutableName = "setup.exe";
        private const string SessionLogFileName = "install-session.log";

        private readonly IScriptRunner _scriptRunner;
        private readonly ValidationService _validationService;

        public InstallationService()
        {
            _scriptRunner = new ScriptRunner();
            _validationService = new ValidationService();
        }

        public event EventHandler<InstallationProgressEventArgs> ProgressChanged = delegate { };

        public async Task<InstallationExecutionResult> InstallOfficeAsync(
            InstallationConfig config,
            IReadOnlyList<string> excludeApps,
            UiLanguage language = UiLanguage.Chinese)
        {
            if (config == null)
            {
                throw new ArgumentNullException(nameof(config));
            }

            string sessionDirectory = PrepareSessionDirectory();
            string logFilePath = Path.Combine(sessionDirectory, SessionLogFileName);
            using var logWriter = CreateLogWriter(logFilePath);

            try
            {
                Log(logWriter, T(language, $"安装开始，时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}", $"Installation started at: {DateTime.Now:yyyy-MM-dd HH:mm:ss}"));
                Log(logWriter, T(language, $"执行模式: {(config.DownloadOnly ? "仅下载" : "下载并安装")}", $"Execution mode: {(config.DownloadOnly ? "Download only" : "Download and install")}"));
                Log(logWriter, T(language, $"下载目录: {config.DownloadPath}", $"Download path: {config.DownloadPath}"));

                if (!config.IsValid())
                {
                    return FailureWithLog(logWriter, logFilePath, T(language, "安装配置无效，请检查下载路径和语言设置。", "Invalid install configuration. Check download path and language settings."), language);
                }

                ReportProgress(logWriter, T(language, "正在验证系统要求...", "Validating system requirements..."), 5);
                ValidationReport validationReport = await _validationService.GetValidationReportAsync(language);
                InstallationExecutionResult? validationFailure = ValidateBeforeInstall(logWriter, logFilePath, validationReport, language);
                if (validationFailure != null)
                {
                    return validationFailure;
                }

                ReportProgress(logWriter, T(language, "正在准备安装环境...", "Preparing installation environment..."), 15);
                Directory.CreateDirectory(config.DownloadPath);

                string configurationXmlPath = Path.Combine(sessionDirectory, "OfficeInstall.xml");
                WriteConfigurationXml(config, excludeApps, configurationXmlPath);
                Log(logWriter, T(language, $"已生成安装配置: {configurationXmlPath}", $"Generated installation configuration: {configurationXmlPath}"));

                string odtCacheDirectory = ToolPathService.GetOdtCacheDirectory();

                string odtInstallerPath = Path.Combine(odtCacheDirectory, OdtInstallerFileName);
                string setupPath = Path.Combine(config.DownloadPath, OdtSetupExecutableName);

                InstallationExecutionResult? odtPreparationFailure = await EnsureOdtReadyAsync(
                    logWriter,
                    config.DownloadPath,
                    odtInstallerPath,
                    setupPath,
                    logFilePath,
                    language);

                if (odtPreparationFailure != null)
                {
                    return odtPreparationFailure;
                }

                string odtArguments = config.DownloadOnly
                    ? $"/download \"{configurationXmlPath}\""
                    : $"/configure \"{configurationXmlPath}\"";
                string progressMessage = config.DownloadOnly
                    ? T(language, "正在下载 Office 安装源...", "Downloading Office installation source...")
                    : T(language, "正在启动 Office 安装...", "Starting Office installation...");
                string phaseName = config.DownloadOnly ? T(language, "Office 下载", "Office Download") : T(language, "Office 安装", "Office Install");

                ReportProgress(logWriter, progressMessage, 70);
                var executionResult = await _scriptRunner.RunAsync(new ScriptExecutionRequest
                {
                    FileName = setupPath,
                    Arguments = odtArguments,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    WorkingDirectory = config.DownloadPath,
                    OnOutputDataReceived = data =>
                    {
                        if (!string.IsNullOrWhiteSpace(data))
                        {
                            Log(logWriter, $"{phaseName}{T(language, "输出", " output")}: {data}");
                            OnProgressChanged($"{phaseName}{T(language, "输出", " output")}: {data}", -1);
                        }
                    },
                    OnErrorDataReceived = data =>
                    {
                        if (!string.IsNullOrWhiteSpace(data))
                        {
                            Log(logWriter, $"{phaseName}{T(language, "错误", " error")}: {data}");
                            OnProgressChanged($"{phaseName}{T(language, "错误", " error")}: {data}", -1);
                        }
                    }
                });

                LogProcessResult(logWriter, phaseName, executionResult.StandardOutput, executionResult.StandardError, language);

                if (!executionResult.IsSuccess)
                {
                    string diagnostic = BuildDiagnosticMessage(
                        T(language, $"{phaseName}执行失败。", $"{phaseName} failed."),
                        config.DownloadPath,
                        odtInstallerPath,
                        executionResult.StandardOutput,
                        executionResult.StandardError,
                        language);

                    return FailureWithLog(logWriter, logFilePath, diagnostic, language, executionResult.StandardOutput, executionResult.StandardError);
                }

                string completedMessage = config.DownloadOnly
                    ? T(language, "Office 安装源下载完成。", "Office installation source download completed.")
                    : T(language, "Office 安装程序已启动完成。", "Office installer started successfully.");
                ReportProgress(logWriter, completedMessage, 100);
                Log(logWriter, T(language, "安装流程结束，结果: 成功", "Installation workflow finished: Success"));

                return InstallationExecutionResult.FromSuccess(
                    executionResult.StandardOutput,
                    executionResult.StandardError,
                    logFilePath,
                    config.DownloadOnly,
                    language);
            }
            catch (Exception ex)
            {
                Log(logWriter, T(language, "发生异常", "Exception") + ": " + ex);
                return FailureWithLog(logWriter, logFilePath, T(language, $"执行过程中发生异常: {ex.Message}", $"Exception during execution: {ex.Message}"), language);
            }
        }

        private async Task<InstallationExecutionResult?> EnsureOdtReadyAsync(
            StreamWriter logWriter,
            string workingDirectory,
            string odtInstallerPath,
            string setupPath,
            string logFilePath,
            UiLanguage language)
        {
            if (File.Exists(setupPath))
            {
                ReportProgress(logWriter, T(language, "检测到已解压的 ODT，复用现有 setup.exe。", "Detected extracted ODT; reusing existing setup.exe."), 25);
                return null;
            }

            ReportProgress(logWriter, T(language, "正在获取微软官方 ODT 下载地址...", "Fetching official Microsoft ODT download URL..."), 25);
            string odtDownloadUrl = await GetOdtDownloadUrlAsync();
            Log(logWriter, T(language, $"ODT 下载地址: {odtDownloadUrl}", $"ODT download URL: {odtDownloadUrl}"));

            if (File.Exists(odtInstallerPath) && new FileInfo(odtInstallerPath).Length > 0)
            {
                ReportProgress(logWriter, T(language, "检测到本地 ODT 安装包缓存，跳过重复下载。", "Detected local ODT package cache; skipping re-download."), 35);
            }
            else
            {
                ReportProgress(logWriter, T(language, "正在下载 Office Deployment Tool...", "Downloading Office Deployment Tool..."), 35);
                await DownloadFileAsync(odtDownloadUrl, odtInstallerPath);
                Log(logWriter, T(language, $"ODT 安装包已下载到: {odtInstallerPath}", $"ODT package downloaded to: {odtInstallerPath}"));
            }

            ReportProgress(logWriter, T(language, "正在解压 Office Deployment Tool...", "Extracting Office Deployment Tool..."), 50);
            var extractResult = await _scriptRunner.RunAsync(new ScriptExecutionRequest
            {
                FileName = odtInstallerPath,
                Arguments = $"/quiet /extract:\"{workingDirectory}\"",
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                WorkingDirectory = workingDirectory
            });

            LogProcessResult(logWriter, T(language, "ODT 解压", "ODT Extract"), extractResult.StandardOutput, extractResult.StandardError, language);

            if (!extractResult.IsSuccess)
            {
                string diagnostic = BuildDiagnosticMessage(
                    T(language, "ODT 解压失败。", "ODT extract failed."),
                    workingDirectory,
                    odtInstallerPath,
                    extractResult.StandardOutput,
                    extractResult.StandardError,
                    language);

                return FailureWithLog(logWriter, logFilePath, diagnostic, language, extractResult.StandardOutput, extractResult.StandardError);
            }

            if (!File.Exists(setupPath))
            {
                string diagnostic = BuildDiagnosticMessage(T(language, "未找到解压后的 setup.exe。", "setup.exe was not found after extraction."), workingDirectory, odtInstallerPath, language: language);
                return FailureWithLog(logWriter, logFilePath, diagnostic, language);
            }

            return null;
        }

        private static void WriteConfigurationXml(InstallationConfig config, IReadOnlyList<string> excludeApps, string configurationXmlPath)
        {
            XDocument xml = OfficeDeploymentXmlBuilder.Build(config, excludeApps);
            xml.Save(configurationXmlPath);
        }

        private static async Task DownloadFileAsync(string url, string filePath)
        {
            using var client = new HttpClient();
            client.DefaultRequestHeaders.UserAgent.ParseAdd("M365Tool/1.0");

            using var response = await client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead);
            response.EnsureSuccessStatusCode();

            await using var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None);
            await response.Content.CopyToAsync(fileStream);
        }

        private static async Task<string> GetOdtDownloadUrlAsync()
        {
            using var client = new HttpClient();
            client.DefaultRequestHeaders.UserAgent.ParseAdd("M365Tool/1.0");
            string pageContent = await client.GetStringAsync(OdtDetailsUrl);

            string[] patterns =
            {
                "\"url\":\"(https:\\\\/\\\\/download\\\\.microsoft\\\\.com\\\\/[^\\\"]*officedeploymenttool[^\\\"]*)\"",
                "href=\"(https://download\\.microsoft\\.com/[^\"]*officedeploymenttool[^\"]*)\"",
                "href='(https://download\\.microsoft\\.com/[^']*officedeploymenttool[^']*)'",
                "(https://download\\.microsoft\\.com/\\S*officedeploymenttool\\S*\\.exe)"
            };

            foreach (string pattern in patterns)
            {
                var match = Regex.Match(pageContent, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    return match.Groups[1].Value.Replace("\\/", "/");
                }
            }

            throw new InvalidOperationException("未能从微软官方下载页解析出 ODT 下载地址。");
        }

        private static string PrepareSessionDirectory()
        {
            return ToolPathService.CreateSession("Install").LogsDirectory;
        }

        private static StreamWriter CreateLogWriter(string logFilePath)
        {
            var stream = new FileStream(logFilePath, FileMode.Create, FileAccess.Write, FileShare.Read);
            return new StreamWriter(stream, new UTF8Encoding(false)) { AutoFlush = true };
        }

        private static void Log(StreamWriter writer, string message)
        {
            writer.WriteLine($"[{DateTime.Now:HH:mm:ss}] {message}");
        }

        private void ReportProgress(StreamWriter writer, string message, int progress)
        {
            Log(writer, message);
            OnProgressChanged(message, progress);
        }

        private static void LogProcessResult(
            StreamWriter writer,
            string phase,
            string standardOutput,
            string standardError,
            UiLanguage language)
        {
            if (!string.IsNullOrWhiteSpace(standardOutput))
            {
                Log(writer, $"{phase} {T(language, "标准输出", "Standard output")}: {TrimForMessage(standardOutput)}");
            }

            if (!string.IsNullOrWhiteSpace(standardError))
            {
                Log(writer, $"{phase} {T(language, "标准错误", "Standard error")}: {TrimForMessage(standardError)}");
            }
        }

        private static string BuildDiagnosticMessage(
            string summary,
            string workingDirectory,
            string odtInstallerPath,
            string standardOutput = "",
            string standardError = "",
            UiLanguage language = UiLanguage.Chinese)
        {
            var diagnostics = new List<string>
            {
                summary,
                $"{T(language, "工作目录", "Working directory")}: {workingDirectory}",
                $"{T(language, "ODT 安装包", "ODT package")}: {odtInstallerPath}",
                $"{T(language, "ODT 安装包存在", "ODT package exists")}: {File.Exists(odtInstallerPath)}"
            };

            if (Directory.Exists(workingDirectory))
            {
                string[] files = Directory.GetFiles(workingDirectory);
                diagnostics.Add($"{T(language, "工作目录文件数", "Working directory file count")}: {files.Length}");

                if (files.Length > 0)
                {
                    diagnostics.Add(T(language, "目录文件", "Files") + ": " + string.Join(", ", GetTopFileNames(files, 8)));
                }
            }
            else
            {
                diagnostics.Add(T(language, "工作目录不存在。", "Working directory does not exist."));
            }

            if (!string.IsNullOrWhiteSpace(standardOutput))
            {
                diagnostics.Add(T(language, "标准输出", "Standard output") + ": " + TrimForMessage(standardOutput));
            }

            if (!string.IsNullOrWhiteSpace(standardError))
            {
                diagnostics.Add(T(language, "标准错误", "Standard error") + ": " + TrimForMessage(standardError));
            }

            return string.Join(Environment.NewLine, diagnostics);
        }

        private InstallationExecutionResult? ValidateBeforeInstall(
            StreamWriter logWriter,
            string logFilePath,
            ValidationReport validationReport,
            UiLanguage language)
        {
            var blockingFailures = new List<ValidationCheckResult>();
            var warningFailures = new List<ValidationCheckResult>();

            foreach (ValidationCheckResult item in validationReport.Items)
            {
                if (!item.RequiresAttention)
                {
                    continue;
                }

                if (item.IsBlocking)
                {
                    blockingFailures.Add(item);
                }
                else
                {
                    warningFailures.Add(item);
                }
            }

            if (warningFailures.Count > 0)
            {
                foreach (ValidationCheckResult warning in warningFailures)
                {
                    Log(logWriter, $"系统检查风险项: [{warning.Name}] {warning.Detail}");
                }

                string warningDetails = string.Join(language == UiLanguage.Chinese ? "；" : "; ", warningFailures.Select(warning => $"[{warning.Name}] {warning.Detail}"));
                string warningSummary = T(language, $"检测到可继续执行的风险项（不会阻止安装）：{warningDetails}", $"Non-blocking risk items were detected: {warningDetails}");
                ReportProgress(logWriter, warningSummary, 10);
            }

            if (blockingFailures.Count == 0)
            {
                return null;
            }

            var failureLines = new List<string>
            {
                T(language, "系统要求验证失败，无法继续执行。", "System requirement validation failed and execution cannot continue.")
            };

            foreach (ValidationCheckResult failure in blockingFailures)
            {
                failureLines.Add($"[{failure.Name}] {failure.Detail}");
            }

            return FailureWithLog(logWriter, logFilePath, string.Join(Environment.NewLine, failureLines), language);
        }

        private static IEnumerable<string> GetTopFileNames(IEnumerable<string> filePaths, int maxCount)
        {
            int count = 0;
            foreach (string filePath in filePaths)
            {
                yield return Path.GetFileName(filePath);
                count++;
                if (count >= maxCount)
                {
                    yield break;
                }
            }
        }

        private static string TrimForMessage(string value)
        {
            string normalized = value.Replace("\r", " ").Replace("\n", " ").Trim();
            return normalized.Length <= 300 ? normalized : normalized.Substring(0, 300) + "...";
        }

        private static InstallationExecutionResult FailureWithLog(
            StreamWriter writer,
            string logFilePath,
            string message,
            UiLanguage language,
            string standardOutput = "",
            string standardError = "")
        {
            Log(writer, T(language, "安装流程结束，结果: 失败", "Installation workflow finished: Failed"));
            Log(writer, T(language, "失败原因", "Failure reason") + ": " + TrimForMessage(message));

            return InstallationExecutionResult.Failure(message, standardOutput, standardError, logFilePath);
        }

        private void OnProgressChanged(string message, int progress)
        {
            ProgressChanged(this, new InstallationProgressEventArgs(message, progress));
        }

        private static string T(UiLanguage language, string zh, string en)
        {
            return LocalizationService.Localize(language, zh, en);
        }
    }

    public class InstallationProgressEventArgs : EventArgs
    {
        public string Message { get; }

        public int Progress { get; }

        public InstallationProgressEventArgs(string message, int progress)
        {
            Message = message;
            Progress = progress;
        }
    }

    public class InstallationExecutionResult
    {
        public bool Success { get; private set; }

        public string Message { get; private set; } = string.Empty;

        public string StandardOutput { get; private set; } = string.Empty;

        public string StandardError { get; private set; } = string.Empty;

        public string LogFilePath { get; private set; } = string.Empty;

        public bool DownloadOnly { get; private set; }

        public static InstallationExecutionResult FromSuccess(
            string standardOutput,
            string standardError,
            string logFilePath,
            bool downloadOnly,
            UiLanguage language = UiLanguage.Chinese)
        {
            return new InstallationExecutionResult
            {
                Success = true,
                Message = downloadOnly
                    ? LocalizationService.Localize(language, "Office 安装源已下载完成。", "Office installation source download completed.")
                    : LocalizationService.Localize(language, "Office 安装程序已启动。", "Office installer started."),
                StandardOutput = standardOutput,
                StandardError = standardError,
                LogFilePath = logFilePath,
                DownloadOnly = downloadOnly
            };
        }

        public static InstallationExecutionResult Failure(
            string message,
            string standardOutput = "",
            string standardError = "",
            string logFilePath = "")
        {
            return new InstallationExecutionResult
            {
                Success = false,
                Message = message,
                StandardOutput = standardOutput,
                StandardError = standardError,
                LogFilePath = logFilePath
            };
        }
    }
}
