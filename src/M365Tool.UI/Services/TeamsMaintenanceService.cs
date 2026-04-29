using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace Office365CleanupTool.Services
{
    public class TeamsMaintenanceService
    {
        private static readonly string[] SaraExecutableCandidates =
        {
            "SaRAcmd.exe",
            "GetHelpCmd.exe"
        };

        private const string SaraDownloadUrl = "https://aka.ms/SaRA_EnterpriseVersionFiles";
        private static readonly string[] TeamsProcessNames =
        {
            "Teams",
            "ms-teams",
            "msteams",
            "TeamsBootstrapper"
        };

        public Task<TeamsMaintenanceResult> ResetAndClearCacheAsync(UiLanguage language = UiLanguage.Chinese)
        {
            return Task.Run(() =>
            {
                var result = new TeamsMaintenanceResult();

                StopTeamsProcesses(result.StoppedProcesses, result.Warnings, language);
                ResetTeams(result, language);
                ClearTeamsCache(result, language);

                result.Success = string.IsNullOrWhiteSpace(result.ResetError) && result.CacheTargets.Count > 0;
                result.Message = BuildMaintenanceSummary(result, language);
                return result;
            });
        }

        public Task<TeamsSignInRecordCleanupResult> DeleteSignInRecordsAsync(UiLanguage language = UiLanguage.Chinese)
        {
            return Task.Run(() =>
            {
                var result = new TeamsSignInRecordCleanupResult();

                StopTeamsProcesses(result.StoppedProcesses, result.Warnings, language);

                IReadOnlyList<TeamsPathTarget> allTargets = GetSignInRecordTargets(language);
                foreach (TeamsPathTarget target in allTargets)
                {
                    if (!Directory.Exists(target.Path))
                    {
                        result.Warnings.Add(T(language, $"未找到目录：{target.Path}", $"Directory not found: {target.Path}"));
                        continue;
                    }

                    result.ProcessedTargets.Add(target);
                    ClearDirectoryContents(target.Path, result);
                }

                result.Success = result.ProcessedTargets.Count == allTargets.Count;
                result.Message = BuildSignInRecordSummary(result, language);
                return result;
            });
        }

        public async Task<TeamsAddinRepairResult> RepairOutlookAddinAsync(UiLanguage language = UiLanguage.Chinese)
        {
            var result = new TeamsAddinRepairResult();

            ToolSessionPaths session = ToolPathService.CreateSession("TeamsAddinRepair");
            string saraZipPath = Path.Combine(session.DownloadsDirectory, "SaRA.zip");
            string extractRoot = Path.Combine(session.ExtractedDirectory, "SaRA");
            string logRoot = session.LogsDirectory;

            using (var client = new HttpClient())
            using (var response = await client.GetAsync(SaraDownloadUrl, HttpCompletionOption.ResponseHeadersRead))
            {
                response.EnsureSuccessStatusCode();
                await using var fileStream = new FileStream(saraZipPath, FileMode.Create, FileAccess.Write, FileShare.None);
                await response.Content.CopyToAsync(fileStream);
            }

            ZipFile.ExtractToDirectory(saraZipPath, extractRoot, true);

            string executablePath = FindSaraExecutable(extractRoot, language);
            var startInfo = new ProcessStartInfo
            {
                FileName = executablePath,
                Arguments = $"-S TeamsAddinScenario -AcceptEula -LogFolder \"{logRoot}\"",
                WorkingDirectory = Path.GetDirectoryName(executablePath),
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using Process process = Process.Start(startInfo)
                ?? throw new InvalidOperationException(T(language, "无法启动 Teams Meeting Add-in 修复命令。", "Failed to start Teams Meeting Add-in repair command."));

            string output = process.StandardOutput.ReadToEnd().Trim();
            string error = process.StandardError.ReadToEnd().Trim();
            process.WaitForExit();

            result.ExitCode = process.ExitCode;
            result.StandardOutput = output;
            result.StandardError = error;
            result.LogDirectory = logRoot;
            result.HasArgumentError = ContainsInvalidArguments(output, error);
            result.HasDiagnosticFailure = ContainsDiagnosticFailure(output);
            result.TeamsNotInstalledByDiagnostic = ContainsTeamsNotInstalled(output);
            result.Success = (process.ExitCode == 0 || process.ExitCode == 23 || process.ExitCode == 24)
                && !result.HasArgumentError
                && !result.HasDiagnosticFailure;

            string executionLogPath = Path.Combine(logRoot, "execution.log");
            File.WriteAllText(executionLogPath, BuildAddinRepairExecutionLog(result), Encoding.UTF8);

            result.Message = BuildAddinRepairSummary(result, language);
            return result;
        }

        public TeamsPathTarget GetPrimaryCacheTarget(UiLanguage language = UiLanguage.Chinese)
        {
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return new TeamsPathTarget(
                T(language, "新版 Teams 缓存", "New Teams Cache"),
                Path.Combine(localAppData, "Packages", "MSTeams_8wekyb3d8bbwe", "LocalCache", "Microsoft", "MSTeams"));
        }

        public IReadOnlyList<TeamsPathTarget> GetSignInRecordTargets(UiLanguage language = UiLanguage.Chinese)
        {
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

            return new List<TeamsPathTarget>
            {
                new TeamsPathTarget("IdentityCache", Path.Combine(localAppData, "Microsoft", "IdentityCache")),
                new TeamsPathTarget("TokenBroker", Path.Combine(localAppData, "Microsoft", "TokenBroker")),
                new TeamsPathTarget(T(language, "新版 Teams 包目录", "New Teams Package"), Path.Combine(localAppData, "Packages", "MSTeams_8wekyb3d8bbwe")),
                new TeamsPathTarget("AAD Broker Plugin", Path.Combine(localAppData, "Packages", "Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy")),
                new TeamsPathTarget("OneAuth", Path.Combine(localAppData, "Microsoft", "OneAuth"))
            };
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

        private static void StopTeamsProcesses(List<string> stoppedProcesses, List<string> warnings, UiLanguage language)
        {
            foreach (string processName in TeamsProcessNames)
            {
                foreach (Process process in Process.GetProcessesByName(processName))
                {
                    try
                    {
                        process.Kill(true);
                        stoppedProcesses.Add(process.ProcessName);
                    }
                    catch (Exception ex)
                    {
                        warnings.Add(T(language, $"无法结束进程 {process.ProcessName}: {ex.Message}", $"Failed to close process {process.ProcessName}: {ex.Message}"));
                    }
                    finally
                    {
                        process.Dispose();
                    }
                }
            }
        }

        private static void ResetTeams(TeamsMaintenanceResult result, UiLanguage language)
        {
            string script = @"
$pkg = Get-AppxPackage | Where-Object { $_.Name -like '*MSTeams*' -or $_.Name -like '*MicrosoftTeams*' } | Select-Object -First 1
if (-not $pkg) {
  Write-Output 'STATUS:PKG_NOT_FOUND'
  exit 0
}

$command = Get-Command Reset-AppxPackage -ErrorAction SilentlyContinue
if (-not $command) {
  Write-Output ('STATUS:RESET_UNSUPPORTED|' + $pkg.Name + '|' + $pkg.PackageFullName)
  exit 0
}

try {
  Reset-AppxPackage -Package $pkg.PackageFullName -ErrorAction Stop
  Write-Output ('STATUS:RESET_OK|' + $pkg.Name + '|' + $pkg.PackageFullName)
  exit 0
}
catch {
  Write-Error $_.Exception.Message
  exit 1
}";

            string encodedScript = Convert.ToBase64String(Encoding.Unicode.GetBytes(script));

            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -ExecutionPolicy Bypass -EncodedCommand {encodedScript}",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using Process process = Process.Start(startInfo)
                ?? throw new InvalidOperationException(T(language, "无法启动 Teams 重置命令。", "Failed to start Teams reset command."));

            string output = process.StandardOutput.ReadToEnd().Trim();
            string error = process.StandardError.ReadToEnd().Trim();
            process.WaitForExit();

            if (process.ExitCode != 0)
            {
                result.ResetError = string.IsNullOrWhiteSpace(error)
                    ? (string.IsNullOrWhiteSpace(output) ? T(language, "Teams 重置失败。", "Teams reset failed.") : output)
                    : error;
                return;
            }

            string statusLine = output
                .Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries)
                .FirstOrDefault(line => line.StartsWith("STATUS:", StringComparison.OrdinalIgnoreCase))
                ?? string.Empty;

            if (statusLine.StartsWith("STATUS:RESET_OK|", StringComparison.OrdinalIgnoreCase))
            {
                string[] parts = statusLine.Split('|');
                result.PackageId = parts.Length > 1 ? parts[1].Trim() : string.Empty;
                result.PackageFullName = parts.Length > 2 ? parts[2].Trim() : string.Empty;
                return;
            }

            if (statusLine.StartsWith("STATUS:RESET_UNSUPPORTED|", StringComparison.OrdinalIgnoreCase))
            {
                string[] parts = statusLine.Split('|');
                result.PackageId = parts.Length > 1 ? parts[1].Trim() : string.Empty;
                result.PackageFullName = parts.Length > 2 ? parts[2].Trim() : string.Empty;
                result.Warnings.Add(T(language, "当前系统不支持 Reset-AppxPackage，已继续执行 Teams 缓存清理。", "Reset-AppxPackage is not supported on this system. Cache cleanup continues."));
                return;
            }

            if (statusLine.StartsWith("STATUS:PKG_NOT_FOUND", StringComparison.OrdinalIgnoreCase))
            {
                result.Warnings.Add(T(language, "未检测到已安装的新版 Teams 应用包，已跳过重置步骤。", "No installed new Teams package was detected. Reset step was skipped."));
                return;
            }

            if (!string.IsNullOrWhiteSpace(output))
            {
                string[] parts = output.Split('|');
                result.PackageId = parts.Length > 0 ? parts[0].Trim() : string.Empty;
                result.PackageFullName = parts.Length > 1 ? parts[1].Trim() : string.Empty;
            }
        }

        private static string FindSaraExecutable(string extractRoot, UiLanguage language)
        {
            foreach (string executableName in SaraExecutableCandidates)
            {
                string[] candidates = Directory.GetFiles(extractRoot, executableName, SearchOption.AllDirectories);
                if (candidates.Length > 0)
                {
                    return candidates[0];
                }
            }

            throw new FileNotFoundException(
                T(language, $"未在解压目录中找到 {string.Join(" 或 ", SaraExecutableCandidates)}。", $"Could not find {string.Join(" or ", SaraExecutableCandidates)} in extracted directory."),
                extractRoot);
        }

        private void ClearTeamsCache(TeamsMaintenanceResult result, UiLanguage language)
        {
            TeamsPathTarget target = GetPrimaryCacheTarget(language);
            if (!Directory.Exists(target.Path))
            {
                result.Warnings.Add(T(language, "未找到新版 Teams 缓存目录。", "New Teams cache directory was not found."));
                return;
            }

            result.CacheTargets.Add(target);
            ClearDirectoryContents(target.Path, result);
        }

        private static void ClearDirectoryContents(string rootPath, TeamsMaintenanceResult result)
        {
            foreach (string directory in Directory.GetDirectories(rootPath))
            {
                Directory.Delete(directory, true);
                result.DeletedDirectories++;
            }

            foreach (string file in Directory.GetFiles(rootPath))
            {
                File.SetAttributes(file, FileAttributes.Normal);
                File.Delete(file);
                result.DeletedFiles++;
            }
        }

        private static void ClearDirectoryContents(string rootPath, TeamsSignInRecordCleanupResult result)
        {
            foreach (string directory in Directory.GetDirectories(rootPath))
            {
                Directory.Delete(directory, true);
                result.DeletedDirectories++;
            }

            foreach (string file in Directory.GetFiles(rootPath))
            {
                File.SetAttributes(file, FileAttributes.Normal);
                File.Delete(file);
                result.DeletedFiles++;
            }
        }

        private static string BuildMaintenanceSummary(TeamsMaintenanceResult result, UiLanguage language)
        {
            string friendlyPackageName = string.IsNullOrWhiteSpace(result.PackageId)
                ? T(language, "未返回应用包标识", "Package identifier not returned")
                : GetFriendlyPackageName(result.PackageId, language);

            string headline = string.IsNullOrWhiteSpace(result.ResetError)
                ? T(language, "Teams 重置与缓存清理已执行完成。", "Teams reset and cache cleanup completed.")
                : T(language, "Teams 缓存已清理，但应用重置未完全成功。", "Teams cache was cleaned, but app reset did not fully succeed.");

            var lines = new List<string>
            {
                T(language, "执行结果", "Result"),
                headline,
                string.Empty,
                T(language, "应用信息", "App Info"),
                $"- {T(language, "应用包", "Package")}: {friendlyPackageName}",
                $"- {T(language, "包全名", "Full package name")}: {(string.IsNullOrWhiteSpace(result.PackageFullName) ? T(language, "未返回包全名", "Full package name not returned") : result.PackageFullName)}",
                string.Empty,
                T(language, "已关闭进程", "Closed Processes"),
                BuildProcessSummary(result.StoppedProcesses, language),
                string.Empty,
                T(language, "已处理位置", "Processed Paths"),
                result.CacheTargets.Count > 0
                    ? string.Join(Environment.NewLine, result.CacheTargets.Select(target => $"- {target.DisplayName}: {target.Path}"))
                    : $"- {T(language, "未找到可清理的缓存目录", "No cache directory found to clean")}",
                string.Empty,
                T(language, "清理统计", "Cleanup Stats"),
                $"- {T(language, "已删除目录", "Deleted directories")}: {result.DeletedDirectories}",
                $"- {T(language, "已删除文件", "Deleted files")}: {result.DeletedFiles}",
                string.Empty,
                T(language, "后续建议", "Next Steps"),
                $"- {T(language, "重新启动 Teams", "Restart Teams")}",
                $"- {T(language, "重新登录账号", "Sign in again")}",
                $"- {T(language, "首次启动可能稍慢，因为本地数据需要重新生成", "First launch may be slower while local data is rebuilt")}"
            };

            if (!string.IsNullOrWhiteSpace(result.ResetError))
            {
                lines.Add(string.Empty);
                lines.Add(T(language, "重置提示", "Reset Notice"));
                lines.Add($"- {result.ResetError}");
            }

            if (result.Warnings.Count > 0)
            {
                lines.Add(string.Empty);
                lines.Add(T(language, "提示", "Notice"));
                lines.AddRange(result.Warnings.Select(warning => $"- {warning}"));
            }

            return string.Join(Environment.NewLine, lines);
        }

        private string BuildSignInRecordSummary(TeamsSignInRecordCleanupResult result, UiLanguage language)
        {
            var lines = new List<string>
            {
                T(language, "执行结果", "Result"),
                result.Success ? T(language, "Teams 登录记录清理已完成。", "Teams sign-in record cleanup completed.") : T(language, "Teams 登录记录清理未完全完成，请查看提示。", "Teams sign-in record cleanup did not fully complete. See notices."),
                string.Empty,
                T(language, "已关闭进程", "Closed Processes"),
                BuildProcessSummary(result.StoppedProcesses, language),
                string.Empty,
                T(language, "已处理位置", "Processed Paths"),
                result.ProcessedTargets.Count > 0
                    ? string.Join(Environment.NewLine, result.ProcessedTargets.Select(target => $"- {target.DisplayName}: {target.Path}"))
                    : $"- {T(language, "未成功处理任何目录", "No directories were processed")}",
                string.Empty,
                T(language, "清理统计", "Cleanup Stats"),
                $"- {T(language, "已删除目录", "Deleted directories")}: {result.DeletedDirectories}",
                $"- {T(language, "已删除文件", "Deleted files")}: {result.DeletedFiles}",
                string.Empty,
                T(language, "重要提示", "Important"),
                $"- {T(language, "完成后必须重启电脑", "Restart your computer after completion")}",
                $"- {T(language, "如果不重启就直接打开 New Teams，可能出现 Error 1001", "Opening New Teams without restart may cause Error 1001")}",
                $"- {T(language, "重启后再打开 New Teams 登录，此时不会再看到之前的账号记录", "After restart, previous sign-in records should not appear")}"
            };

            if (result.Warnings.Count > 0)
            {
                lines.Add(string.Empty);
                lines.Add(T(language, "提示", "Notice"));
                lines.AddRange(result.Warnings.Select(warning => $"- {warning}"));
            }

            return string.Join(Environment.NewLine, lines);
        }

        private static string BuildProcessSummary(List<string> processes, UiLanguage language)
        {
            return processes.Count > 0
                ? string.Join(Environment.NewLine, processes.Distinct(StringComparer.OrdinalIgnoreCase).Select(process => $"- {process}"))
                : $"- {T(language, "未检测到正在运行的 Teams 进程", "No running Teams process was detected")}";
        }

        private static string GetFriendlyPackageName(string packageId, UiLanguage language)
        {
            if (packageId.Contains("MSTeams", StringComparison.OrdinalIgnoreCase))
            {
                return T(language, "新版 Teams", "New Teams");
            }

            if (packageId.Contains("MicrosoftTeams", StringComparison.OrdinalIgnoreCase))
            {
                return "Teams";
            }

            return packageId;
        }

        private static string BuildAddinRepairSummary(TeamsAddinRepairResult result, UiLanguage language)
        {
            if (result.HasArgumentError)
            {
                var badArgsLines = new List<string>
                {
                    T(language, "执行结果", "Result"),
                    T(language, "命令行参数无效，修复场景未执行成功。", "Command-line arguments are invalid. Repair scenario did not run successfully."),
                    string.Empty,
                    T(language, "场景信息", "Scenario Info"),
                    $"- {T(language, "目标", "Target")}: {T(language, "经典 Outlook 的 Teams Meeting Add-in", "Teams Meeting Add-in for classic Outlook")}",
                    $"- {T(language, "返回码", "Return code")}: {result.ExitCode}",
                    $"- {T(language, "日志目录", "Log directory")}: {result.LogDirectory}",
                    string.Empty,
                    T(language, "建议", "Suggestion"),
                    $"- {T(language, "已自动切换到兼容参数模板，请重试一次。", "A compatible argument template has been applied. Please retry once.")}",
                    $"- {T(language, "如仍失败，请把本次“标准输出/标准错误”发给我继续定位。", "If it still fails, share Standard Output/Error for further analysis.")}"
                };

                if (!string.IsNullOrWhiteSpace(result.StandardOutput))
                {
                    badArgsLines.Add(string.Empty);
                    badArgsLines.Add(T(language, "标准输出", "Standard Output"));
                    badArgsLines.Add(result.StandardOutput);
                }

                if (!string.IsNullOrWhiteSpace(result.StandardError))
                {
                    badArgsLines.Add(string.Empty);
                    badArgsLines.Add(T(language, "标准错误", "Standard Error"));
                    badArgsLines.Add(result.StandardError);
                }

                return string.Join(Environment.NewLine, badArgsLines);
            }

            if (result.HasDiagnosticFailure)
            {
                var failedLines = new List<string>
                {
                    T(language, "执行结果", "Result"),
                    result.TeamsNotInstalledByDiagnostic
                        ? T(language, "诊断显示未检测到 Teams，会议插件修复未完成。", "Diagnostic reports Teams is not installed. Meeting add-in repair did not complete.")
                        : T(language, "诊断已返回失败结果，会议插件修复未完成。", "Diagnostic returned a failed result. Meeting add-in repair did not complete."),
                    string.Empty,
                    T(language, "场景信息", "Scenario Info"),
                    $"- {T(language, "目标", "Target")}: {T(language, "经典 Outlook 的 Teams Meeting Add-in", "Teams Meeting Add-in for classic Outlook")}",
                    $"- {T(language, "返回码", "Return code")}: {result.ExitCode}",
                    $"- {T(language, "日志目录", "Log directory")}: {result.LogDirectory}",
                    string.Empty,
                    T(language, "建议", "Suggestion"),
                    result.TeamsNotInstalledByDiagnostic
                        ? $"- {T(language, "请先确认当前用户已安装并至少启动过一次 New Teams，然后再重试。", "Ensure New Teams is installed for current user and launched at least once, then retry.")}"
                        : $"- {T(language, "请打开日志目录查看诊断明细并重试。", "Open log directory, review diagnostic details, and retry.")}"
                };

                if (!string.IsNullOrWhiteSpace(result.StandardOutput))
                {
                    failedLines.Add(string.Empty);
                    failedLines.Add(T(language, "标准输出", "Standard Output"));
                    failedLines.Add(result.StandardOutput);
                }

                if (!string.IsNullOrWhiteSpace(result.StandardError))
                {
                    failedLines.Add(string.Empty);
                    failedLines.Add(T(language, "标准错误", "Standard Error"));
                    failedLines.Add(result.StandardError);
                }

                return string.Join(Environment.NewLine, failedLines);
            }

            string outcome = result.ExitCode switch
            {
                0 => T(language, "场景已成功完成。请退出并重新启动 Outlook。", "Scenario completed successfully. Exit and restart Outlook."),
                23 => T(language, "已修复相关注册表项。请退出并重新启动 Outlook。", "Registry items were repaired. Exit and restart Outlook."),
                24 => T(language, "已重新注册 Microsoft.Teams.AddinLoader.dll。请先重启 Teams，再重启 Outlook。", "Microsoft.Teams.AddinLoader.dll was re-registered. Restart Teams, then Outlook."),
                20 => T(language, "未检测到已安装的 Teams。", "No installed Teams was detected."),
                21 => T(language, "未检测到 Outlook 2013 或更高版本。", "Outlook 2013 or later was not detected."),
                22 => T(language, "系统前置条件不足。", "System prerequisites are not satisfied."),
                _ => T(language, "修复场景未成功完成，请查看输出详情。", "Repair scenario did not complete successfully. See details.")
            };

            var lines = new List<string>
            {
                T(language, "执行结果", "Result"),
                outcome,
                string.Empty,
                T(language, "场景信息", "Scenario Info"),
                $"- {T(language, "目标", "Target")}: {T(language, "经典 Outlook 的 Teams Meeting Add-in", "Teams Meeting Add-in for classic Outlook")}",
                $"- {T(language, "返回码", "Return code")}: {result.ExitCode}",
                $"- {T(language, "日志目录", "Log directory")}: {result.LogDirectory}"
            };

            if (!string.IsNullOrWhiteSpace(result.StandardOutput))
            {
                lines.Add(string.Empty);
                lines.Add(T(language, "标准输出", "Standard Output"));
                lines.Add(result.StandardOutput);
            }

            if (!string.IsNullOrWhiteSpace(result.StandardError))
            {
                lines.Add(string.Empty);
                lines.Add(T(language, "标准错误", "Standard Error"));
                lines.Add(result.StandardError);
            }

            return string.Join(Environment.NewLine, lines);
        }

        private static bool ContainsInvalidArguments(string output, string error)
        {
            string allText = (output ?? string.Empty) + "\n" + (error ?? string.Empty);
            return allText.IndexOf("Invalid command line arguments", StringComparison.OrdinalIgnoreCase) >= 0
                || allText.IndexOf("Usage: GetHelpCmd.exe", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool ContainsDiagnosticFailure(string output)
        {
            if (string.IsNullOrWhiteSpace(output))
            {
                return false;
            }

            return output.IndexOf("ExecutionResult: Failed", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool ContainsTeamsNotInstalled(string output)
        {
            if (string.IsNullOrWhiteSpace(output))
            {
                return false;
            }

            return output.IndexOf("Teams is not installed", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string BuildAddinRepairExecutionLog(TeamsAddinRepairResult result)
        {
            var lines = new List<string>
            {
                $"Time: {DateTime.Now:yyyy-MM-dd HH:mm:ss}",
                "Scenario: TeamsAddinScenario",
                $"ExitCode: {result.ExitCode}",
                $"Success: {result.Success}",
                $"LogDirectory: {result.LogDirectory}",
                string.Empty,
                "[StandardOutput]",
                result.StandardOutput,
                string.Empty,
                "[StandardError]",
                result.StandardError
            };

            return string.Join(Environment.NewLine, lines);
        }

        private static string T(UiLanguage language, string zh, string en)
        {
            return LocalizationService.Localize(language, zh, en);
        }
    }

    public class TeamsMaintenanceResult
    {
        public bool Success { get; set; }

        public string PackageId { get; set; } = string.Empty;

        public string PackageFullName { get; set; } = string.Empty;

        public string ResetError { get; set; } = string.Empty;

        public string Message { get; set; } = string.Empty;

        public List<string> StoppedProcesses { get; } = new List<string>();

        public List<string> Warnings { get; } = new List<string>();

        public List<TeamsPathTarget> CacheTargets { get; } = new List<TeamsPathTarget>();

        public int DeletedDirectories { get; set; }

        public int DeletedFiles { get; set; }
    }

    public class TeamsSignInRecordCleanupResult
    {
        public bool Success { get; set; }

        public string Message { get; set; } = string.Empty;

        public List<string> StoppedProcesses { get; } = new List<string>();

        public List<string> Warnings { get; } = new List<string>();

        public List<TeamsPathTarget> ProcessedTargets { get; } = new List<TeamsPathTarget>();

        public int DeletedDirectories { get; set; }

        public int DeletedFiles { get; set; }
    }

    public class TeamsPathTarget
    {
        public TeamsPathTarget(string displayName, string path)
        {
            DisplayName = displayName;
            Path = path;
        }

        public string DisplayName { get; }

        public string Path { get; }
    }

    public class TeamsAddinRepairResult
    {
        public bool Success { get; set; }

        public int ExitCode { get; set; }

        public bool HasArgumentError { get; set; }

        public bool HasDiagnosticFailure { get; set; }

        public bool TeamsNotInstalledByDiagnostic { get; set; }

        public string StandardOutput { get; set; } = string.Empty;

        public string StandardError { get; set; } = string.Empty;

        public string LogDirectory { get; set; } = string.Empty;

        public string Message { get; set; } = string.Empty;
    }
}
