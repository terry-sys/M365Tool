using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Win32;
using Office365CleanupTool.Models;

namespace Office365CleanupTool.Services
{
    public class OfficeChannelOption
    {
        public UpdateChannel Channel { get; set; }

        public string DisplayName { get; set; } = string.Empty;

        public string CdnBaseUrl { get; set; } = string.Empty;

        public override string ToString() => DisplayName;
    }

    public class OfficeVersionOption
    {
        public string Version { get; set; } = string.Empty;

        public string Build { get; set; } = string.Empty;

        public string ReleaseDate { get; set; } = string.Empty;

        public bool IsLatest { get; set; }

        public string DisplayName =>
            IsLatest
                ? $"Latest Version ({Version} / Build {Build})"
                : $"Version {Version} (Build {Build}) - {ReleaseDate}";

        public override string ToString() => DisplayName;
    }

    public class OfficeChannelService
    {
        private const string ClickToRunConfigPath = @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration";
        private const string ClickToRunUpdatesPath = @"SOFTWARE\Microsoft\Office\ClickToRun\Updates";
        private const string OfficeClientRelativePath = @"microsoft shared\ClickToRun\OfficeC2RClient.exe";
        private const string VersionHistoryUrl = "https://learn.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date";
        private const int UpdateProcessWaitTimeoutMs = 15000;

        public IReadOnlyList<OfficeChannelOption> GetSupportedChannels(UiLanguage language = UiLanguage.Chinese)
        {
            return new List<OfficeChannelOption>
            {
                new OfficeChannelOption
                {
                    Channel = UpdateChannel.Monthly,
                    DisplayName = T(language, "当前频道", "Current Channel"),
                    CdnBaseUrl = "https://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60"
                },
                new OfficeChannelOption
                {
                    Channel = UpdateChannel.Broad,
                    DisplayName = T(language, "半年频道", "Semi-Annual Channel"),
                    CdnBaseUrl = "https://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114"
                }
            };
        }

        public string GetOfficeClientPath()
        {
            var commonProgramFiles = Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles);
            return Path.Combine(commonProgramFiles, OfficeClientRelativePath);
        }

        public bool IsOfficeClickToRunAvailable()
        {
            return File.Exists(GetOfficeClientPath());
        }

        public async Task<IReadOnlyList<OfficeVersionOption>> GetAvailableVersionsAsync(UpdateChannel channel, int maxResults = 36)
        {
            string html = await DownloadVersionHistoryAsync();
            return ParseVersionOptions(html, channel, maxResults);
        }

        public void SwitchChannel(
            OfficeChannelOption channel,
            bool disableUpdates,
            OfficeVersionOption? targetVersion = null,
            UiLanguage language = UiLanguage.Chinese)
        {
            if (channel == null)
            {
                throw new ArgumentNullException(nameof(channel));
            }

            using (var configKey = Registry.LocalMachine.OpenSubKey(ClickToRunConfigPath, writable: true))
            {
                if (configKey == null)
                {
                    throw new InvalidOperationException(T(language, "找不到 Office ClickToRun 配置注册表。", "Office ClickToRun configuration registry key was not found."));
                }

                configKey.SetValue("CDNBaseUrl", channel.CdnBaseUrl, RegistryValueKind.String);
                // Ensure one update run can be triggered first; apply disable-after-switch later.
                configKey.SetValue("UpdatesEnabled", "True", RegistryValueKind.String);
            }

            using (var updatesKey = Registry.LocalMachine.OpenSubKey(ClickToRunUpdatesPath, writable: true))
            {
                if (updatesKey == null)
                {
                    throw new InvalidOperationException(T(language, "找不到 Office ClickToRun 更新注册表。", "Office ClickToRun updates registry key was not found."));
                }

                if (targetVersion == null || targetVersion.IsLatest)
                {
                    updatesKey.DeleteValue("UpdateToVersion", throwOnMissingValue: false);
                }
                else
                {
                    updatesKey.SetValue("UpdateToVersion", targetVersion.Build, RegistryValueKind.String);
                }
            }

            var clientPath = GetOfficeClientPath();
            if (!File.Exists(clientPath))
            {
                throw new FileNotFoundException(T(language, "找不到 OfficeC2RClient.exe。", "OfficeC2RClient.exe was not found."), clientPath);
            }

            string arguments = targetVersion != null && !targetVersion.IsLatest
                ? $"/update user updatetoversion={targetVersion.Build} displaylevel=true updatepromptuser=true forceappshutdown=true"
                : "/update user displaylevel=true updatepromptuser=true forceappshutdown=true";

            StartOfficeUpdateCommand(clientPath, arguments, language);

            if (disableUpdates)
            {
                using var configKey = Registry.LocalMachine.OpenSubKey(ClickToRunConfigPath, writable: true);
                if (configKey != null)
                {
                    configKey.SetValue("UpdatesEnabled", "False", RegistryValueKind.String);
                }
            }
        }

        private static async Task<string> DownloadVersionHistoryAsync()
        {
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.UserAgent.ParseAdd("M365Tool/1.0");
                using (var response = await client.GetAsync(VersionHistoryUrl))
                {
                    response.EnsureSuccessStatusCode();
                    return await response.Content.ReadAsStringAsync();
                }
            }
        }

        private static IReadOnlyList<OfficeVersionOption> ParseVersionOptions(string html, UpdateChannel channel, int maxResults)
        {
            var historyOptions = ParseVersionHistory(html, channel).ToList();
            var supportedOptions = ParseSupportedVersions(html, channel).ToList();

            var merged = new List<OfficeVersionOption>();
            merged.AddRange(supportedOptions);

            foreach (var option in historyOptions)
            {
                if (merged.All(existing => !string.Equals(existing.Build, option.Build, StringComparison.OrdinalIgnoreCase)))
                {
                    merged.Add(option);
                }
            }

            if (merged.Count == 0)
            {
                return Array.Empty<OfficeVersionOption>();
            }

            var latestSource = supportedOptions.FirstOrDefault() ?? historyOptions.First();
            var latestOption = new OfficeVersionOption
            {
                Version = latestSource.Version,
                Build = latestSource.Build,
                ReleaseDate = latestSource.ReleaseDate,
                IsLatest = true
            };

            var finalOptions = new List<OfficeVersionOption> { latestOption };
            foreach (var option in merged)
            {
                if (!string.Equals(option.Build, latestOption.Build, StringComparison.OrdinalIgnoreCase))
                {
                    option.IsLatest = false;
                    finalOptions.Add(option);
                }
            }

            return finalOptions.Take(maxResults + 1).ToList();
        }

        private static IEnumerable<OfficeVersionOption> ParseSupportedVersions(string html, UpdateChannel channel)
        {
            string normalizedHtml = Regex.Replace(WebUtility.HtmlDecode(html), @"\s+", " ");
            string channelPattern = channel == UpdateChannel.Monthly
                ? "Current Channel"
                : "Semi-Annual Enterprise Channel";

            var pattern = new Regex(
                $@"{Regex.Escape(channelPattern)}\s+(?<version>\d{{4}}\*?)\s+(?<build>\d+\.\d+)\s+(?<date>[A-Za-z]+\s+\d{{1,2}},\s+\d{{4}})",
                RegexOptions.IgnoreCase);

            foreach (Match match in pattern.Matches(normalizedHtml))
            {
                yield return new OfficeVersionOption
                {
                    Version = match.Groups["version"].Value.TrimEnd('*'),
                    Build = match.Groups["build"].Value,
                    ReleaseDate = match.Groups["date"].Value,
                    IsLatest = false
                };
            }
        }

        private static IEnumerable<OfficeVersionOption> ParseVersionHistory(string html, UpdateChannel channel)
        {
            string normalizedHtml = Regex.Replace(html, @"\s+", " ");
            var rowPattern = new Regex(
                @"<tr[^>]*>\s*<td[^>]*>(?<year>\d{4})</td>\s*<td[^>]*>(?<date>.*?)</td>\s*<td[^>]*>(?<current>.*?)</td>\s*<td[^>]*>(?<monthly>.*?)</td>\s*<td[^>]*>(?<preview>.*?)</td>\s*<td[^>]*>(?<semi>.*?)</td>\s*</tr>",
                RegexOptions.IgnoreCase);

            foreach (Match rowMatch in rowPattern.Matches(normalizedHtml))
            {
                string releaseDate = $"{CleanCellText(rowMatch.Groups["year"].Value)} {CleanCellText(rowMatch.Groups["date"].Value)}".Trim();
                string sourceCell = channel == UpdateChannel.Monthly
                    ? rowMatch.Groups["current"].Value
                    : rowMatch.Groups["semi"].Value;

                string channelCell = CleanCellText(sourceCell);
                foreach (Match versionMatch in Regex.Matches(channelCell, @"Version\s+(?<version>\d{4})\s+\(Build\s+(?<build>\d+\.\d+)\)", RegexOptions.IgnoreCase))
                {
                    yield return new OfficeVersionOption
                    {
                        Version = versionMatch.Groups["version"].Value,
                        Build = versionMatch.Groups["build"].Value,
                        ReleaseDate = releaseDate,
                        IsLatest = false
                    };
                }
            }
        }

        private static string CleanCellText(string html)
        {
            string noTags = Regex.Replace(html, "<.*?>", " ");
            string decoded = WebUtility.HtmlDecode(noTags);
            return Regex.Replace(decoded, @"\s+", " ").Trim();
        }

        private static string T(UiLanguage language, string zh, string en)
        {
            return LocalizationService.Localize(language, zh, en);
        }

        private static void StartOfficeUpdateCommand(string clientPath, string arguments, UiLanguage language)
        {
            using var process = new Process();
            process.StartInfo.FileName = clientPath;
            process.StartInfo.Arguments = arguments;
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;

            if (!process.Start())
            {
                throw new InvalidOperationException(
                    T(language, "无法启动 Office 更新命令。", "Unable to start Office update command."));
            }

            bool exited = process.WaitForExit(UpdateProcessWaitTimeoutMs);
            if (!exited)
            {
                // OfficeC2RClient may hand off update work to background processes.
                // If the command started successfully, treat update as triggered.
                return;
            }

            if (process.ExitCode is not 0 and not -1)
            {
                throw new InvalidOperationException(
                    T(language,
                        $"Office 更新命令执行失败，退出代码: {process.ExitCode}",
                        $"Office update command failed. Exit code: {process.ExitCode}"));
            }
        }
    }
}
