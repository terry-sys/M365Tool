using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
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
        private const string OfficeClientRelativePath = @"microsoft shared\ClickToRun\OfficeC2RClient.exe";
        private const string VersionHistoryUrl = "https://learn.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date";
        private const int UpdateProcessWaitTimeoutMs = 15000;

        private readonly IScriptRunner _scriptRunner;

        public OfficeChannelService(IScriptRunner? scriptRunner = null)
        {
            _scriptRunner = scriptRunner ?? new ScriptRunner();
        }

        internal static string GetChannelName(UpdateChannel channel) => channel switch
        {
            UpdateChannel.Monthly => "Current",
            UpdateChannel.MonthlyEnterprise => "MonthlyEnterprise",
            UpdateChannel.Broad => "SemiAnnual",
            _ => throw new ArgumentOutOfRangeException(nameof(channel))
        };

        internal static string GetChannelDisplayName(UpdateChannel channel, UiLanguage language = UiLanguage.Chinese) => channel switch
        {
            UpdateChannel.Monthly => T(language, "当前频道", "Current Channel"),
            UpdateChannel.MonthlyEnterprise => T(language, "月度企业频道", "Monthly Enterprise Channel"),
            UpdateChannel.Broad => T(language, "半年企业频道（兼容）", "Semi-Annual Enterprise (legacy)"),
            _ => throw new ArgumentOutOfRangeException(nameof(channel))
        };

        public IReadOnlyList<OfficeChannelOption> GetSupportedChannels(UiLanguage language = UiLanguage.Chinese)
        {
            return new List<OfficeChannelOption>
            {
                new OfficeChannelOption
                {
                    Channel = UpdateChannel.Monthly,
                    DisplayName = GetChannelDisplayName(UpdateChannel.Monthly, language),
                    CdnBaseUrl = "https://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60"
                },
                new OfficeChannelOption
                {
                    Channel = UpdateChannel.MonthlyEnterprise,
                    DisplayName = GetChannelDisplayName(UpdateChannel.MonthlyEnterprise, language),
                    CdnBaseUrl = "https://officecdn.microsoft.com/pr/55336b82-a18d-4dd6-b5f6-9e5095c314a6"
                },
                new OfficeChannelOption
                {
                    Channel = UpdateChannel.Broad,
                    DisplayName = GetChannelDisplayName(UpdateChannel.Broad, language),
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

        public virtual async Task<IReadOnlyList<OfficeVersionOption>> GetAvailableVersionsAsync(UpdateChannel channel, int maxResults = 36)
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
            using var operation = MaintenanceCoordinator.Shared.Begin(language);
            if (channel == null)
            {
                throw new ArgumentNullException(nameof(channel));
            }

            string clientPath = GetOfficeClientPath();
            if (!File.Exists(clientPath))
            {
                throw new FileNotFoundException(T(language, "找不到 OfficeC2RClient.exe。", "OfficeC2RClient.exe was not found."), clientPath);
            }

            // Stage the channel using Microsoft's supported ODT Updates configuration.
            // Do not overwrite install-time CDN values or higher-precedence organization policies.
            XDocument configuration = BuildChannelConfiguration(channel.Channel, disableUpdates, targetVersion);
            var session = ToolPathService.CreateSession("Channel");
            string setupPath = new OdtToolService(_scriptRunner).PrepareAsync(
                ToolPathService.GetOdtCacheDirectory(), CancellationToken.None, _ => { }, language).GetAwaiter().GetResult();
            ApplyChannelConfigurationAsync(setupPath, clientPath, configuration,
                session.SessionRoot, targetVersion, language).GetAwaiter().GetResult();
        }

        internal static XDocument BuildChannelConfiguration(UpdateChannel channel, bool disableUpdates, OfficeVersionOption? targetVersion)
        {
            string targetBuild = targetVersion != null && !targetVersion.IsLatest
                ? NormalizeBuildVersion(targetVersion.Build) : string.Empty;
            return new XDocument(new XElement("Configuration",
                new XElement("Updates",
                    new XAttribute("Channel", GetChannelName(channel)),
                    new XAttribute("Enabled", disableUpdates ? "FALSE" : "TRUE"),
                    // An empty TargetVersion clears a previously pinned build.
                    new XAttribute("TargetVersion", targetBuild))));
        }

        internal async Task ApplyChannelConfigurationAsync(string setupPath, string clientPath,
            XDocument configuration, string sessionDirectory, OfficeVersionOption? targetVersion, UiLanguage language)
        {
            string configurationPath = Path.Combine(sessionDirectory, "OfficeChannel.xml");
            configuration.Save(configurationPath);
            var result = await _scriptRunner.RunAsync(new ScriptExecutionRequest
            {
                FileName = setupPath,
                Arguments = $"/configure \"{configurationPath}\"",
                WorkingDirectory = sessionDirectory,
                Timeout = TimeSpan.FromMinutes(5),
                RedirectStandardOutput = false,
                RedirectStandardError = false
            });
            if (!result.IsSuccess)
            {
                throw new InvalidOperationException(result.IsStillRunning
                    ? T(language, "频道配置仍在后台运行，请等待完成后再检查更新。", "Channel configuration is still running. Wait for it to finish before checking for updates.")
                    : T(language, $"ODT 频道配置失败，退出代码: {result.ExitCode}", $"ODT channel configuration failed. Exit code: {result.ExitCode}"));
            }

            // Disabling automatic checks still permits this explicit manual update.
            string? targetBuild = targetVersion != null && !targetVersion.IsLatest ? targetVersion.Build : null;
            await StartOfficeUpdateCommandAsync(clientPath, BuildUpdateArguments(targetBuild), language);
        }

        internal static string NormalizeBuildVersion(string build)
        {
            string fullVersion = build.Trim();
            if (Regex.IsMatch(fullVersion, @"^\d+\.\d+$"))
            {
                fullVersion = "16.0." + fullVersion;
            }

            if (!System.Version.TryParse(fullVersion, out var version)
                || version.Major != 16 || version.Minor != 0
                || version.Build < 0 || version.Revision < 0)
            {
                throw new ArgumentException("Invalid Office build version.", nameof(build));
            }

            return version.ToString(4);
        }

        internal static string BuildUpdateArguments(string? targetBuild)
        {
            string version = targetBuild == null ? string.Empty : $" updatetoversion={NormalizeBuildVersion(targetBuild)}";
            return $"/update user{version} displaylevel=true updatepromptuser=true forceappshutdown=false";
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
            foreach (var rows in ReadTableRows(html))
            {
                string[] headers = rows[0];
                int channelIndex = Array.FindIndex(headers, value => value.Equals("Channel", StringComparison.OrdinalIgnoreCase));
                int versionIndex = Array.FindIndex(headers, value => value.Equals("Version", StringComparison.OrdinalIgnoreCase));
                int buildIndex = Array.FindIndex(headers, value => value.Equals("Build", StringComparison.OrdinalIgnoreCase));
                int dateIndex = Array.FindIndex(headers, value => value.Equals("Latest release date", StringComparison.OrdinalIgnoreCase));
                if (channelIndex < 0 || versionIndex < 0 || buildIndex < 0 || dateIndex < 0) continue;
                foreach (string[] cells in rows.Skip(1))
                {
                    if (cells.Length <= Math.Max(Math.Max(channelIndex, versionIndex), Math.Max(buildIndex, dateIndex))
                        || !MatchesChannel(cells[channelIndex], channel)) continue;
                    if (!Regex.IsMatch(cells[versionIndex], @"^\d{4}\*?$")
                        || !Regex.IsMatch(cells[buildIndex], @"^\d+\.\d+$")) continue;
                    yield return new OfficeVersionOption
                    {
                        Version = cells[versionIndex].TrimEnd('*'),
                        Build = cells[buildIndex],
                        ReleaseDate = cells[dateIndex]
                    };
                }
            }
        }

        private static IEnumerable<OfficeVersionOption> ParseVersionHistory(string html, UpdateChannel channel)
        {
            foreach (var rows in ReadTableRows(html))
            {
                string[] headers = rows[0];
                int yearIndex = Array.FindIndex(headers, value => value.Equals("Year", StringComparison.OrdinalIgnoreCase));
                int dateIndex = Array.FindIndex(headers, value => value.Equals("Release Date", StringComparison.OrdinalIgnoreCase));
                int channelIndex = Array.FindIndex(headers, value => MatchesChannel(value, channel));
                if (yearIndex < 0 || dateIndex < 0 || channelIndex < 0) continue;
                foreach (string[] cells in rows.Skip(1))
                {
                    if (cells.Length <= Math.Max(channelIndex, Math.Max(yearIndex, dateIndex))) continue;
                    string releaseDate = $"{cells[yearIndex]} {cells[dateIndex]}".Trim();
                    foreach (Match match in Regex.Matches(cells[channelIndex], @"Version\s+(?<version>\d{4})\s+\(Build\s+(?<build>\d+\.\d+)\)", RegexOptions.IgnoreCase))
                    {
                        yield return new OfficeVersionOption
                        {
                            Version = match.Groups["version"].Value,
                            Build = match.Groups["build"].Value,
                            ReleaseDate = releaseDate
                        };
                    }
                }
            }
        }

        private static bool MatchesChannel(string text, UpdateChannel channel)
        {
            string name = channel switch
            {
                UpdateChannel.Monthly => "Current Channel",
                UpdateChannel.MonthlyEnterprise => "Monthly Enterprise Channel",
                UpdateChannel.Broad => "Semi-Annual Enterprise Channel",
                _ => throw new ArgumentOutOfRangeException(nameof(channel))
            };
            return text.StartsWith(name, StringComparison.OrdinalIgnoreCase)
                && !text[name.Length..].TrimStart().StartsWith("(Preview)", StringComparison.OrdinalIgnoreCase);
        }

        private static IEnumerable<List<string[]>> ReadTableRows(string html)
        {
            foreach (Match table in Regex.Matches(html, @"<table\b[^>]*>(.*?)</table>", RegexOptions.IgnoreCase | RegexOptions.Singleline))
            {
                var rows = new List<string[]>();
                foreach (Match row in Regex.Matches(table.Groups[1].Value, @"<tr\b[^>]*>(.*?)</tr>", RegexOptions.IgnoreCase | RegexOptions.Singleline))
                {
                    string[] cells = Regex.Matches(row.Groups[1].Value, @"<t[hd]\b[^>]*>(.*?)</t[hd]>", RegexOptions.IgnoreCase | RegexOptions.Singleline)
                        .Select(cell => CleanCellText(cell.Groups[1].Value)).ToArray();
                    if (cells.Length > 0) rows.Add(cells);
                }
                if (rows.Count > 0) yield return rows;
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

        private async Task StartOfficeUpdateCommandAsync(string clientPath, string arguments, UiLanguage language)
        {
            var result = await _scriptRunner.RunAsync(new ScriptExecutionRequest
            {
                FileName = clientPath,
                Arguments = arguments,
                Timeout = TimeSpan.FromMilliseconds(UpdateProcessWaitTimeoutMs),
                RedirectStandardOutput = false,
                RedirectStandardError = false
            });

            // Keep the shared operation busy while a timed-out client is still running.
            if (result.IsStillRunning) return;
            if (result.ExitCode != 0)
            {
                throw new InvalidOperationException(
                    T(language,
                        $"Office 更新命令执行失败，退出代码: {result.ExitCode}",
                        $"Office update command failed. Exit code: {result.ExitCode}"));
            }
        }
    }
}
