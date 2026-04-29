using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Security.Principal;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace Office365CleanupTool.Services
{
    public class ValidationService
    {
        public async Task<bool> ValidateSystemRequirementsAsync(UiLanguage language = UiLanguage.Chinese)
        {
            ValidationReport report = await GetValidationReportAsync(language);
            return report.IsSuccess;
        }

        public async Task<ValidationReport> GetValidationReportAsync(UiLanguage language = UiLanguage.Chinese)
        {
            var items = new List<ValidationCheckResult>();

            items.Add(CreateOperatingSystemValidationResult(language));

            bool isAdmin = ValidateAdministratorPrivileges();
            items.Add(new ValidationCheckResult(
                T(language, "管理员权限", "Administrator Privileges"),
                isAdmin,
                isAdmin
                    ? T(language, "当前已使用管理员权限运行。", "Currently running with administrator privileges.")
                    : T(language, "安装 Office 建议使用管理员权限运行工具。", "Office installation is recommended to run with administrator privileges.")));

            double diskSpaceGb = await GetAvailableDiskSpaceGbAsync();
            bool diskValid = diskSpaceGb >= 3.0;
            items.Add(new ValidationCheckResult(
                T(language, "磁盘空间", "Disk Space"),
                diskValid,
                diskValid
                    ? T(language, $"当前可用空间约 {diskSpaceGb:F1} GB。", $"Available space is about {diskSpaceGb:F1} GB.")
                    : T(language, $"当前可用空间约 {diskSpaceGb:F1} GB，建议至少保留 3 GB。", $"Available space is about {diskSpaceGb:F1} GB, at least 3 GB is recommended.")));

            (bool networkValid, string networkDetail) = await ValidateNetworkConnectionAsync(language);
            items.Add(new ValidationCheckResult(
                T(language, "网络连接", "Network Connectivity"),
                networkValid,
                networkDetail));

            int powerShellVersion = await GetPowerShellMajorVersionAsync();
            bool powerShellValid = powerShellVersion >= 5;
            items.Add(new ValidationCheckResult(
                "PowerShell",
                powerShellValid,
                powerShellValid
                    ? T(language, $"当前 PowerShell 主版本为 {powerShellVersion}。", $"Current PowerShell major version is {powerShellVersion}.")
                    : T(language, $"当前 PowerShell 主版本为 {powerShellVersion}，建议至少为 5。", $"Current PowerShell major version is {powerShellVersion}, version 5 or later is recommended.")));

            return new ValidationReport(items);
        }

        private ValidationCheckResult CreateOperatingSystemValidationResult(UiLanguage language)
        {
            string itemName = T(language, "操作系统", "Operating System");

            try
            {
                WindowsVersionInfo versionInfo = GetWindowsVersionInfo();

                if (versionInfo.IsWindows11OrNewer)
                {
                    return new ValidationCheckResult(
                        itemName,
                        ValidationSeverity.Success,
                        T(language, $"已满足最低系统要求（当前系统：{versionInfo.DisplayName}）。", $"Meets the minimum system requirement (current system: {versionInfo.DisplayName})."));
                }

                if (versionInfo.IsWindows10)
                {
                    return new ValidationCheckResult(
                        itemName,
                        ValidationSeverity.Warning,
                        T(
                            language,
                            $"当前系统为 {versionInfo.DisplayName}，不符合最低系统要求，但仍可继续安装。Windows 10 已不在官方支持列表中。",
                            $"Current system is {versionInfo.DisplayName}. It does not meet the minimum system requirement, but installation can continue. Windows 10 is no longer on the official support list."));
                }

                return new ValidationCheckResult(
                    itemName,
                    ValidationSeverity.Error,
                    string.IsNullOrWhiteSpace(versionInfo.DisplayName)
                        ? T(language, "需要 Windows 11 或更高版本。", "Requires Windows 11 or later.")
                        : T(language, $"当前系统为 {versionInfo.DisplayName}，需要 Windows 11 或更高版本。", $"Current system is {versionInfo.DisplayName}; Windows 11 or later is required."));
            }
            catch
            {
                return new ValidationCheckResult(
                    itemName,
                    ValidationSeverity.Error,
                    T(language, "无法识别当前 Windows 版本，需要 Windows 11 或更高版本。", "Unable to identify the current Windows version. Windows 11 or later is required."));
            }
        }

        private static WindowsVersionInfo GetWindowsVersionInfo()
        {
            string productName = string.Empty;
            string editionId = string.Empty;
            string displayVersion = string.Empty;
            int buildNumber = 0;

            try
            {
                using RegistryKey? key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion");
                productName = key?.GetValue("ProductName")?.ToString() ?? string.Empty;
                editionId = key?.GetValue("EditionID")?.ToString() ?? string.Empty;
                displayVersion = key?.GetValue("DisplayVersion")?.ToString() ?? string.Empty;
                _ = int.TryParse(key?.GetValue("CurrentBuildNumber")?.ToString(), out buildNumber);
            }
            catch
            {
                productName = string.Empty;
                editionId = string.Empty;
                displayVersion = string.Empty;
                buildNumber = 0;
            }

            Version version = Environment.OSVersion.Version;
            bool isServer = productName.Contains("Server", StringComparison.OrdinalIgnoreCase);
            bool isWindows11ByName = productName.Contains("Windows 11", StringComparison.OrdinalIgnoreCase);
            bool isWindows10ByName = productName.Contains("Windows 10", StringComparison.OrdinalIgnoreCase);
            bool isWindows11ByBuild = version.Major >= 10 && buildNumber >= 22000;
            bool isWindows10ByBuild = version.Major >= 10 && buildNumber > 0 && buildNumber < 22000;
            bool isWindows11OrNewer = isWindows11ByName || isWindows11ByBuild;
            bool isWindows10 = !isServer && !isWindows11OrNewer && (isWindows10ByName || isWindows10ByBuild || version.Major == 10);

            string displayName = BuildWindowsDisplayName(
                productName,
                editionId,
                displayVersion,
                buildNumber,
                version,
                isServer,
                isWindows10,
                isWindows11OrNewer);

            return new WindowsVersionInfo(displayName, isWindows10, isWindows11OrNewer);
        }

        private static string BuildWindowsDisplayName(
            string productName,
            string editionId,
            string displayVersion,
            int buildNumber,
            Version version,
            bool isServer,
            bool isWindows10,
            bool isWindows11OrNewer)
        {
            if (isServer && !string.IsNullOrWhiteSpace(productName))
            {
                return AppendVersionSegments(productName, displayVersion, buildNumber);
            }

            string baseName = isWindows11OrNewer
                ? "Windows 11"
                : isWindows10
                    ? "Windows 10"
                    : !string.IsNullOrWhiteSpace(productName)
                        ? productName
                        : $"Windows {version.Major}.{version.Minor}";

            string editionName = GetEditionDisplayName(editionId, productName, baseName);
            string fullName = string.IsNullOrWhiteSpace(editionName) ? baseName : $"{baseName} {editionName}";
            return AppendVersionSegments(fullName, displayVersion, buildNumber);
        }

        private static string GetEditionDisplayName(string editionId, string productName, string baseName)
        {
            if (!string.IsNullOrWhiteSpace(editionId))
            {
                return editionId.Replace("Core", "Home", StringComparison.OrdinalIgnoreCase);
            }

            if (string.IsNullOrWhiteSpace(productName))
            {
                return string.Empty;
            }

            string normalized = productName.Trim();
            if (normalized.StartsWith(baseName + " ", StringComparison.OrdinalIgnoreCase))
            {
                return normalized.Substring(baseName.Length).Trim();
            }

            return string.Empty;
        }

        private static string AppendVersionSegments(string name, string displayVersion, int buildNumber)
        {
            string result = name;

            if (!string.IsNullOrWhiteSpace(displayVersion) &&
                !result.Contains(displayVersion, StringComparison.OrdinalIgnoreCase))
            {
                result += $" {displayVersion}";
            }

            if (buildNumber > 0)
            {
                result += $" (Build {buildNumber})";
            }

            return result;
        }

        private bool ValidateAdministratorPrivileges()
        {
            try
            {
                using var identity = WindowsIdentity.GetCurrent();
                var principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch
            {
                return false;
            }
        }

        private async Task<double> GetAvailableDiskSpaceGbAsync()
        {
            return await Task.Run(() =>
            {
                try
                {
                    var drive = new DriveInfo(Path.GetPathRoot(Environment.SystemDirectory) ?? string.Empty);
                    return drive.AvailableFreeSpace / (1024.0 * 1024 * 1024);
                }
                catch
                {
                    return 0d;
                }
            });
        }

        private async Task<(bool IsSuccess, string Detail)> ValidateNetworkConnectionAsync(UiLanguage language)
        {
            try
            {
                using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(8) };
                using var response = await client.GetAsync("https://www.microsoft.com");

                if (response.IsSuccessStatusCode)
                {
                    return (true, T(language, "可以访问 Microsoft 网络。", "Can access Microsoft network."));
                }

                return
                (
                    false,
                    T(
                        language,
                        $"无法访问 Microsoft 网络，状态码：{(int)response.StatusCode} {response.StatusCode}。",
                        $"Cannot access Microsoft network. Status code: {(int)response.StatusCode} {response.StatusCode}.")
                );
            }
            catch (Exception ex)
            {
                return
                (
                    false,
                    T(
                        language,
                        $"无法访问 Microsoft 网络，可能影响 ODT 下载与安装。底层错误：{ex.Message}",
                        $"Cannot access Microsoft network, which may affect ODT download and installation. Error: {ex.Message}")
                );
            }
        }

        private async Task<int> GetPowerShellMajorVersionAsync()
        {
            return await Task.Run(() =>
            {
                try
                {
                    using var process = new System.Diagnostics.Process();
                    process.StartInfo.FileName = "powershell.exe";
                    process.StartInfo.Arguments = "-Command \"$PSVersionTable.PSVersion.Major\"";
                    process.StartInfo.UseShellExecute = false;
                    process.StartInfo.RedirectStandardOutput = true;
                    process.StartInfo.CreateNoWindow = true;

                    process.Start();
                    string output = process.StandardOutput.ReadToEnd();
                    process.WaitForExit();

                    return int.TryParse(output.Trim(), out int version) ? version : 0;
                }
                catch
                {
                    return 0;
                }
            });
        }

        public bool ValidateInstallationPath(string path)
        {
            try
            {
                if (string.IsNullOrEmpty(path))
                {
                    return false;
                }

                if (!Path.IsPathRooted(path))
                {
                    return false;
                }

                if (path.Length > 260)
                {
                    return false;
                }

                return path.IndexOfAny(Path.GetInvalidPathChars()) < 0;
            }
            catch
            {
                return false;
            }
        }

        public string GetSystemInfo(UiLanguage language = UiLanguage.Chinese)
        {
            try
            {
                var info = new System.Text.StringBuilder();
                info.AppendLine(T(language, "操作系统", "Operating System") + $": {Environment.OSVersion}");
                info.AppendLine(T(language, "系统架构", "OS Architecture") + $": {(Environment.Is64BitOperatingSystem ? "64-bit" : "32-bit")}");
                info.AppendLine(T(language, "处理器架构", "Processor Architecture") + $": {Environment.GetEnvironmentVariable("PROCESSOR_ARCHITECTURE")}");
                info.AppendLine($".NET {T(language, "版本", "Version")}: {Environment.Version}");
                return info.ToString();
            }
            catch (Exception ex)
            {
                return T(language, $"获取系统信息失败: {ex.Message}", $"Failed to get system information: {ex.Message}");
            }
        }

        private static string T(UiLanguage language, string zh, string en)
        {
            return language == UiLanguage.Chinese ? zh : en;
        }
    }

    public class ValidationReport
    {
        public IReadOnlyList<ValidationCheckResult> Items { get; }

        public bool IsSuccess { get; }

        public bool HasAttentionItems { get; }

        public bool HasBlockingItems { get; }

        public ValidationReport(IReadOnlyList<ValidationCheckResult> items)
        {
            Items = items;
            IsSuccess = true;
            HasAttentionItems = false;
            HasBlockingItems = false;
            foreach (ValidationCheckResult item in items)
            {
                if (item.RequiresAttention)
                {
                    HasAttentionItems = true;
                }

                if (item.IsBlocking)
                {
                    IsSuccess = false;
                    HasBlockingItems = true;
                }
            }
        }
    }

    public enum ValidationSeverity
    {
        Success,
        Warning,
        Error
    }

    public class ValidationCheckResult
    {
        public string Name { get; }

        public bool IsSuccess { get; }

        public bool RequiresAttention { get; }

        public bool IsBlocking { get; }

        public ValidationSeverity Severity { get; }

        public string Detail { get; }

        public ValidationCheckResult(string name, bool isSuccess, string detail)
            : this(name, isSuccess ? ValidationSeverity.Success : ValidationSeverity.Error, detail)
        {
        }

        public ValidationCheckResult(string name, ValidationSeverity severity, string detail)
        {
            Name = name;
            Severity = severity;
            IsSuccess = severity == ValidationSeverity.Success;
            RequiresAttention = severity != ValidationSeverity.Success;
            IsBlocking = severity == ValidationSeverity.Error;
            Detail = detail;
        }
    }

    internal sealed class WindowsVersionInfo
    {
        public WindowsVersionInfo(string displayName, bool isWindows10, bool isWindows11OrNewer)
        {
            DisplayName = displayName;
            IsWindows10 = isWindows10;
            IsWindows11OrNewer = isWindows11OrNewer;
        }

        public string DisplayName { get; }

        public bool IsWindows10 { get; }

        public bool IsWindows11OrNewer { get; }
    }
}
