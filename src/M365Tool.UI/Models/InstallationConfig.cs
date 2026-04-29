using System.Collections.Generic;
using System.IO;
using Office365CleanupTool.Services;

namespace Office365CleanupTool.Models
{
    public class InstallationConfig
    {
        public OfficeEdition Edition { get; set; }

        public OfficeArchitecture Architecture { get; set; }

        public UpdateChannel Channel { get; set; }

        public bool SilentInstall { get; set; } = true;

        public bool DisplayInstall { get; set; }

        public bool DownloadOnly { get; set; }

        public List<string> ExcludeApps { get; set; } = new List<string>();

        public List<string> LanguageIDs { get; set; } = new List<string>();

        public bool IncludeProject { get; set; }

        public bool IncludeVisio { get; set; }

        public bool SetFileFormat { get; set; }

        public bool ChangeArch { get; set; }

        public bool RemoveAllProducts { get; set; }

        public bool ForceOpenAppShutdown { get; set; }

        public bool OfficeMgmtCom { get; set; }

        public bool DeviceBasedLicensing { get; set; }

        public bool AllowCdnFallback { get; set; } = true;

        public string InstallVersion { get; set; } = string.Empty;

        private string _downloadPath = @"C:\Office365Install";

        public string DownloadPath
        {
            get => _downloadPath;
            set => _downloadPath = value ?? string.Empty;
        }

        public bool EnableUpdates { get; set; } = true;

        public bool PinItemsToTaskbar { get; set; } = true;

        public bool KeepMSI { get; set; }

        public bool SharedComputerLicensing { get; set; }

        public string SourcePath { get; set; } = string.Empty;

        public InstallationConfig()
        {
        }

        public InstallationConfig(OfficeVersionInfo versionInfo)
        {
            Edition = versionInfo.Edition;
            Architecture = versionInfo.Architecture;
            Channel = versionInfo.Channel;
        }

        public bool IsValid()
        {
            if (string.IsNullOrWhiteSpace(DownloadPath))
            {
                return false;
            }

            try
            {
                if (!Directory.Exists(DownloadPath))
                {
                    Directory.CreateDirectory(DownloadPath);
                }
            }
            catch
            {
                return false;
            }

            return true;
        }

        public string GetSummary(UiLanguage language = UiLanguage.Chinese)
        {
            string productName = IncludeProject
                ? "Project"
                : IncludeVisio
                    ? "Visio"
                    : new OfficeVersionInfo(Edition, Architecture, Channel).Name;

            string languageSummary = LanguageIDs.Count > 0
                ? string.Join(", ", LanguageIDs)
                : T(language, "匹配系统语言", "Match system language");
            string excludeSummary = ExcludeApps.Count > 0
                ? string.Join(", ", ExcludeApps)
                : T(language, "无", "None");

            return $"{T(language, "安装产品", "Product")}: {productName}\n" +
                   $"{T(language, "安装方式", "Install mode")}: {(SilentInstall ? T(language, "静默安装", "Silent install") : T(language, "显示安装界面", "Show installer UI"))}\n" +
                   $"{T(language, "下载路径", "Download path")}: {DownloadPath}\n" +
                   $"{T(language, "语言", "Language")}: {languageSummary}\n" +
                   $"{T(language, "排除应用", "Exclude apps")}: {excludeSummary}\n" +
                   $"{T(language, "固定到任务栏", "Pin to taskbar")}: {(PinItemsToTaskbar ? T(language, "是", "Yes") : T(language, "否", "No"))}";
        }

        public static List<string> GetAvailableApps()
        {
            return new List<string>
            {
                "Groove",
                "Outlook",
                "OneNote",
                "Access",
                "OneDrive",
                "Publisher",
                "Word",
                "Excel",
                "PowerPoint",
                "Teams",
                "Lync",
                "OutlookForWindows"
            };
        }

        public static string GetAppDisplayName(string appId)
        {
            return appId switch
            {
                "Groove" => "OneDrive (Groove)",
                "Outlook" => "Outlook",
                "OneNote" => "OneNote",
                "Access" => "Access",
                "OneDrive" => "OneDrive",
                "Publisher" => "Publisher",
                "Word" => "Word",
                "Excel" => "Excel",
                "PowerPoint" => "PowerPoint",
                "Teams" => "Teams",
                "Lync" => "Skype for Business (Lync)",
                "OutlookForWindows" => "New Outlook",
                _ => appId
            };
        }

        private static string T(UiLanguage language, string zh, string en)
        {
            return LocalizationService.Localize(language, zh, en);
        }
    }
}
