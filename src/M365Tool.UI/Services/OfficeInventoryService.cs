using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Win32;
using Office365CleanupTool.Models;

namespace Office365CleanupTool.Services
{
    public class OfficeInventoryService
    {
        public IReadOnlyList<InstalledOfficeProduct> DetectInstalledProducts()
        {
            var items = new List<InstalledOfficeProduct>();

            CollectFromUninstallKey(RegistryView.Registry64, items);
            CollectFromUninstallKey(RegistryView.Registry32, items);
            CollectFromClickToRun(items);

            return items
                .GroupBy(
                    item => $"{item.ProductId}|{item.DisplayName}|{item.Version}|{item.Architecture}|{item.Channel}|{item.Source}",
                    StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .OrderBy(item => item.ProductFamily)
                .ThenBy(item => item.DisplayName)
                .ToList();
        }

        public string BuildSummary(IReadOnlyList<InstalledOfficeProduct> products)
        {
            if (products.Count == 0)
            {
                return "未检测到已安装的 Office / Visio / Project。";
            }

            return string.Join(Environment.NewLine, products.Select(item => $"• {item.ToSummaryLine()}"));
        }

        private static void CollectFromUninstallKey(RegistryView view, List<InstalledOfficeProduct> items)
        {
            using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view);
            using var uninstallKey = baseKey.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");
            if (uninstallKey == null)
            {
                return;
            }

            foreach (string subKeyName in uninstallKey.GetSubKeyNames())
            {
                using var subKey = uninstallKey.OpenSubKey(subKeyName);
                string? displayName = subKey?.GetValue("DisplayName") as string;
                string? publisher = subKey?.GetValue("Publisher") as string;
                if (!LooksLikeOfficeProduct(displayName, publisher) || IsOfficeRelatedAddIn(displayName))
                {
                    continue;
                }

                string version = subKey?.GetValue("DisplayVersion") as string ?? string.Empty;
                items.Add(new InstalledOfficeProduct
                {
                    DisplayName = displayName ?? string.Empty,
                    Version = version,
                    ProductId = subKeyName,
                    OfficeVersion = ResolveOfficeVersion(displayName, string.Empty, version),
                    ProductFamily = ResolveProductFamily(displayName),
                    Source = view == RegistryView.Registry64 ? "卸载注册表(64 位)" : "卸载注册表(32 位)"
                });
            }
        }

        private static void CollectFromClickToRun(List<InstalledOfficeProduct> items)
        {
            using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
            using var clickToRunKey = baseKey.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration");
            if (clickToRunKey == null)
            {
                return;
            }

            string productReleaseIds = clickToRunKey.GetValue("ProductReleaseIds") as string ?? string.Empty;
            string version = clickToRunKey.GetValue("VersionToReport") as string ?? string.Empty;
            string platform = clickToRunKey.GetValue("Platform") as string ?? string.Empty;
            string updateChannel = clickToRunKey.GetValue("UpdateChannel") as string ?? string.Empty;
            string cdnBaseUrl = clickToRunKey.GetValue("CDNBaseUrl") as string ?? string.Empty;

            foreach (string rawProductId in productReleaseIds.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string productId = rawProductId.Trim();
                if (string.IsNullOrWhiteSpace(productId))
                {
                    continue;
                }

                items.Add(new InstalledOfficeProduct
                {
                    DisplayName = GetClickToRunDisplayName(productId),
                    Version = version,
                    ProductId = productId,
                    OfficeVersion = ResolveOfficeVersion(string.Empty, productId, version),
                    ProductFamily = ResolveProductFamily(productId),
                    Source = "ClickToRun",
                    Architecture = string.IsNullOrWhiteSpace(platform) ? string.Empty : $"{platform} 位",
                    Channel = ResolveChannelDisplayName(updateChannel, cdnBaseUrl, version, productId)
                });
            }
        }

        internal static bool LooksLikeOfficeProduct(string? displayName, string? publisher = null)
        {
            if (string.IsNullOrWhiteSpace(displayName))
            {
                return false;
            }

            bool microsoftProduct = string.IsNullOrWhiteSpace(publisher)
                ? displayName.StartsWith("Microsoft ", StringComparison.OrdinalIgnoreCase)
                : publisher.Equals("Microsoft Corporation", StringComparison.OrdinalIgnoreCase)
                    || publisher.Equals("Microsoft", StringComparison.OrdinalIgnoreCase);

            return microsoftProduct && Regex.IsMatch(
                displayName, @"\b(?:Office|Visio|Project)\b|\bMicrosoft\s+365\b", RegexOptions.IgnoreCase);
        }

        internal static string ResolveOfficeVersion(string? displayName, string productId, string version)
        {
            if (productId.StartsWith("O365", StringComparison.OrdinalIgnoreCase)
                || productId.StartsWith("M365", StringComparison.OrdinalIgnoreCase)
                || Regex.IsMatch(displayName ?? string.Empty, @"\b(?:Microsoft|Office)\s*365\b", RegexOptions.IgnoreCase))
            {
                return "M365";
            }

            Match year = Regex.Match($"{productId} {displayName}", @"(?<!\d)(2024|2021|2019|2016|2013|2010|2007)(?!\d)");
            if (year.Success)
            {
                return year.Value;
            }

            if (System.Version.TryParse(version, out var parsedVersion))
            {
                return parsedVersion.Major switch
                {
                    12 => "2007",
                    14 => "2010",
                    15 => "2013",
                    _ => string.Empty
                };
            }

            return string.Empty;
        }

        private static bool IsOfficeRelatedAddIn(string? displayName)
        {
            if (string.IsNullOrWhiteSpace(displayName))
            {
                return false;
            }

            return displayName.Contains("Teams Meeting Add-in", StringComparison.OrdinalIgnoreCase);
        }

        private static string ResolveProductFamily(string? text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return "Office";
            }

            if (text.Contains("Visio", StringComparison.OrdinalIgnoreCase))
            {
                return "Visio";
            }

            if (text.Contains("Project", StringComparison.OrdinalIgnoreCase))
            {
                return "Project";
            }

            return "Office";
        }

        internal static string GetClickToRunDisplayName(string productId, UiLanguage language = UiLanguage.Chinese)
        {
            string id = productId.Trim();
            switch (id.ToLowerInvariant())
            {
                case "o365proplusretail":
                    return LocalizationService.Localize(language, "Microsoft 365 企业应用版", "Microsoft 365 Apps for enterprise");
                case "o365businessretail":
                    return LocalizationService.Localize(language, "Microsoft 365 商业应用版", "Microsoft 365 Apps for business");
            }

            if (id.Contains("Project", StringComparison.OrdinalIgnoreCase)) return "Project (ClickToRun)";
            if (id.Contains("Visio", StringComparison.OrdinalIgnoreCase)) return "Visio (ClickToRun)";
            if (id.StartsWith("O365", StringComparison.OrdinalIgnoreCase)
                || id.StartsWith("M365", StringComparison.OrdinalIgnoreCase))
                return $"Microsoft 365 Apps ({id})";

            return LocalizationService.Localize(language, $"Office 产品 ({id})", $"Office product ({id})");
        }

        internal static string ResolveChannelDisplayName(string? updateChannel, string? cdnBaseUrl = null, string? version = null, string? productId = null)
        {
            // UpdateChannel reflects the current update configuration; CDNBaseUrl is the install-time fallback.
            string value = string.IsNullOrWhiteSpace(updateChannel) ? cdnBaseUrl ?? string.Empty : updateChannel;
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;

            string id = value.Trim();
            if (Uri.TryCreate(id, UriKind.Absolute, out Uri? uri))
                id = uri.Segments.LastOrDefault()?.Trim('/') ?? string.Empty;

            // The old CDN ID remains valid after unification; only M365 Apps on 2606+ use the new experience.
            bool unified = ResolveOfficeVersion(string.Empty, productId ?? string.Empty, version ?? string.Empty) == "M365"
                && System.Version.TryParse(version, out var installedVersion)
                && installedVersion.Major == 16 && installedVersion.Minor == 0
                && installedVersion > new System.Version(16, 0, 20131, 20000);

            // Office stores CDN URLs containing channel GUIDs, rather than friendly channel names.
            // Mapping: learn.microsoft.com/intune/configmgr/sum/deploy-use/manage-office-365-proplus-updates
            return id.ToLowerInvariant() switch
            {
                "492350f6-3a01-4f97-b9c0-c7c6ddf67d60" or "current" or "monthly" => "当前频道",
                "7ffbc6bf-bc32-4f92-8982-f9dd17fd3114" or "broad" or "semiannual" => unified ? "月度企业频道（原半年配置）" : "半年企业频道",
                "55336b82-a18d-4dd6-b5f6-9e5095c314a6" or "monthlyenterprise" => "月度企业频道",
                "64256afe-f5d9-4f86-8936-8840a6a4f5be" or "currentpreview" or "targeted" => "当前频道（预览）",
                "b8f9b850-328d-4355-9145-c59439a0c4cf" or "semiannualpreview" or "semiannualtargeted" => "半年企业频道（预览）",
                "5440fd1f-7ecb-4221-8110-145efaa6372f" or "beta" or "insiderfast" => "Beta 频道",
                _ => "其他频道"
            };
        }
    }
}
