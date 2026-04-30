using System;
using System.Collections.Generic;
using System.Linq;
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
                    item => $"{item.DisplayName}|{item.Version}|{item.Architecture}|{item.Channel}|{item.Source}",
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
                if (!LooksLikeOfficeProduct(displayName) || IsOfficeRelatedAddIn(displayName))
                {
                    continue;
                }

                items.Add(new InstalledOfficeProduct
                {
                    DisplayName = displayName ?? string.Empty,
                    Version = subKey?.GetValue("DisplayVersion") as string ?? string.Empty,
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
            string updateChannel = clickToRunKey.GetValue("CDNBaseUrl") as string ?? string.Empty;

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
                    ProductFamily = ResolveProductFamily(productId),
                    Source = "ClickToRun",
                    Architecture = string.IsNullOrWhiteSpace(platform) ? string.Empty : $"{platform} 位",
                    Channel = ResolveChannelDisplayName(updateChannel)
                });
            }
        }

        private static bool LooksLikeOfficeProduct(string? displayName)
        {
            if (string.IsNullOrWhiteSpace(displayName))
            {
                return false;
            }

            return displayName.Contains("Office", StringComparison.OrdinalIgnoreCase)
                || displayName.Contains("Microsoft 365", StringComparison.OrdinalIgnoreCase)
                || displayName.Contains("Visio", StringComparison.OrdinalIgnoreCase)
                || displayName.Contains("Project", StringComparison.OrdinalIgnoreCase);
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

        private static string GetClickToRunDisplayName(string productId)
        {
            if (productId.Contains("Project", StringComparison.OrdinalIgnoreCase))
            {
                return "Project (ClickToRun)";
            }

            if (productId.Contains("Visio", StringComparison.OrdinalIgnoreCase))
            {
                return "Visio (ClickToRun)";
            }

            if (productId.Contains("O365", StringComparison.OrdinalIgnoreCase)
                || productId.Contains("ProPlus", StringComparison.OrdinalIgnoreCase))
            {
                return "Microsoft 365 Apps (ClickToRun)";
            }

            return $"Office 产品 ({productId})";
        }

        private static string ResolveChannelDisplayName(string updateChannel)
        {
            if (string.IsNullOrWhiteSpace(updateChannel))
            {
                return string.Empty;
            }

            if (updateChannel.Contains("Current", StringComparison.OrdinalIgnoreCase))
            {
                return "当前频道";
            }

            if (updateChannel.Contains("Broad", StringComparison.OrdinalIgnoreCase)
                || updateChannel.Contains("SemiAnnual", StringComparison.OrdinalIgnoreCase))
            {
                return "半年频道";
            }

            return "其他频道";
        }
    }
}
