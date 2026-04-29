using System.Collections.Generic;
using System.Xml.Linq;
using Office365CleanupTool.Models;

namespace Office365CleanupTool.Services
{
    public static class OfficeDeploymentXmlBuilder
    {
        public static XDocument Build(InstallationConfig config, IReadOnlyList<string> excludeApps)
        {
            string channel = config.Channel == UpdateChannel.Monthly ? "Current" : "SemiAnnual";
            string officeEdition = config.Edition == OfficeEdition.Business ? "O365BusinessRetail" : "O365ProPlusRetail";
            string officeArch = config.Architecture == OfficeArchitecture.x64 ? "64" : "32";
            string displayLevel = config.SilentInstall ? "None" : "Full";

            var addElement = new XElement("Add",
                new XAttribute("OfficeClientEdition", officeArch),
                new XAttribute("Channel", channel));

            if (!string.IsNullOrWhiteSpace(config.SourcePath))
            {
                addElement.Add(new XAttribute("SourcePath", config.SourcePath));
            }

            if (!string.IsNullOrWhiteSpace(config.InstallVersion))
            {
                addElement.Add(new XAttribute("Version", config.InstallVersion));
            }

            if (config.AllowCdnFallback)
            {
                addElement.Add(new XAttribute("AllowCdnFallback", "True"));
            }

            if (config.ChangeArch)
            {
                addElement.Add(new XAttribute("MigrateArch", "TRUE"));
            }

            if (config.OfficeMgmtCom)
            {
                addElement.Add(new XAttribute("OfficeMgmtCOM", "True"));
            }

            string primaryProductId = config.IncludeProject
                ? "ProjectProRetail"
                : config.IncludeVisio
                    ? "VisioProRetail"
                    : officeEdition;

            addElement.Add(BuildProductElement(primaryProductId, config.LanguageIDs, excludeApps));

            var root = new XElement("Configuration", addElement);
            root.Add(new XElement("Property",
                new XAttribute("Name", "PinIconsToTaskbar"),
                new XAttribute("Value", config.PinItemsToTaskbar ? "TRUE" : "FALSE")));
            root.Add(new XElement("Property",
                new XAttribute("Name", "FORCEAPPSHUTDOWN"),
                new XAttribute("Value", config.ForceOpenAppShutdown ? "TRUE" : "FALSE")));
            root.Add(new XElement("Property",
                new XAttribute("Name", "SharedComputerLicensing"),
                new XAttribute("Value", config.SharedComputerLicensing ? "1" : "0")));
            root.Add(new XElement("Property",
                new XAttribute("Name", "DeviceBasedLicensing"),
                new XAttribute("Value", config.DeviceBasedLicensing ? "1" : "0")));
            root.Add(new XElement("Display",
                new XAttribute("Level", displayLevel),
                new XAttribute("AcceptEULA", "TRUE")));
            root.Add(new XElement("Updates",
                new XAttribute("Enabled", config.EnableUpdates ? "TRUE" : "FALSE")));

            if (!config.KeepMSI)
            {
                root.Add(new XElement("RemoveMSI"));
            }

            if (config.RemoveAllProducts)
            {
                root.Add(new XElement("Remove", new XAttribute("All", "TRUE")));
            }

            return new XDocument(new XDeclaration("1.0", "utf-8", "yes"), root);
        }

        public static InstallationConfig Parse(XDocument document)
        {
            XElement? root = document.Root;
            XElement? addElement = root?.Element("Add");
            XElement? displayElement = root?.Element("Display");
            XElement? updatesElement = root?.Element("Updates");

            if (root == null || addElement == null)
            {
                throw new InvalidOperationException("XML 中缺少 Configuration 或 Add 节点。");
            }

            var config = new InstallationConfig
            {
                Architecture = GetArchitecture(addElement.Attribute("OfficeClientEdition")?.Value),
                Channel = GetChannel(addElement.Attribute("Channel")?.Value),
                SourcePath = addElement.Attribute("SourcePath")?.Value ?? string.Empty,
                InstallVersion = addElement.Attribute("Version")?.Value ?? string.Empty,
                AllowCdnFallback = ParseBool(addElement.Attribute("AllowCdnFallback")?.Value, defaultValue: false),
                ChangeArch = ParseBool(addElement.Attribute("MigrateArch")?.Value),
                OfficeMgmtCom = ParseBool(addElement.Attribute("OfficeMgmtCOM")?.Value),
                SilentInstall = string.Equals(displayElement?.Attribute("Level")?.Value, "None", System.StringComparison.OrdinalIgnoreCase),
                EnableUpdates = ParseBool(updatesElement?.Attribute("Enabled")?.Value, defaultValue: true),
                KeepMSI = root.Element("RemoveMSI") == null
            };

            config.DisplayInstall = !config.SilentInstall;

            foreach (XElement property in root.Elements("Property"))
            {
                string? name = property.Attribute("Name")?.Value;
                string? value = property.Attribute("Value")?.Value;

                switch (name)
                {
                    case "PinIconsToTaskbar":
                        config.PinItemsToTaskbar = ParseBool(value, defaultValue: true);
                        break;
                    case "FORCEAPPSHUTDOWN":
                        config.ForceOpenAppShutdown = ParseBool(value);
                        break;
                    case "SharedComputerLicensing":
                        config.SharedComputerLicensing = ParseBool(value);
                        break;
                    case "DeviceBasedLicensing":
                        config.DeviceBasedLicensing = ParseBool(value);
                        break;
                }
            }

            XElement? removeElement = root.Element("Remove");
            if (removeElement != null && string.Equals(removeElement.Attribute("All")?.Value, "TRUE", System.StringComparison.OrdinalIgnoreCase))
            {
                config.RemoveAllProducts = true;
            }

            foreach (XElement product in addElement.Elements("Product"))
            {
                string? productId = product.Attribute("ID")?.Value;
                if (string.IsNullOrWhiteSpace(productId))
                {
                    continue;
                }

                switch (productId)
                {
                    case "O365BusinessRetail":
                        config.Edition = OfficeEdition.Business;
                        AddLanguagesAndExcludedApps(config, product);
                        break;
                    case "O365ProPlusRetail":
                        config.Edition = OfficeEdition.Enterprise;
                        AddLanguagesAndExcludedApps(config, product);
                        break;
                    case "ProjectProRetail":
                        config.IncludeProject = true;
                        break;
                    case "VisioProRetail":
                        config.IncludeVisio = true;
                        break;
                }
            }

            return config;
        }

        private static XElement BuildProductElement(string productId, IReadOnlyList<string> languageIds, IReadOnlyList<string> excludeApps)
        {
            var product = new XElement("Product", new XAttribute("ID", productId));

            if (languageIds.Count > 0)
            {
                foreach (string languageId in languageIds)
                {
                    product.Add(new XElement("Language", new XAttribute("ID", languageId)));
                }
            }
            else
            {
                product.Add(new XElement("Language", new XAttribute("ID", "MatchOS")));
            }

            foreach (string excludeApp in excludeApps)
            {
                product.Add(new XElement("ExcludeApp", new XAttribute("ID", excludeApp)));
            }

            return product;
        }

        private static void AddLanguagesAndExcludedApps(InstallationConfig config, XElement product)
        {
            config.LanguageIDs.Clear();
            foreach (XElement language in product.Elements("Language"))
            {
                string? id = language.Attribute("ID")?.Value;
                if (!string.IsNullOrWhiteSpace(id) && !string.Equals(id, "MatchOS", System.StringComparison.OrdinalIgnoreCase))
                {
                    config.LanguageIDs.Add(id);
                }
            }

            config.ExcludeApps.Clear();
            foreach (XElement exclude in product.Elements("ExcludeApp"))
            {
                string? id = exclude.Attribute("ID")?.Value;
                if (!string.IsNullOrWhiteSpace(id))
                {
                    config.ExcludeApps.Add(id);
                }
            }
        }

        private static OfficeArchitecture GetArchitecture(string? officeClientEdition)
        {
            return officeClientEdition == "32" ? OfficeArchitecture.x86 : OfficeArchitecture.x64;
        }

        private static UpdateChannel GetChannel(string? channel)
        {
            return string.Equals(channel, "SemiAnnual", System.StringComparison.OrdinalIgnoreCase)
                ? UpdateChannel.Broad
                : UpdateChannel.Monthly;
        }

        private static bool ParseBool(string? value, bool defaultValue = false)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return defaultValue;
            }

            return value == "1" ||
                   value.Equals("TRUE", System.StringComparison.OrdinalIgnoreCase) ||
                   value.Equals("True", System.StringComparison.OrdinalIgnoreCase);
        }
    }
}
