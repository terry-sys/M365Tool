using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using Office365CleanupTool.Models;

namespace Office365CleanupTool.Services
{
    public class InstallationPresetService
    {
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true
        };

        public string RecentConfigPath => Path.Combine(GetPresetDirectory(), "recent-installation-config.json");

        public string GetPresetDirectory()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "M365Tool",
                "InstallPresets");
        }

        public async Task SaveTemplateAsync(InstallationConfig config, string filePath)
        {
            string json = JsonSerializer.Serialize(config, JsonOptions);
            string? directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            await File.WriteAllTextAsync(filePath, json);
        }

        public async Task<InstallationConfig?> LoadTemplateAsync(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return null;
            }

            string json = await File.ReadAllTextAsync(filePath);
            return JsonSerializer.Deserialize<InstallationConfig>(json, JsonOptions);
        }

        public async Task SaveRecentConfigAsync(InstallationConfig config)
        {
            Directory.CreateDirectory(GetPresetDirectory());
            await SaveTemplateAsync(config, RecentConfigPath);
        }

        public async Task<InstallationConfig?> LoadRecentConfigAsync()
        {
            return await LoadTemplateAsync(RecentConfigPath);
        }

        public async Task SaveNamedTemplateAsync(string name, string remark, InstallationConfig config)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentException("模板名称不能为空。", nameof(name));
            }

            Directory.CreateDirectory(GetPresetDirectory());
            string filePath = GetNamedTemplatePath(name);

            var template = new InstallationTemplate
            {
                Name = name.Trim(),
                Remark = remark.Trim(),
                CreatedAt = DateTime.Now,
                Config = config
            };

            string json = JsonSerializer.Serialize(template, JsonOptions);
            await File.WriteAllTextAsync(filePath, json);
        }

        public async Task<InstallationTemplate?> LoadNamedTemplateAsync(string name)
        {
            string filePath = GetNamedTemplatePath(name);
            if (!File.Exists(filePath))
            {
                return null;
            }

            string json = await File.ReadAllTextAsync(filePath);
            return JsonSerializer.Deserialize<InstallationTemplate>(json, JsonOptions);
        }

        public IReadOnlyList<InstallationTemplateListItem> ListTemplates()
        {
            string directory = GetPresetDirectory();
            if (!Directory.Exists(directory))
            {
                return Array.Empty<InstallationTemplateListItem>();
            }

            var items = new List<InstallationTemplateListItem>();
            foreach (string filePath in Directory.GetFiles(directory, "*.json"))
            {
                if (string.Equals(Path.GetFileName(filePath), "recent-installation-config.json", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                try
                {
                    string json = File.ReadAllText(filePath);
                    InstallationTemplate? template = JsonSerializer.Deserialize<InstallationTemplate>(json, JsonOptions);
                    if (template == null)
                    {
                        continue;
                    }

                    items.Add(new InstallationTemplateListItem
                    {
                        Name = template.Name,
                        Remark = template.Remark,
                        CreatedAt = template.CreatedAt,
                        FilePath = filePath
                    });
                }
                catch
                {
                    // Ignore unreadable template files.
                }
            }

            return items
                .OrderByDescending(item => item.CreatedAt)
                .ThenBy(item => item.Name)
                .ToList();
        }

        public bool DeleteNamedTemplate(string name)
        {
            string filePath = GetNamedTemplatePath(name);
            if (!File.Exists(filePath))
            {
                return false;
            }

            File.Delete(filePath);
            return true;
        }

        private string GetNamedTemplatePath(string name)
        {
            string safeFileName = string.Join(
                "_",
                name.Split(Path.GetInvalidFileNameChars(), StringSplitOptions.RemoveEmptyEntries)).Trim();

            if (string.IsNullOrWhiteSpace(safeFileName))
            {
                safeFileName = "Template";
            }

            return Path.Combine(GetPresetDirectory(), $"{safeFileName}.json");
        }
    }

    public class InstallationTemplateListItem
    {
        public string Name { get; set; } = string.Empty;

        public string Remark { get; set; } = string.Empty;

        public DateTime CreatedAt { get; set; }

        public string FilePath { get; set; } = string.Empty;

        public override string ToString()
        {
            string remark = string.IsNullOrWhiteSpace(Remark) ? "无备注" : Remark;
            return $"{Name} ({CreatedAt:yyyy-MM-dd HH:mm}) - {remark}";
        }
    }
}
