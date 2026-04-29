using System;
using System.IO;
using System.Text.Json;

namespace Office365CleanupTool.Services
{
    public sealed class AppSettings
    {
        public string LanguageMode { get; set; } = "Auto";

        public bool HasValidated21VAccount { get; set; }

        public string LastValidatedAccount { get; set; } = string.Empty;

        public string LastValidatedDisplayName { get; set; } = string.Empty;

        public string LastValidatedAvatarUrl { get; set; } = string.Empty;

        public DateTime? LastValidatedAtUtc { get; set; }
    }

    public sealed class AppSettingsService
    {
        private static readonly JsonSerializerOptions JsonOptions = new()
        {
            WriteIndented = true
        };

        private readonly string _settingsPath;

        public AppSettingsService()
        {
            string settingsDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "M365Tool");
            Directory.CreateDirectory(settingsDirectory);
            _settingsPath = Path.Combine(settingsDirectory, "settings.json");
        }

        public AppSettings Load()
        {
            try
            {
                if (!File.Exists(_settingsPath))
                {
                    return new AppSettings();
                }

                string json = File.ReadAllText(_settingsPath);
                AppSettings? settings = JsonSerializer.Deserialize<AppSettings>(json, JsonOptions);
                return settings ?? new AppSettings();
            }
            catch
            {
                return new AppSettings();
            }
        }

        public void Save(AppSettings settings)
        {
            if (settings == null)
            {
                throw new ArgumentNullException(nameof(settings));
            }

            string json = JsonSerializer.Serialize(settings, JsonOptions);
            File.WriteAllText(_settingsPath, json);
        }
    }
}
