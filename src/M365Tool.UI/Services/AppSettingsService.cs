using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
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

        public string LastValidatedAvatarPath { get; set; } = string.Empty;

        public DateTime? LastValidatedAtUtc { get; set; }
    }

    public sealed class AppSettingsService
    {
        private const string ProtectedSettingsFormat = "M365Tool.AppSettings.Protected.v1";

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
                AppSettings? settings = TryReadProtectedSettings(json) ??
                                        JsonSerializer.Deserialize<AppSettings>(json, JsonOptions);
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
            byte[] plainBytes = Encoding.UTF8.GetBytes(json);
            byte[] protectedBytes = ProtectedData.Protect(
                plainBytes,
                optionalEntropy: null,
                DataProtectionScope.CurrentUser);

            var protectedFile = new ProtectedAppSettingsFile
            {
                Format = ProtectedSettingsFormat,
                ProtectedData = Convert.ToBase64String(protectedBytes)
            };
            string protectedJson = JsonSerializer.Serialize(protectedFile, JsonOptions);
            File.WriteAllText(_settingsPath, protectedJson);
        }

        private static AppSettings? TryReadProtectedSettings(string json)
        {
            try
            {
                ProtectedAppSettingsFile? protectedFile = JsonSerializer.Deserialize<ProtectedAppSettingsFile>(json, JsonOptions);
                if (protectedFile == null ||
                    !string.Equals(protectedFile.Format, ProtectedSettingsFormat, StringComparison.Ordinal) ||
                    string.IsNullOrWhiteSpace(protectedFile.ProtectedData))
                {
                    return null;
                }

                byte[] protectedBytes = Convert.FromBase64String(protectedFile.ProtectedData);
                byte[] plainBytes = ProtectedData.Unprotect(
                    protectedBytes,
                    optionalEntropy: null,
                    DataProtectionScope.CurrentUser);
                string plainJson = Encoding.UTF8.GetString(plainBytes);
                return JsonSerializer.Deserialize<AppSettings>(plainJson, JsonOptions);
            }
            catch
            {
                return null;
            }
        }

        private sealed class ProtectedAppSettingsFile
        {
            public string Format { get; set; } = string.Empty;

            public string ProtectedData { get; set; } = string.Empty;
        }
    }
}
