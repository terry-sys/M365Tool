using System;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Resources;

namespace Office365CleanupTool.Services
{
    public enum UiLanguage
    {
        Chinese,
        English
    }

    public static class LocalizationService
    {
        private const string ResourceSuffix = ".UiStrings.resources";
        private const uint FnvOffsetBasis = 2166136261;
        private const uint FnvPrime = 16777619;
        private static readonly Lazy<ResourceManager?> ResourceManager = new(ResolveResourceManager);

        public static UiLanguage ResolveLanguage(string? languageMode)
        {
            string normalized = (languageMode ?? string.Empty).Trim();
            if (string.Equals(normalized, "auto", StringComparison.OrdinalIgnoreCase))
            {
                return IsChineseCulture(CultureInfo.CurrentUICulture) ? UiLanguage.Chinese : UiLanguage.English;
            }

            return normalized.ToLowerInvariant() switch
            {
                "zh-cn" => UiLanguage.Chinese,
                "en-us" => UiLanguage.English,
                "zh" => UiLanguage.Chinese,
                "en" => UiLanguage.English,
                _ => IsChineseCulture(CultureInfo.CurrentUICulture) ? UiLanguage.Chinese : UiLanguage.English
            };
        }

        public static bool IsChineseCulture(CultureInfo culture)
        {
            return culture.TwoLetterISOLanguageName == "zh";
        }

        public static string Localize(UiLanguage language, string zh, string en)
        {
            // English should always prefer in-code English text and never fall back
            // to resource bundles, which may be incomplete or contain stale values.
            if (language == UiLanguage.English)
            {
                return string.IsNullOrWhiteSpace(en) ? zh : en;
            }

            if (string.IsNullOrWhiteSpace(en))
            {
                return zh;
            }

            string resourceKey = BuildResourceKey(en);
            CultureInfo culture = CultureInfo.GetCultureInfo("zh-CN");

            string? resourceText = ResourceManager.Value?.GetString(resourceKey, culture);
            if (!string.IsNullOrWhiteSpace(resourceText))
            {
                return resourceText;
            }

            return zh;
        }

        private static string BuildResourceKey(string en)
        {
            uint hash = FnvOffsetBasis;
            foreach (byte b in System.Text.Encoding.UTF8.GetBytes(en))
            {
                hash ^= b;
                hash *= FnvPrime;
            }

            return $"TXT_{hash:X8}";
        }

        private static ResourceManager? ResolveResourceManager()
        {
            try
            {
                Assembly assembly = typeof(LocalizationService).Assembly;
                string? resourceName = assembly
                    .GetManifestResourceNames()
                    .FirstOrDefault(name => name.EndsWith(ResourceSuffix, StringComparison.OrdinalIgnoreCase));

                if (string.IsNullOrWhiteSpace(resourceName))
                {
                    return null;
                }

                string baseName = resourceName[..^".resources".Length];
                return new ResourceManager(baseName, assembly);
            }
            catch
            {
                return null;
            }
        }
    }
}
