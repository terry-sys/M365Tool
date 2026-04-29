using System;
using System.IO;

namespace Office365CleanupTool.Services
{
    internal sealed class ToolSessionPaths
    {
        public ToolSessionPaths(string sessionRoot, string downloadsDirectory, string extractedDirectory, string logsDirectory)
        {
            SessionRoot = sessionRoot;
            DownloadsDirectory = downloadsDirectory;
            ExtractedDirectory = extractedDirectory;
            LogsDirectory = logsDirectory;
        }

        public string SessionRoot { get; }

        public string DownloadsDirectory { get; }

        public string ExtractedDirectory { get; }

        public string LogsDirectory { get; }
    }

    internal static class ToolPathService
    {
        private const string ProductFolderName = "M365Tool";

        public static string GetRootDirectory()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                ProductFolderName);
        }

        public static string GetModuleRootDirectory(string moduleName)
        {
            return Path.Combine(GetRootDirectory(), "Sessions", moduleName);
        }

        public static string GetOdtCacheDirectory()
        {
            string cacheDirectory = Path.Combine(GetRootDirectory(), "Cache", "ODT");
            Directory.CreateDirectory(cacheDirectory);
            return cacheDirectory;
        }

        public static ToolSessionPaths CreateSession(string moduleName, string? scenarioName = null)
        {
            string moduleRoot = string.IsNullOrWhiteSpace(scenarioName)
                ? Path.Combine(GetRootDirectory(), "Sessions", moduleName)
                : Path.Combine(GetRootDirectory(), "Sessions", moduleName, scenarioName);

            string sessionRoot = Path.Combine(moduleRoot, DateTime.Now.ToString("yyyyMMdd-HHmmss"));
            string downloadsDirectory = Path.Combine(sessionRoot, "Downloads");
            string extractedDirectory = Path.Combine(sessionRoot, "Extracted");
            string logsDirectory = Path.Combine(sessionRoot, "Logs");

            Directory.CreateDirectory(downloadsDirectory);
            Directory.CreateDirectory(extractedDirectory);
            Directory.CreateDirectory(logsDirectory);

            return new ToolSessionPaths(sessionRoot, downloadsDirectory, extractedDirectory, logsDirectory);
        }
    }
}
