namespace Office365CleanupTool.Models
{
    public class OfficeUninstallTargetOption
    {
        public string DisplayName { get; set; } = string.Empty;

        public string Description { get; set; } = string.Empty;

        public string? SaraOfficeVersion { get; set; }

        public bool UsesDetectedInstalledVersion => string.IsNullOrWhiteSpace(SaraOfficeVersion);

        public override string ToString() => DisplayName;
    }

    public class OfficeUninstallRequest
    {
        public OfficeUninstallTargetOption Target { get; set; } = new OfficeUninstallTargetOption();
    }

    public class OfficeUninstallResult
    {
        public int ExitCode { get; set; }

        public string Summary { get; set; } = string.Empty;

        public string Details { get; set; } = string.Empty;

        public string NextStepAdvice { get; set; } = string.Empty;

        public string DetectedOfficeSummary { get; set; } = string.Empty;

        public string LogDirectory { get; set; } = string.Empty;

        public string StandardOutput { get; set; } = string.Empty;

        public string StandardError { get; set; } = string.Empty;

        public string ProcessCleanupSummary { get; set; } = string.Empty;

        public string PostCheckOfficeSummary { get; set; } = string.Empty;

        public bool BackgroundUninstallReported { get; set; }

        public bool RequiresAttention { get; set; }

        public bool IsSuccess => ExitCode == 0 && !RequiresAttention;
    }
}
