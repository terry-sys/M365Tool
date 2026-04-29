namespace Office365CleanupTool.Models
{
    public class InstalledOfficeProduct
    {
        public string DisplayName { get; set; } = string.Empty;

        public string Version { get; set; } = string.Empty;

        public string ProductFamily { get; set; } = string.Empty;

        public string Source { get; set; } = string.Empty;

        public string Architecture { get; set; } = string.Empty;

        public string Channel { get; set; } = string.Empty;

        public string ToSummaryLine()
        {
            string detail = string.Empty;

            if (!string.IsNullOrWhiteSpace(Architecture))
            {
                detail = AppendDetail(detail, Architecture);
            }

            if (!string.IsNullOrWhiteSpace(Channel))
            {
                detail = AppendDetail(detail, Channel);
            }

            if (!string.IsNullOrWhiteSpace(Version))
            {
                detail = AppendDetail(detail, $"版本 {Version}");
            }

            if (!string.IsNullOrWhiteSpace(Source))
            {
                detail = AppendDetail(detail, Source);
            }

            return string.IsNullOrWhiteSpace(detail)
                ? DisplayName
                : $"{DisplayName} ({detail})";
        }

        private static string AppendDetail(string current, string value)
        {
            return string.IsNullOrWhiteSpace(current) ? value : $"{current} / {value}";
        }
    }
}
