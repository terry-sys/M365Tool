namespace Office365CleanupTool.Models
{
    public enum OfficeEdition
    {
        Business,
        Enterprise
    }

    public enum OfficeArchitecture
    {
        x86,
        x64
    }

    public enum UpdateChannel
    {
        Monthly,
        Broad
    }

    public class OfficeVersionInfo
    {
        public string Name { get; set; }

        public string Description { get; set; }

        public OfficeEdition Edition { get; set; }

        public OfficeArchitecture Architecture { get; set; }

        public UpdateChannel Channel { get; set; }

        public string ScriptFileName { get; set; }

        public OfficeVersionInfo()
        {
            Name = string.Empty;
            Description = string.Empty;
            ScriptFileName = string.Empty;
        }

        public OfficeVersionInfo(OfficeEdition edition, OfficeArchitecture architecture, UpdateChannel channel)
        {
            Edition = edition;
            Architecture = architecture;
            Channel = channel;

            string editionName = edition == OfficeEdition.Business ? "Business" : "Enterprise";
            string archName = architecture == OfficeArchitecture.x64 ? "64" : "32";
            string channelName = channel == UpdateChannel.Monthly ? "Current" : "Broad";

            ScriptFileName = $"{editionName}{channelName}{archName}.ps1";

            string editionDisplay = edition == OfficeEdition.Business ? "商业版" : "企业版";
            string archDisplay = architecture == OfficeArchitecture.x64 ? "64位" : "32位";
            string channelDisplay = channel == UpdateChannel.Monthly ? "当前频道" : "半年频道";

            Name = $"{editionDisplay} {archDisplay} {channelDisplay}";
            Description = $"Microsoft 365 Apps {editionDisplay} ({archDisplay}) - {channelDisplay}";
        }
    }
}
