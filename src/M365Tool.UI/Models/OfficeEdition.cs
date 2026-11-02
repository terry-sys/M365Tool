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
        Monthly = 0, // Legacy enum name for Current Channel.
        Broad = 1,   // Preserve the existing SemiAnnual value.
        MonthlyEnterprise = 2
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
            string channelName = channel switch
            {
                UpdateChannel.Monthly => "Current",
                UpdateChannel.Broad => "Broad",
                UpdateChannel.MonthlyEnterprise => "MonthlyEnterprise",
                _ => throw new System.ArgumentOutOfRangeException(nameof(channel))
            };

            ScriptFileName = $"{editionName}{channelName}{archName}.ps1";

            string editionDisplay = edition == OfficeEdition.Business ? "商业版" : "企业版";
            string archDisplay = architecture == OfficeArchitecture.x64 ? "64位" : "32位";
            string channelDisplay = Services.OfficeChannelService.GetChannelDisplayName(channel);

            Name = $"{editionDisplay} {archDisplay} {channelDisplay}";
            Description = $"Microsoft 365 Apps {editionDisplay} ({archDisplay}) - {channelDisplay}";
        }
    }
}
