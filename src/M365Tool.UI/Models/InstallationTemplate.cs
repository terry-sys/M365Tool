using System;

namespace Office365CleanupTool.Models
{
    public class InstallationTemplate
    {
        public string Name { get; set; } = string.Empty;

        public string Remark { get; set; } = string.Empty;

        public DateTime CreatedAt { get; set; } = DateTime.Now;

        public InstallationConfig Config { get; set; } = new InstallationConfig();
    }
}
