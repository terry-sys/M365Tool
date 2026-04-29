using System.Windows.Forms;
using Office365CleanupTool.Services;

namespace Office365CleanupTool
{
    internal static class WorkbenchTextLocalizer
    {
        public static void Apply(Control root, UiLanguage language)
        {
            // Global text replacement has been retired.
            // Each page now owns its own localization logic.
        }

        public static string Translate(string? text, UiLanguage language)
        {
            return text ?? string.Empty;
        }
    }
}
