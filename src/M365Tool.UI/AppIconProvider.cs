using System;
using System.Drawing;
using System.Windows.Forms;

namespace Office365CleanupTool
{
    internal static class AppIconProvider
    {
        private static Icon? _cachedIcon;

        public static Icon? GetAppIcon()
        {
            if (_cachedIcon != null)
            {
                return _cachedIcon;
            }

            try
            {
                string executablePath = Application.ExecutablePath;
                Icon? extracted = Icon.ExtractAssociatedIcon(executablePath);
                if (extracted != null)
                {
                    _cachedIcon = extracted;
                    return _cachedIcon;
                }
            }
            catch
            {
                // Ignore extraction failures and let caller fall back to default icon.
            }

            return null;
        }
    }
}

