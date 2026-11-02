using System;
using System.Drawing;

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
                using var stream = typeof(AppIconProvider).Assembly
                    .GetManifestResourceStream("M365Tool.Assets.AppIcon.ico");
                if (stream != null)
                {
                    using var icon = new Icon(stream);
                    _cachedIcon = (Icon)icon.Clone();
                    return _cachedIcon;
                }
            }
            catch
            {
                // Ignore resource loading failures and let caller fall back to default icon.
            }

            return null;
        }
    }
}

