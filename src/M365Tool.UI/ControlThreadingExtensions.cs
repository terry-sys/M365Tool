using System;
using System.Windows.Forms;

namespace Office365CleanupTool
{
    internal static class ControlThreadingExtensions
    {
        public static void BeginInvokeWhenReady(this Control control, Action action)
        {
            if (control == null || action == null || control.IsDisposed || control.Disposing)
            {
                return;
            }

            if (control.IsHandleCreated)
            {
                try
                {
                    control.BeginInvoke(action);
                }
                catch (InvalidOperationException)
                {
                    // Handle was destroyed between checks; ignore.
                }

                return;
            }

            void OnHandleCreated(object? sender, EventArgs args)
            {
                control.HandleCreated -= OnHandleCreated;
                if (control.IsDisposed || control.Disposing || !control.IsHandleCreated)
                {
                    return;
                }

                try
                {
                    control.BeginInvoke(action);
                }
                catch (InvalidOperationException)
                {
                    // Handle was destroyed between checks; ignore.
                }
            }

            control.HandleCreated += OnHandleCreated;
        }

        public static bool TryBeginInvokeIfReady(this Control control, Action action)
        {
            if (control == null || action == null || control.IsDisposed || control.Disposing || !control.IsHandleCreated)
            {
                return false;
            }

            try
            {
                control.BeginInvoke(action);
                return true;
            }
            catch (InvalidOperationException)
            {
                return false;
            }
        }
    }
}
