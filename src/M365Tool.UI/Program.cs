using System;
using System.Security.Principal;
using System.Windows.Forms;

namespace Office365CleanupTool
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            if (!IsRunningAsAdministrator())
            {
                MessageBox.Show(
                    "请使用“以管理员身份运行”启动 21V M365助手。\r\n\r\nThis tool must be run as administrator.",
                    "21V M365助手",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            Application.Run(new WorkbenchForm());
        }

        private static bool IsRunningAsAdministrator()
        {
            try
            {
                using WindowsIdentity identity = WindowsIdentity.GetCurrent();
                WindowsPrincipal principal = new(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch
            {
                return false;
            }
        }
    }
}
