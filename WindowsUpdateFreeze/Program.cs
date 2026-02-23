using System;
using System.Diagnostics;
using System.Security.Principal;
using System.Windows.Forms;
using static System.Windows.Forms.DataFormats;

namespace WinUtilityAppliance
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            if (!IsAdministrator())
            {
                TryRelaunchAsAdmin();
                return;
            }

            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());
        }

        private static bool IsAdministrator()
        {
            using var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        private static void TryRelaunchAsAdmin()
        {
            try
            {
                var exe = Process.GetCurrentProcess().MainModule?.FileName;
                if (string.IsNullOrWhiteSpace(exe))
                {
                    MessageBox.Show("Unable to determine executable path for elevation.", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var psi = new ProcessStartInfo(exe)
                {
                    UseShellExecute = true,
                    Verb = "runas"
                };

                Process.Start(psi);
            }
            catch
            {
                MessageBox.Show("Admin elevation was cancelled or failed. This app must be run as Administrator.",
                    "Elevation Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}