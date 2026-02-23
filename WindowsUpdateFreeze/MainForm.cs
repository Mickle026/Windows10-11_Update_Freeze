using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Windows.Forms;

namespace WinUtilityAppliance
{
    public sealed class MainForm : Form
    {
        private readonly Button _btnFreeze = new() { Text = "Freeze (Disable Updates + Tasks)", Width = 260, Height = 40 };
        private readonly Button _btnUnfreeze = new() { Text = "Unfreeze (Restore Snapshot)", Width = 260, Height = 40 };
        private readonly Button _btnAlwaysOn = new() { Text = "Power: Always On", Width = 260, Height = 40 };
        private readonly Button _btnRestorePower = new() { Text = "Power: Restore", Width = 260, Height = 40 };
        private readonly Button _btnRefresh = new() { Text = "Refresh Status", Width = 260, Height = 34 };
        private readonly TextBox _log = new() { Multiline = true, ScrollBars = ScrollBars.Vertical, ReadOnly = true, WordWrap = false };
        private readonly Label _status = new() { AutoSize = true };

        private const string AppName = "WinUtilityAppliance";
        private static readonly string SnapshotPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
            AppName,
            "snapshot.txt"
        );

        public MainForm()
        {
            Text = "Windows Utility Appliance (Freeze / Unfreeze)";
            Width = 900;
            Height = 620;
            StartPosition = FormStartPosition.CenterScreen;

            var left = new FlowLayoutPanel
            {
                Dock = DockStyle.Left,
                Width = 290,
                FlowDirection = FlowDirection.TopDown,
                Padding = new Padding(12),
                WrapContents = false,
                AutoScroll = true
            };

            left.Controls.Add(_btnFreeze);
            left.Controls.Add(_btnUnfreeze);
            left.Controls.Add(new Label { Height = 10 });
            left.Controls.Add(_btnAlwaysOn);
            left.Controls.Add(_btnRestorePower);
            left.Controls.Add(new Label { Height = 10 });
            left.Controls.Add(_btnRefresh);
            left.Controls.Add(new Label { Height = 10 });
            left.Controls.Add(_status);

            _log.Dock = DockStyle.Fill;
            _log.Font = new System.Drawing.Font("Consolas", 10);

            Controls.Add(_log);
            Controls.Add(left);

            _btnFreeze.Click += (_, _) => RunSafe("FREEZE", Freeze);
            _btnUnfreeze.Click += (_, _) => RunSafe("UNFREEZE", Unfreeze);
            _btnAlwaysOn.Click += (_, _) => RunSafe("POWER_ALWAYS_ON", ApplyAlwaysOnPower);
            _btnRestorePower.Click += (_, _) => RunSafe("POWER_RESTORE", RestorePower);
            _btnRefresh.Click += (_, _) => RunSafe("REFRESH", RefreshStatus);

            Directory.CreateDirectory(Path.GetDirectoryName(SnapshotPath)!);
            RefreshStatus();
        }

        // -----------------------------
        // Top-level actions
        // -----------------------------
        private void Freeze()
        {
            if (File.Exists(SnapshotPath))
            {
                Log($"Snapshot already exists: {SnapshotPath}");
                Log("Freeze will overwrite it with a new snapshot.");
            }

            var snap = new Snapshot();
            snap.Capture();

            SaveSnapshot(snap);

            Log("Snapshot saved. Applying freeze changes...");

            // Services: disable update engines
            SetServiceState("wuauserv", ServiceStartMode.Disabled, stopIfRunning: true);
            SetServiceState("usosvc", ServiceStartMode.Disabled, stopIfRunning: true);

            // Policies: set Windows Update policies (reversible)
            ApplyWindowsUpdatePolicies();

            // Tasks: disable common update tasks / reboot triggers (reversible via snapshot)
            DisableTasksByPrefix(@"\Microsoft\Windows\UpdateOrchestrator\");
            DisableTasksByPrefix(@"\Microsoft\Windows\WindowsUpdate\");
            DisableTasksByPrefix(@"\Microsoft\Windows\InstallService\");

            // Common noisy updaters
            DisableTaskIfExists(@"\MicrosoftEdgeUpdateTaskMachineCore");
            DisableTaskIfExists(@"\MicrosoftEdgeUpdateTaskMachineUA");
            DisableTasksByWildcard(@"\OneDrive Standalone Update Task*");
            DisableTasksByWildcard(@"\Mozilla\Firefox Background Update*");

            // Optional: UNP update nags
            DisableTaskIfExists(@"\Microsoft\Windows\UNP\RunUpdateNotificationMgr");

            Log("Freeze complete.");
            RefreshStatus();
        }

        private void Unfreeze()
        {
            var snap = LoadSnapshot();
            if (snap is null)
            {
                Log("No snapshot found. Nothing to restore.");
                return;
            }

            Log("Restoring from snapshot...");

            // Restore services
            foreach (var kvp in snap.Services)
            {
                var name = kvp.Key;
                var mode = kvp.Value;
                SetServiceState(name, mode, stopIfRunning: false);

                // If it was running before and not disabled now, try start
                if (snap.ServicesWasRunning.TryGetValue(name, out var wasRunning) && wasRunning && mode != ServiceStartMode.Disabled)
                {
                    TryStartService(name);
                }
            }

            // Restore registry policy keys we touched
            RestoreWindowsUpdatePolicies(snap);

            // Restore scheduled tasks enabled/disabled state
            foreach (var task in snap.Tasks)
            {
                SetTaskEnabled(task.Key, task.Value);
            }

            // Restore power settings
            if (snap.Power is not null)
            {
                RestorePowerFromSnapshot(snap.Power);
            }

            Log("Unfreeze complete.");

            // Keep snapshot (so you can re-restore), but you can delete if you want:
            // File.Delete(SnapshotPath);

            RefreshStatus();
        }

        private void ApplyAlwaysOnPower()
        {
            var snap = LoadSnapshot() ?? new Snapshot();
            if (snap.Power is null)
            {
                snap.Power = PowerSnapshot.Capture();
                SaveSnapshot(snap);
                Log("Power snapshot captured & saved.");
            }

            Log("Applying Always-On power settings...");
            Run("powercfg", "-h off");
            Run("powercfg", "/change standby-timeout-ac 0");
            Run("powercfg", "/change monitor-timeout-ac 0");
            Run("powercfg", "/change hibernate-timeout-ac 0");
            Run("powercfg", "/change standby-timeout-dc 0");
            Run("powercfg", "/change monitor-timeout-dc 0");
            Run("powercfg", "/change hibernate-timeout-dc 0");

            Log("Power: Always-On applied.");
        }

        private void RestorePower()
        {
            var snap = LoadSnapshot();
            if (snap?.Power is null)
            {
                Log("No power snapshot saved yet. Click 'Power: Always On' once to capture it.");
                return;
            }

            RestorePowerFromSnapshot(snap.Power);
            Log("Power restored from snapshot.");
        }

        private void RefreshStatus()
        {
            var sb = new StringBuilder();
            sb.AppendLine("=== Service Status ===");
            sb.AppendLine(ServiceLine("wuauserv"));
            sb.AppendLine(ServiceLine("usosvc"));
            sb.AppendLine(ServiceLine("WaaSMedicSvc"));
            sb.AppendLine();
            sb.AppendLine("=== Snapshot ===");
            sb.AppendLine(File.Exists(SnapshotPath) ? $"Snapshot: {SnapshotPath}" : "Snapshot: (none)");
            _status.Text = File.Exists(SnapshotPath) ? "Snapshot: Present" : "Snapshot: Missing";

            Log(sb.ToString().TrimEnd());
        }

        // -----------------------------
        // Registry policies (Windows Update)
        // -----------------------------
        private static readonly string WUPath = @"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate";
        private static readonly string AUPath = @"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU";

        private void ApplyWindowsUpdatePolicies()
        {
            // These mirror the “quiet/appliance” posture. Services being disabled is the main kill-switch.
            SetRegDword(Registry.LocalMachine, AUPath, "NoAutoUpdate", 1);
            SetRegDword(Registry.LocalMachine, AUPath, "AUOptions", 2);
            SetRegDword(Registry.LocalMachine, AUPath, "NoAutoRebootWithLoggedOnUsers", 1);
            SetRegDword(Registry.LocalMachine, AUPath, "AlwaysAutoRebootAtScheduledTime", 0);
            SetRegDword(Registry.LocalMachine, AUPath, "AlwaysAutoRebootAtScheduledTimeMinutes", 0);
            Log("Windows Update policies applied.");
        }

        private void RestoreWindowsUpdatePolicies(Snapshot snap)
        {
            foreach (var kvp in snap.RegistryDwords)
            {
                var keyPath = kvp.Key.keyPath;
                var name = kvp.Key.name;
                var val = kvp.Value;

                if (val is null)
                {
                    DeleteRegValueIfExists(Registry.LocalMachine, keyPath, name);
                }
                else
                {
                    SetRegDword(Registry.LocalMachine, keyPath, name, val.Value);
                }
            }

            Log("Windows Update policies restored.");
        }

        private void SetRegDword(RegistryKey root, string subKey, string name, int value)
        {
            using var k = root.CreateSubKey(subKey, true);
            k.SetValue(name, value, RegistryValueKind.DWord);
        }

        private void DeleteRegValueIfExists(RegistryKey root, string subKey, string name)
        {
            using var k = root.OpenSubKey(subKey, true);
            if (k is null) return;
            if (k.GetValue(name) is null) return;
            k.DeleteValue(name, false);
        }

        // -----------------------------
        // Services & Tasks
        // -----------------------------
        private void SetServiceState(string name, ServiceStartMode mode, bool stopIfRunning)
        {
            // Use sc.exe because Set-Service StartType has edge cases with some system services.
            // Also avoids PowerShell alias issues entirely.
            if (stopIfRunning)
            {
                TryStopService(name);
            }

            Run("sc.exe", $"config {name} start= {ToScStart(mode)}");
            Log($"Service '{name}' set to {mode}.");
        }

        private void TryStopService(string name)
        {
            try
            {
                using var sc = new ServiceController(name);
                if (sc.Status == ServiceControllerStatus.Running ||
                    sc.Status == ServiceControllerStatus.StartPending)
                {
                    Run("sc.exe", $"stop {name}");
                }
            }
            catch
            {
                // ignore
            }
        }

        private void TryStartService(string name)
        {
            try
            {
                using var sc = new ServiceController(name);
                if (sc.Status == ServiceControllerStatus.Stopped ||
                    sc.Status == ServiceControllerStatus.StopPending)
                {
                    Run("sc.exe", $"start {name}");
                }
            }
            catch
            {
                // ignore
            }
        }

        private static string ToScStart(ServiceStartMode mode) => mode switch
        {
            ServiceStartMode.Disabled => "disabled",
            ServiceStartMode.Auto => "auto",
            ServiceStartMode.Manual => "demand",
            _ => "demand"
        };

        private void DisableTasksByPrefix(string folderPrefix)
        {
            var tasks = ListAllTasks();
            var matched = tasks.Where(t => t.StartsWith(folderPrefix, StringComparison.OrdinalIgnoreCase)).ToList();
            foreach (var t in matched)
            {
                DisableTaskIfExists(t);
            }
            Log($"Disabled {matched.Count} tasks under {folderPrefix}");
        }

        private void DisableTasksByWildcard(string wildcard)
        {
            var tasks = ListAllTasks();
            var matched = tasks.Where(t => WildcardMatch(t, wildcard)).ToList();
            foreach (var t in matched)
            {
                DisableTaskIfExists(t);
            }
            Log($"Disabled {matched.Count} tasks matching {wildcard}");
        }

        private void DisableTaskIfExists(string taskName)
        {
            if (!TaskExists(taskName)) return;
            Run("schtasks", $"/Change /TN \"{taskName}\" /Disable");
        }

        private void SetTaskEnabled(string taskName, bool enabled)
        {
            if (!TaskExists(taskName)) return;
            var flag = enabled ? "/Enable" : "/Disable";
            Run("schtasks", $"/Change /TN \"{taskName}\" {flag}");
        }

        private bool TaskExists(string taskName)
        {
            var r = Run("schtasks", $"/Query /TN \"{taskName}\"", throwOnFail: false);
            return r.ExitCode == 0;
        }

        private List<string> ListAllTasks()
        {
            var r = Run("schtasks", "/Query /FO LIST", throwOnFail: false);
            var lines = r.StdOut.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            var tasks = new List<string>();
            foreach (var line in lines)
            {
                if (line.StartsWith("TaskName:", StringComparison.OrdinalIgnoreCase))
                {
                    var tn = line.Substring("TaskName:".Length).Trim();
                    if (!string.IsNullOrWhiteSpace(tn))
                        tasks.Add(tn);
                }
            }
            return tasks;
        }

        private static bool WildcardMatch(string input, string pattern)
        {
            // Very small wildcard: '*' only
            // pattern like "\OneDrive Standalone Update Task*"
            var parts = pattern.Split('*');
            if (parts.Length == 1) return input.Equals(pattern, StringComparison.OrdinalIgnoreCase);

            var idx = 0;
            foreach (var part in parts)
            {
                if (string.IsNullOrEmpty(part)) continue;
                var found = input.IndexOf(part, idx, StringComparison.OrdinalIgnoreCase);
                if (found < 0) return false;
                idx = found + part.Length;
            }
            return true;
        }

        // -----------------------------
        // Process runner + UI logging
        // -----------------------------
        private (int ExitCode, string StdOut, string StdErr) Run(string file, string args, bool throwOnFail = true)
        {
            Log($"> {file} {args}");

            var psi = new ProcessStartInfo(file, args)
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            using var p = Process.Start(psi)!;
            var stdout = p.StandardOutput.ReadToEnd();
            var stderr = p.StandardError.ReadToEnd();
            p.WaitForExit();

            if (!string.IsNullOrWhiteSpace(stdout)) Log(stdout.TrimEnd());
            if (!string.IsNullOrWhiteSpace(stderr)) Log(stderr.TrimEnd());

            if (throwOnFail && p.ExitCode != 0)
                throw new InvalidOperationException($"{file} failed with exit code {p.ExitCode}");

            return (p.ExitCode, stdout, stderr);
        }

        private void Log(string text)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => Log(text)));
                return;
            }

            _log.AppendText($"[{DateTime.Now:HH:mm:ss}] {text}\r\n");
            _log.SelectionStart = _log.TextLength;
            _log.ScrollToCaret();
        }

        private void RunSafe(string label, Action action)
        {
            try
            {
                Log($"--- {label} ---");
                action();
            }
            catch (Exception ex)
            {
                Log($"ERROR: {ex.Message}");
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string ServiceLine(string name)
        {
            try
            {
                using var sc = new ServiceController(name);
                var status = sc.Status.ToString();
                var startMode = QueryStartMode(name);
                return $"{name,-12}  Status={status,-12}  StartType={startMode}";
            }
            catch
            {
                return $"{name,-12}  (not found / access denied)";
            }
        }

        private static string QueryStartMode(string name)
        {
            // Use sc qc to avoid WMI permission quirks
            try
            {
                var psi = new ProcessStartInfo("sc.exe", $"qc {name}")
                {
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };
                using var p = Process.Start(psi)!;
                var outp = p.StandardOutput.ReadToEnd();
                p.WaitForExit();

                // Look for: START_TYPE         : 4   DISABLED
                var line = outp.Split('\n').FirstOrDefault(l => l.Contains("START_TYPE", StringComparison.OrdinalIgnoreCase));
                if (line is null) return "?";
                if (line.Contains("DISABLED", StringComparison.OrdinalIgnoreCase)) return "Disabled";
                if (line.Contains("AUTO_START", StringComparison.OrdinalIgnoreCase)) return "Auto";
                if (line.Contains("DEMAND_START", StringComparison.OrdinalIgnoreCase)) return "Manual";
                return line.Trim();
            }
            catch
            {
                return "?";
            }
        }

        // -----------------------------
        // Snapshot persistence (simple text format)
        // -----------------------------
        private void SaveSnapshot(Snapshot snap)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(SnapshotPath)!);
            File.WriteAllText(SnapshotPath, snap.Serialize(), Encoding.UTF8);
        }

        private Snapshot? LoadSnapshot()
        {
            if (!File.Exists(SnapshotPath)) return null;
            var txt = File.ReadAllText(SnapshotPath, Encoding.UTF8);
            return Snapshot.Deserialize(txt);
        }

        private void RestorePowerFromSnapshot(PowerSnapshot ps)
        {
            // Restore hibernate toggle + timeouts
            if (ps.HibernateEnabled)
            {
                Run("powercfg", "-h on", throwOnFail: false);
            }
            else
            {
                Run("powercfg", "-h off", throwOnFail: false);
            }

            Run("powercfg", $"/change standby-timeout-ac {ps.StandbyAc}", throwOnFail: false);
            Run("powercfg", $"/change monitor-timeout-ac {ps.MonitorAc}", throwOnFail: false);
            Run("powercfg", $"/change hibernate-timeout-ac {ps.HibernateAc}", throwOnFail: false);

            Run("powercfg", $"/change standby-timeout-dc {ps.StandbyDc}", throwOnFail: false);
            Run("powercfg", $"/change monitor-timeout-dc {ps.MonitorDc}", throwOnFail: false);
            Run("powercfg", $"/change hibernate-timeout-dc {ps.HibernateDc}", throwOnFail: false);
        }
    }

    internal enum ServiceStartMode
    {
        Auto,
        Manual,
        Disabled
    }

    internal sealed class PowerSnapshot
    {
        public bool HibernateEnabled { get; set; }
        public int StandbyAc { get; set; }
        public int MonitorAc { get; set; }
        public int HibernateAc { get; set; }
        public int StandbyDc { get; set; }
        public int MonitorDc { get; set; }
        public int HibernateDc { get; set; }

        public static PowerSnapshot Capture()
        {
            // Use powercfg /query? It’s verbose. For appliance use, capture the visible /change values:
            // We’ll read them via `powercfg /getactivescheme` + `powercfg /query` is heavy.
            // Instead: best-effort capture by parsing `powercfg /q` minimal signals is complex.
            // Practical approach: store what *we last set*? But you asked to reverse itself.
            // So we capture using `powercfg /q` and parse the relevant AC/DC timeouts.

            // NOTE: This capture is a best-effort parser; it works on standard Win11 outputs.
            var ps = new PowerSnapshot();

            ps.HibernateEnabled = File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "hiberfil.sys"));

            // Defaults if parse fails
            ps.StandbyAc = 30;
            ps.MonitorAc = 10;
            ps.HibernateAc = 0;
            ps.StandbyDc = 15;
            ps.MonitorDc = 5;
            ps.HibernateDc = 0;

            try
            {
                var active = RunCapture("powercfg", "/getactivescheme");
                var guid = ParseFirstGuid(active);
                if (guid is null) return ps;

                var q = RunCapture("powercfg", $"/q {guid}");
                // Parse likely lines:
                // "Current AC Power Setting Index: 0x0000000a" etc.
                // We search for known setting GUIDs:
                // Sleep idle timeout: 29f6c1db-86da-48c5-9fdb-f2b67b1f44da (SUB_SLEEP)
                // Monitor timeout: 3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e (VIDEOIDLE)
                // Hibernate after: 9d7815a6-7ee4-497e-8888-515a05f02364 (HIBERNATEIDLE)
                ps.StandbyAc = ParseSettingSeconds(q, "29f6c1db-86da-48c5-9fdb-f2b67b1f44da", ac: true) / 60;
                ps.StandbyDc = ParseSettingSeconds(q, "29f6c1db-86da-48c5-9fdb-f2b67b1f44da", ac: false) / 60;

                ps.MonitorAc = ParseSettingSeconds(q, "3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e", ac: true) / 60;
                ps.MonitorDc = ParseSettingSeconds(q, "3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e", ac: false) / 60;

                ps.HibernateAc = ParseSettingSeconds(q, "9d7815a6-7ee4-497e-8888-515a05f02364", ac: true) / 60;
                ps.HibernateDc = ParseSettingSeconds(q, "9d7815a6-7ee4-497e-8888-515a05f02364", ac: false) / 60;

                // If parser returns 0 due to fail, keep defaults (best effort)
            }
            catch
            {
                // best effort only
            }

            return ps;
        }

        private static string RunCapture(string file, string args)
        {
            var psi = new ProcessStartInfo(file, args)
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };
            using var p = Process.Start(psi)!;
            var stdout = p.StandardOutput.ReadToEnd();
            p.WaitForExit();
            return stdout;
        }

        private static string? ParseFirstGuid(string text)
        {
            // finds first GUID in output
            var start = text.IndexOf('{');
            var end = text.IndexOf('}');
            if (start >= 0 && end > start)
            {
                return text.Substring(start + 1, end - start - 1).Trim();
            }
            return null;
        }

        private static int ParseSettingSeconds(string powercfgQ, string settingGuid, bool ac)
        {
            // Find section containing setting GUID; then read "Current AC Power Setting Index"
            var idx = powercfgQ.IndexOf(settingGuid, StringComparison.OrdinalIgnoreCase);
            if (idx < 0) return 0;

            var chunk = powercfgQ.Substring(idx);
            var lineKey = ac ? "Current AC Power Setting Index" : "Current DC Power Setting Index";

            // Take next ~2000 chars, find line
            var limit = Math.Min(chunk.Length, 2500);
            chunk = chunk.Substring(0, limit);

            var lines = chunk.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            var line = lines.FirstOrDefault(l => l.IndexOf(lineKey, StringComparison.OrdinalIgnoreCase) >= 0);
            if (line is null) return 0;

            // line like "...: 0x0000003c"
            var hexIdx = line.IndexOf("0x", StringComparison.OrdinalIgnoreCase);
            if (hexIdx < 0) return 0;

            var hex = new string(line.Substring(hexIdx).TakeWhile(c => char.IsLetterOrDigit(c) || c == 'x').ToArray());
            if (!hex.StartsWith("0x", StringComparison.OrdinalIgnoreCase)) return 0;

            if (int.TryParse(hex.Substring(2), System.Globalization.NumberStyles.HexNumber, null, out var val))
            {
                // For timeouts, powercfg stores seconds.
                return val;
            }
            return 0;
        }
    }

    internal sealed class Snapshot
    {
        public Dictionary<string, ServiceStartMode> Services { get; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, bool> ServicesWasRunning { get; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, bool> Tasks { get; } = new(StringComparer.OrdinalIgnoreCase);

        // store the exact DWORD values we touch; null means “didn’t exist”
        public Dictionary<(string keyPath, string name), int?> RegistryDwords { get; } = new();

        public PowerSnapshot? Power { get; set; }

        public void Capture()
        {
            CaptureService("wuauserv");
            CaptureService("usosvc");

            CaptureWindowsUpdatePolicyDword(@"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU", "NoAutoUpdate");
            CaptureWindowsUpdatePolicyDword(@"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU", "AUOptions");
            CaptureWindowsUpdatePolicyDword(@"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU", "NoAutoRebootWithLoggedOnUsers");
            CaptureWindowsUpdatePolicyDword(@"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU", "AlwaysAutoRebootAtScheduledTime");
            CaptureWindowsUpdatePolicyDword(@"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU", "AlwaysAutoRebootAtScheduledTimeMinutes");

            CaptureTasksWithPrefix(@"\Microsoft\Windows\UpdateOrchestrator\");
            CaptureTasksWithPrefix(@"\Microsoft\Windows\WindowsUpdate\");
            CaptureTasksWithPrefix(@"\Microsoft\Windows\InstallService\");

            CaptureTask(@"\MicrosoftEdgeUpdateTaskMachineCore");
            CaptureTask(@"\MicrosoftEdgeUpdateTaskMachineUA");
            CaptureTasksByWildcard(@"\OneDrive Standalone Update Task*");
            CaptureTasksByWildcard(@"\Mozilla\Firefox Background Update*");
            CaptureTask(@"\Microsoft\Windows\UNP\RunUpdateNotificationMgr");
        }

        private void CaptureService(string name)
        {
            try
            {
                using var sc = new ServiceController(name);
                ServicesWasRunning[name] = sc.Status == ServiceControllerStatus.Running;

                // query start mode from sc qc output
                var mode = QueryStartMode(name);
                Services[name] = mode;
            }
            catch
            {
                // ignore
            }
        }

        private static ServiceStartMode QueryStartMode(string name)
        {
            var psi = new ProcessStartInfo("sc.exe", $"qc {name}")
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };
            using var p = Process.Start(psi)!;
            var outp = p.StandardOutput.ReadToEnd();
            p.WaitForExit();

            var line = outp.Split('\n').FirstOrDefault(l => l.Contains("START_TYPE", StringComparison.OrdinalIgnoreCase));
            if (line is null) return ServiceStartMode.Manual;

            if (line.Contains("DISABLED", StringComparison.OrdinalIgnoreCase)) return ServiceStartMode.Disabled;
            if (line.Contains("AUTO_START", StringComparison.OrdinalIgnoreCase)) return ServiceStartMode.Auto;
            if (line.Contains("DEMAND_START", StringComparison.OrdinalIgnoreCase)) return ServiceStartMode.Manual;

            return ServiceStartMode.Manual;
        }

        private void CaptureWindowsUpdatePolicyDword(string keyPath, string name)
        {
            int? existing = null;
            using var k = Registry.LocalMachine.OpenSubKey(keyPath, false);
            if (k != null)
            {
                var v = k.GetValue(name);
                if (v is int i) existing = i;
            }
            RegistryDwords[(keyPath, name)] = existing;
        }

        private void CaptureTasksWithPrefix(string prefix)
        {
            var tasks = ListAllTasks();
            foreach (var t in tasks.Where(t => t.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)))
            {
                CaptureTask(t);
            }
        }

        private void CaptureTasksByWildcard(string wildcard)
        {
            var tasks = ListAllTasks();
            foreach (var t in tasks.Where(t => WildcardMatch(t, wildcard)))
            {
                CaptureTask(t);
            }
        }

        private void CaptureTask(string taskName)
        {
            var exists = TaskExists(taskName);
            if (!exists) return;
            var enabled = IsTaskEnabled(taskName);
            Tasks[taskName] = enabled;
        }

        private static bool IsTaskEnabled(string taskName)
        {
            var psi = new ProcessStartInfo("schtasks", $"/Query /TN \"{taskName}\" /FO LIST /V")
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };
            using var p = Process.Start(psi)!;
            var stdout = p.StandardOutput.ReadToEnd();
            p.WaitForExit();

            // "Scheduled Task State: Enabled" (varies), or "Enabled: Yes" sometimes.
            // We'll look for "Enabled" tokens.
            if (stdout.IndexOf("Disabled", StringComparison.OrdinalIgnoreCase) >= 0) return false;
            if (stdout.IndexOf("Enabled", StringComparison.OrdinalIgnoreCase) >= 0) return true;

            // fallback: assume enabled
            return true;
        }

        private static bool TaskExists(string taskName)
        {
            var psi = new ProcessStartInfo("schtasks", $"/Query /TN \"{taskName}\"")
            {
                UseShellExecute = false,
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                CreateNoWindow = true
            };
            using var p = Process.Start(psi)!;
            p.WaitForExit();
            return p.ExitCode == 0;
        }

        private static List<string> ListAllTasks()
        {
            var psi = new ProcessStartInfo("schtasks", "/Query /FO LIST")
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };
            using var p = Process.Start(psi)!;
            var stdout = p.StandardOutput.ReadToEnd();
            p.WaitForExit();

            var lines = stdout.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            var tasks = new List<string>();
            foreach (var line in lines)
            {
                if (line.StartsWith("TaskName:", StringComparison.OrdinalIgnoreCase))
                {
                    var tn = line.Substring("TaskName:".Length).Trim();
                    if (!string.IsNullOrWhiteSpace(tn))
                        tasks.Add(tn);
                }
            }
            return tasks;
        }

        private static bool WildcardMatch(string input, string pattern)
        {
            var parts = pattern.Split('*');
            if (parts.Length == 1) return input.Equals(pattern, StringComparison.OrdinalIgnoreCase);

            var idx = 0;
            foreach (var part in parts)
            {
                if (string.IsNullOrEmpty(part)) continue;
                var found = input.IndexOf(part, idx, StringComparison.OrdinalIgnoreCase);
                if (found < 0) return false;
                idx = found + part.Length;
            }
            return true;
        }

        public string Serialize()
        {
            // Simple line format so you can inspect/edit by hand if needed.
            // SVC|name|mode|wasRunning
            // REG|key|name|valueOrNULL
            // TSK|taskName|enabled
            // PWR|HibernateEnabled|StandbyAc|MonitorAc|HibernateAc|StandbyDc|MonitorDc|HibernateDc

            var sb = new StringBuilder();
            foreach (var s in Services)
            {
                var was = ServicesWasRunning.TryGetValue(s.Key, out var wr) && wr;
                sb.AppendLine($"SVC|{s.Key}|{s.Value}|{was}");
            }

            foreach (var r in RegistryDwords)
            {
                var v = r.Value.HasValue ? r.Value.Value.ToString() : "NULL";
                sb.AppendLine($"REG|{r.Key.keyPath}|{r.Key.name}|{v}");
            }

            foreach (var t in Tasks)
            {
                sb.AppendLine($"TSK|{t.Key}|{t.Value}");
            }

            if (Power is not null)
            {
                sb.AppendLine($"PWR|{Power.HibernateEnabled}|{Power.StandbyAc}|{Power.MonitorAc}|{Power.HibernateAc}|{Power.StandbyDc}|{Power.MonitorDc}|{Power.HibernateDc}");
            }

            return sb.ToString();
        }

        public static Snapshot Deserialize(string text)
        {
            var snap = new Snapshot();

            var lines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in lines)
            {
                var parts = line.Split('|');
                if (parts.Length < 2) continue;

                var kind = parts[0];
                if (kind.Equals("SVC", StringComparison.OrdinalIgnoreCase) && parts.Length >= 4)
                {
                    var name = parts[1];
                    if (Enum.TryParse<ServiceStartMode>(parts[2], out var mode))
                        snap.Services[name] = mode;
                    if (bool.TryParse(parts[3], out var was))
                        snap.ServicesWasRunning[name] = was;
                }
                else if (kind.Equals("REG", StringComparison.OrdinalIgnoreCase) && parts.Length >= 4)
                {
                    var key = parts[1];
                    var name = parts[2];
                    var val = parts[3];
                    int? parsed = null;
                    if (!val.Equals("NULL", StringComparison.OrdinalIgnoreCase) && int.TryParse(val, out var i))
                        parsed = i;
                    snap.RegistryDwords[(key, name)] = parsed;
                }
                else if (kind.Equals("TSK", StringComparison.OrdinalIgnoreCase) && parts.Length >= 3)
                {
                    var task = parts[1];
                    if (bool.TryParse(parts[2], out var en))
                        snap.Tasks[task] = en;
                }
                else if (kind.Equals("PWR", StringComparison.OrdinalIgnoreCase) && parts.Length >= 8)
                {
                    snap.Power = new PowerSnapshot
                    {
                        HibernateEnabled = bool.TryParse(parts[1], out var hb) && hb,
                        StandbyAc = int.TryParse(parts[2], out var sa) ? sa : 0,
                        MonitorAc = int.TryParse(parts[3], out var ma) ? ma : 0,
                        HibernateAc = int.TryParse(parts[4], out var ha) ? ha : 0,
                        StandbyDc = int.TryParse(parts[5], out var sd) ? sd : 0,
                        MonitorDc = int.TryParse(parts[6], out var md) ? md : 0,
                        HibernateDc = int.TryParse(parts[7], out var hd) ? hd : 0
                    };
                }
            }

            return snap;
        }
    }
}
