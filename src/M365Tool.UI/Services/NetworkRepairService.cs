using System;
using System.Threading.Tasks;
using System.Threading;
using System.IO;
using System.Text.Json;
using System.Security.Principal;
using System.Linq;
using Microsoft.Win32;

namespace Office365CleanupTool.Services
{
    public enum NetworkRepairStatus
    {
        Processing,
        Success,
        Warning
    }

    public class NetworkRepairProgressEventArgs : EventArgs
    {
        public NetworkRepairProgressEventArgs(string message, NetworkRepairStatus status)
        {
            Message = message;
            Status = status;
        }

        public string Message { get; }

        public NetworkRepairStatus Status { get; }
    }

    public class NetworkRepairService
    {
        private readonly IScriptRunner _scriptRunner;
        private readonly string _internetSettings;
        private readonly string? _backupDirectory;

        public NetworkRepairService(IScriptRunner scriptRunner) : this(scriptRunner, InternetSettings, null) { }

        internal NetworkRepairService(IScriptRunner scriptRunner, string internetSettings, string? backupDirectory)
        {
            _scriptRunner = scriptRunner;
            _internetSettings = internetSettings;
            _backupDirectory = backupDirectory;
        }

        public event EventHandler<NetworkRepairProgressEventArgs>? ProgressChanged;

        private const string InternetSettings = @"Software\Microsoft\Windows\CurrentVersion\Internet Settings";
        private static readonly string NetshPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "netsh.exe");
        private static readonly string[] ValueNames = { "SecureProtocols", "AutoDetect", "AutoConfigURL" };
        public string LastBackupPath { get; private set; } = string.Empty;

        public string GetRepairPreview(NetworkRepairOptions options, UiLanguage language)
        {
            using var key = Registry.CurrentUser.OpenSubKey(_internetSettings);
            int protocols = key?.GetValue("SecureProtocols") is int number ? number : 0;
            string pac = key?.GetValue("AutoConfigURL", "", RegistryValueOptions.DoNotExpandEnvironmentNames)?.ToString() ?? "";
            var changes = new System.Collections.Generic.List<string>();
            if (options.EnableTls12) changes.Add(T(language, $"TLS 1.2：{protocols} → {protocols | 0x800}", $"TLS 1.2: {protocols} -> {protocols | 0x800}"));
            if (options.EnableAutoDetect) changes.Add(T(language, "启用自动代理检测。", "Enable automatic proxy detection."));
            if (options.ImportUserProxy) changes.Add(T(language, "用当前用户代理替换 WinHTTP 代理。", "Replace WinHTTP proxy with current-user proxy."));
            return string.Join(Environment.NewLine, changes) + Environment.NewLine +
                T(language, $"PAC 地址保留：{pac}。执行前会备份原配置。", $"PAC address will be preserved: {pac}. Original settings will be backed up before changes.");
        }

        public async Task RunAsync(UiLanguage language = UiLanguage.Chinese,
            NetworkRepairOptions? options = null, CancellationToken cancellationToken = default)
        {
            using var operation = MaintenanceCoordinator.Shared.Begin(language);
            options ??= new NetworkRepairOptions();
            if (!options.EnableTls12 && !options.EnableAutoDetect && !options.ImportUserProxy)
                throw new InvalidOperationException(T(language, "请至少选择一项修复操作。", "Select at least one repair action."));

            await CreateBackupAsync(cancellationToken);
            OnProgressChanged(T(language, $"已备份原配置：{LastBackupPath}", $"Original settings backed up: {LastBackupPath}"), NetworkRepairStatus.Processing);
            await Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                using var key = Registry.CurrentUser.OpenSubKey(_internetSettings, true)
                    ?? throw new InvalidOperationException("Internet Settings was not found.");
                if (options.EnableTls12)
                {
                    int current = key.GetValue("SecureProtocols") is int number ? number : 0;
                    key.SetValue("SecureProtocols", current | 0x800, RegistryValueKind.DWord);
                }
                if (options.EnableAutoDetect) key.SetValue("AutoDetect", 1, RegistryValueKind.DWord);
                // A corporate PAC URL must not be cleared as a generic repair step.
            }, cancellationToken);
            if (options.ImportUserProxy)
                await Task.Run(() => ImportProxyFromInternetSettings(language), cancellationToken);
            OnProgressChanged(T(language, "所选网络修复已完成。", "Selected network repairs completed."), NetworkRepairStatus.Success);
        }

        private async Task CreateBackupAsync(CancellationToken cancellationToken)
        {
            var backup = new NetworkRepairBackup { UserSid = CurrentUserSid() };
            using (var key = Registry.CurrentUser.OpenSubKey(_internetSettings))
            {
                foreach (string name in ValueNames)
                {
                    object? value = key?.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
                    backup.Values.Add(new NetworkRegistryValue
                    {
                        Name = name, Exists = value != null,
                        Kind = value == null ? RegistryValueKind.None : key!.GetValueKind(name),
                        Number = value as int?, Text = value as string
                    });
                }
            }
            var proxy = await _scriptRunner.RunAsync(new ScriptExecutionRequest
            {
                FileName = NetshPath, Arguments = "winhttp dump", Timeout = TimeSpan.FromMinutes(1),
                CancellationToken = cancellationToken, KillOnCancellation = true
            });
            cancellationToken.ThrowIfCancellationRequested();
            if (!proxy.IsSuccess) throw new InvalidOperationException("Cannot back up WinHTTP settings: " + proxy.StandardError);
            backup.WinHttpScript = proxy.StandardOutput;
            NetworkRepairBackup.Validate(backup);
            string directory = _backupDirectory ?? ToolPathService.CreateSession("NetworkRepair").LogsDirectory;
            Directory.CreateDirectory(directory);
            string path = Path.Combine(directory, "network-backup-" + Guid.NewGuid().ToString("N") + ".json");
            await File.WriteAllTextAsync(path, JsonSerializer.Serialize(backup, new JsonSerializerOptions { WriteIndented = true }), cancellationToken);
            LastBackupPath = path;
        }

        public async Task RestoreAsync(string path, CancellationToken cancellationToken = default)
        {
            using var operation = MaintenanceCoordinator.Shared.Begin(LocalizationService.ResolveLanguage("Auto"));
            var backup = JsonSerializer.Deserialize<NetworkRepairBackup>(await File.ReadAllTextAsync(path, cancellationToken))
                ?? throw new InvalidDataException("Invalid network backup.");
            NetworkRepairBackup.Validate(backup);
            if (backup.UserSid != CurrentUserSid())
                throw new InvalidOperationException("This network backup belongs to a different Windows user.");
            cancellationToken.ThrowIfCancellationRequested();
            using (var key = Registry.CurrentUser.CreateSubKey(_internetSettings))
            {
                foreach (var value in backup.Values)
                {
                    if (!value.Exists) key.DeleteValue(value.Name, false);
                    else key.SetValue(value.Name, value.Kind == RegistryValueKind.DWord ? (object)value.Number!.Value : value.Text!, value.Kind);
                }
            }
            string scriptPath = Path.Combine(Path.GetDirectoryName(path)!, "restore-winhttp-" + Guid.NewGuid().ToString("N") + ".txt");
            await File.WriteAllTextAsync(scriptPath, backup.WinHttpScript, cancellationToken);
            var result = await _scriptRunner.RunAsync(new ScriptExecutionRequest
            {
                FileName = NetshPath, Arguments = $"exec \"{scriptPath}\"", Timeout = TimeSpan.FromMinutes(1),
                CancellationToken = cancellationToken
            });
            if (!result.IsSuccess) throw new InvalidOperationException("WinHTTP restore failed: " + result.StandardError);
        }

        private static string CurrentUserSid()
        {
            using var identity = WindowsIdentity.GetCurrent();
            return identity.User?.Value ?? string.Empty;
        }

        internal void ImportProxyFromInternetSettings(UiLanguage language)
        {
            OnProgressChanged(T(language, "正在设置网络代理...", "Importing network proxy settings..."), NetworkRepairStatus.Processing);

            var executionResult = _scriptRunner.RunAsync(new ScriptExecutionRequest
            {
                FileName = NetshPath,
                Arguments = "winhttp import proxy source=ie",
                Timeout = TimeSpan.FromMinutes(1)
            }).GetAwaiter().GetResult();

            if (!executionResult.IsSuccess)
            {
                throw new InvalidOperationException(T(
                    language,
                    $"网络代理导入失败（返回码 {executionResult.ExitCode}）：{executionResult.StandardError}",
                    $"Proxy import failed (exit code {executionResult.ExitCode}): {executionResult.StandardError}"));
            }

            OnProgressChanged(T(language, "✓ 网络代理设置成功", "✓ Network proxy imported"), NetworkRepairStatus.Success);
        }

        private void OnProgressChanged(string message, NetworkRepairStatus status)
        {
            ProgressChanged?.Invoke(this, new NetworkRepairProgressEventArgs(message, status));
        }

        private static string T(UiLanguage language, string zh, string en)
        {
            return LocalizationService.Localize(language, zh, en);
        }
    }
}
