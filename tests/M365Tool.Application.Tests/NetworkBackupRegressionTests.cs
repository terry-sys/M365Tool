using Microsoft.Win32;
using Office365CleanupTool.Services;
using Xunit;

namespace M365Tool.Application.Tests;

[Collection("Shared maintenance operation")]
public sealed class NetworkBackupRegressionTests : IDisposable
{
    private readonly string _keyPath = @"Software\M365Tool.Tests\" + Guid.NewGuid().ToString("N");
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "M365Tool-BackupTest-" + Guid.NewGuid().ToString("N"));

    [Fact]
    public async Task RepairPreservesPacAndRestoreRecoversExistingAndMissingValues()
    {
        using (var key = Registry.CurrentUser.CreateSubKey(_keyPath))
        {
            key.SetValue("SecureProtocols", 128, RegistryValueKind.DWord);
            key.SetValue("AutoConfigURL", @"%USERPROFILE%\corporate.pac", RegistryValueKind.ExpandString);
        }
        var runner = new BackupRunner();
        var service = new NetworkRepairService(runner, _keyPath, _directory);

        await service.RunAsync(UiLanguage.English);

        using (var key = Registry.CurrentUser.OpenSubKey(_keyPath, true))
        {
            Assert.Equal(128 | 0x800, key!.GetValue("SecureProtocols"));
            Assert.Null(key.GetValue("AutoDetect"));
            Assert.Equal(@"%USERPROFILE%\corporate.pac", key.GetValue("AutoConfigURL", null, RegistryValueOptions.DoNotExpandEnvironmentNames));
            Assert.Equal(RegistryValueKind.ExpandString, key.GetValueKind("AutoConfigURL"));
            key.SetValue("AutoDetect", 1);
        }
        Assert.True(File.Exists(service.LastBackupPath));
        Assert.DoesNotContain(runner.Arguments, argument => argument.Contains("import proxy"));

        await service.RestoreAsync(service.LastBackupPath);

        using (var key = Registry.CurrentUser.OpenSubKey(_keyPath))
        {
            Assert.Equal(128, key!.GetValue("SecureProtocols"));
            Assert.Null(key.GetValue("AutoDetect"));
        }
        Assert.Contains(runner.Arguments, argument => argument.StartsWith("exec "));
        Assert.Equal(BackupRunner.ProxyScript, await File.ReadAllTextAsync(Assert.Single(Directory.GetFiles(_directory, "restore-winhttp-*.txt"))));
    }

    [Fact]
    public async Task FailedBackupLeavesSettingsUntouched()
    {
        using (var key = Registry.CurrentUser.CreateSubKey(_keyPath)) key.SetValue("SecureProtocols", 128);
        var service = new NetworkRepairService(new BackupRunner { FailDump = true }, _keyPath, _directory);

        await Assert.ThrowsAsync<InvalidOperationException>(() => service.RunAsync(UiLanguage.English));

        using var unchanged = Registry.CurrentUser.OpenSubKey(_keyPath);
        Assert.Equal(128, unchanged!.GetValue("SecureProtocols"));
    }

    public void Dispose()
    {
        // Only this test's GUID key and temporary directory are removed.
        Registry.CurrentUser.DeleteSubKeyTree(_keyPath, throwOnMissingSubKey: false);
        if (Directory.Exists(_directory)) Directory.Delete(_directory, true);
    }

    private sealed class BackupRunner : IScriptRunner
    {
        public const string ProxyScript = "pushd winhttp\r\nset proxy proxy-server=\"proxy.example:8080\"\r\npopd\r\n";
        public bool FailDump { get; init; }
        public List<string> Arguments { get; } = new();
        public Task<ScriptExecutionResult> RunAsync(ScriptExecutionRequest request)
        {
            Arguments.Add(request.Arguments);
            return Task.FromResult(new ScriptExecutionResult
            {
                ExitCode = request.Arguments == "winhttp dump" && FailDump ? 1 : 0,
                StandardOutput = request.Arguments == "winhttp dump" ? ProxyScript : string.Empty
            });
        }
        public void StartDetached(ScriptExecutionRequest request) => throw new InvalidOperationException();
    }
}
