using System.Diagnostics;
using System.Text;
using System.Text.Json;
using Microsoft.Win32;
using Office365CleanupTool.Models;
using Office365CleanupTool.Services;
using Xunit;

namespace M365Tool.Application.Tests;

public class ExecutionReliabilityTests
{
    [Fact]
    public async Task RunnerCollectsBothStreamsAndExitCode()
    {
        var result = await new ScriptRunner().RunAsync(new ScriptExecutionRequest
        {
            FileName = "cmd.exe", Arguments = "/d /c \"echo output & echo diagnostic 1>&2 & exit /b 7\""
        });
        Assert.Equal(7, result.ExitCode);
        Assert.Contains("output", result.StandardOutput);
        Assert.Contains("diagnostic", result.StandardError);
        Assert.False(result.IsSuccess);
    }

    [Fact]
    public async Task AlreadyCanceledRequestDoesNotStartProcess()
    {
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new ScriptRunner().RunAsync(new ScriptExecutionRequest
        {
            FileName = "must-not-be-launched.exe", CancellationToken = new CancellationToken(true)
        }));
    }

    [Fact]
    public async Task TimeoutStopsMonitoringWithoutKillingMaintenanceProcess()
    {
        var result = await new ScriptRunner().RunAsync(new ScriptExecutionRequest
        {
            FileName = "powershell.exe", Arguments = "-NoProfile -NonInteractive -Command Start-Sleep -Seconds 30",
            Timeout = TimeSpan.FromMilliseconds(300)
        });
        try
        {
            Assert.True(result.TimedOut);
            Assert.True(result.IsStillRunning);
            Assert.False(result.IsSuccess);
            using var process = Process.GetProcessById(result.ProcessId);
            Assert.False(process.HasExited);
        }
        finally { await StopTestChild(result.ProcessId); }
    }

    [Fact]
    public async Task CancelStopsMonitoringAndRetainsPartialOutput()
    {
        using var cancellation = new CancellationTokenSource(TimeSpan.FromSeconds(10));
        var result = await new ScriptRunner().RunAsync(new ScriptExecutionRequest
        {
            FileName = "powershell.exe",
            Arguments = "-NoProfile -NonInteractive -Command \"Write-Output 'ready'; Start-Sleep -Seconds 30\"",
            OnOutputDataReceived = _ => cancellation.Cancel(), CancellationToken = cancellation.Token
        });
        try
        {
            Assert.True(result.WasCanceled);
            Assert.False(result.TimedOut);
            Assert.True(result.IsStillRunning);
            Assert.Contains("ready", result.StandardOutput);
        }
        finally { await StopTestChild(result.ProcessId); }
    }

    [Fact]
    public async Task ReadOnlyProbeCanOptIntoTerminationOnTimeout()
    {
        var result = await new ScriptRunner().RunAsync(new ScriptExecutionRequest
        {
            FileName = "powershell.exe", Arguments = "-NoProfile -NonInteractive -Command Start-Sleep -Seconds 30",
            Timeout = TimeSpan.FromMilliseconds(300), KillOnCancellation = true
        });
        Assert.True(result.TimedOut);
        Assert.False(result.IsStillRunning);
        Assert.False(result.IsSuccess);
    }

    [Fact]
    public void StandardRemovalPreservesOtherVersionsAndMsiProducts()
    {
        var target = new OfficeUninstallTargetOption { SaraOfficeVersion = "2021" };
        var inventory = new[]
        {
            new InstalledOfficeInfo { Source = "ClickToRun", OfficeVersion = "2021", ProductId = "ProPlus2021Volume" },
            new InstalledOfficeInfo { Source = "ClickToRun", OfficeVersion = "M365", ProductId = "O365ProPlusRetail" },
            new InstalledOfficeInfo { Source = "MSI", OfficeVersion = "2021", ProductId = "msi-product" }
        };
        var xml = OfficeUninstallService.BuildStandardRemovalXml(target, inventory);
        Assert.Equal("ProPlus2021Volume", Assert.Single(xml.Root!.Element("Remove")!.Elements("Product")).Attribute("ID")!.Value);
        Assert.Null(xml.Root.Element("RemoveMSI"));
        Assert.DoesNotContain("O365ProPlusRetail", xml.ToString());
        Assert.DoesNotContain("msi-product", xml.ToString());
        Assert.Equal(2, OfficeUninstallService.GetPostCheckScope(inventory, OfficeUninstallMode.Standard).Count);
        Assert.Equal(3, OfficeUninstallService.GetPostCheckScope(inventory, OfficeUninstallMode.DeepCleanup).Count);
    }

    [Fact]
    public void StandardAutoDetectionRequiresExplicitChoiceForMixedVersions()
    {
        Assert.Throws<InvalidOperationException>(() => OfficeUninstallService.BuildStandardRemovalXml(
            new OfficeUninstallTargetOption(), new[]
            {
                new InstalledOfficeInfo { Source = "ClickToRun", OfficeVersion = "M365", ProductId = "O365ProPlusRetail" },
                new InstalledOfficeInfo { Source = "ClickToRun", OfficeVersion = "2021", ProductId = "VisioPro2021Volume" }
            }));
    }

    [Fact]
    public void StandardRemovalDoesNotSilentlyFallBackToDestructiveCleanup()
    {
        Assert.Throws<InvalidOperationException>(() => OfficeUninstallService.BuildStandardRemovalXml(
            new OfficeUninstallTargetOption { SaraOfficeVersion = "All" },
            new[] { new InstalledOfficeInfo { Source = "MSI", ProductId = "msi" } }));
    }

    [Fact]
    public async Task InvalidSignatureBlocksToolReuse()
    {
        string directory = Path.Combine(Path.GetTempPath(), "M365Tool-SignatureTest-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        await File.WriteAllTextAsync(Path.Combine(directory, "setup.exe"), "unsigned test file");
        try
        {
            var runner = new RefusingRunner();
            await Assert.ThrowsAsync<InvalidOperationException>(() => new OdtToolService(runner).PrepareAsync(directory, default, _ => { }));
            Assert.Equal(1, runner.Calls);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void SignatureProbeQuotesPathsAndRequiresMicrosoftSigner()
    {
        string arguments = MicrosoftToolVerifier.BuildVerificationArguments(@"C:\Test's Folder\setup.exe");
        string script = Encoding.Unicode.GetString(Convert.FromBase64String(arguments.Split(' ').Last()));
        Assert.Contains("Test''s Folder", script);
        Assert.Contains("Get-AuthenticodeSignature -LiteralPath", script);
        Assert.Contains("$s.Status -ne 'Valid'", script);
        Assert.Contains("Microsoft Corporation", script);
    }

    [Fact]
    public void NetworkBackupRoundTripsMissingValuesAndPacValueKinds()
    {
        var backup = Backup();
        var restored = JsonSerializer.Deserialize<NetworkRepairBackup>(JsonSerializer.Serialize(backup))!;
        NetworkRepairBackup.Validate(restored);
        Assert.False(restored.Values[1].Exists);
        Assert.Equal(RegistryValueKind.ExpandString, restored.Values[2].Kind);
        Assert.Equal(@"%USERPROFILE%\proxy.pac", restored.Values[2].Text);
    }

    [Fact]
    public void InvalidBackupCannotNameUnrelatedRegistryValues()
    {
        var backup = Backup();
        backup.Values[0].Name = "UnrelatedSetting";
        Assert.Throws<InvalidDataException>(() => NetworkRepairBackup.Validate(backup));
    }

    [Fact]
    public void BackupCannotExecuteCommandsOutsideProxyRestoration()
    {
        var backup = Backup();
        backup.WinHttpScript = "pushd winhttp\r\nadd helper untrusted.dll\r\npopd";
        Assert.Throws<InvalidDataException>(() => NetworkRepairBackup.Validate(backup));
    }

    [Fact]
    public void DefaultNetworkRepairPreservesCorporateProxySettings()
    {
        var options = new NetworkRepairOptions();
        Assert.True(options.EnableTls12);
        Assert.False(options.EnableAutoDetect);
        Assert.False(options.ImportUserProxy);
    }

    private static NetworkRepairBackup Backup() => new()
    {
        WinHttpScript = "pushd winhttp\r\nreset proxy\r\npopd",
        Values = new()
        {
            new() { Name = "SecureProtocols", Exists = true, Kind = RegistryValueKind.DWord, Number = 2048 },
            new() { Name = "AutoDetect", Exists = false },
            new() { Name = "AutoConfigURL", Exists = true, Kind = RegistryValueKind.ExpandString, Text = @"%USERPROFILE%\proxy.pac" }
        }
    };

    private static async Task StopTestChild(int id)
    {
        try
        {
            using var process = Process.GetProcessById(id);
            if (!process.HasExited) process.Kill(entireProcessTree: true);
            await process.WaitForExitAsync();
        }
        catch (ArgumentException) { }
    }

    private sealed class RefusingRunner : IScriptRunner
    {
        public int Calls { get; private set; }
        public Task<ScriptExecutionResult> RunAsync(ScriptExecutionRequest request)
        {
            Calls++;
            Assert.EndsWith("powershell.exe", request.FileName, StringComparison.OrdinalIgnoreCase);
            return Task.FromResult(new ScriptExecutionResult { ExitCode = 1, StandardError = "invalid signature" });
        }
        public void StartDetached(ScriptExecutionRequest request) => throw new InvalidOperationException();
    }
}
