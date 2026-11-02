using System.Reflection;
using System.Runtime.ExceptionServices;
using Microsoft.Web.WebView2.Core;
using Office365CleanupTool;
using Office365CleanupTool.Models;
using Office365CleanupTool.Services;
using Xunit;

namespace M365Tool.Application.Tests;

public class FunctionalSafetyRegressionTests
{
    [Theory]
    [InlineData(null)]
    [InlineData("20326.20142")]
    public void ChannelUpdatesNeverForceCloseApplications(string? build)
    {
        string arguments = OfficeChannelService.BuildUpdateArguments(build);
        Assert.Contains("forceappshutdown=false", arguments);
        Assert.DoesNotContain("forceappshutdown=true", arguments);
        if (build != null) Assert.Contains("updatetoversion=16.0.20326.20142", arguments);
    }

    [Theory]
    [InlineData("https://myaccount.windowsazure.cn/", true, false, true)]
    [InlineData("https://myaccount.windowsazure.cn/", false, false, false)]
    [InlineData("https://myaccount.windowsazure.cn/", true, true, false)]
    [InlineData("http://myaccount.windowsazure.cn/", true, false, false)]
    [InlineData("https://login.partner.microsoftonline.cn/", true, false, false)]
    [InlineData("https://myaccount.windowsazure.cn.example.com/", true, false, false)]
    public void ProfileIdentityFallbackRequiresSuccessfulHttpsPortalNavigation(string url, bool completed, bool error, bool expected) =>
        Assert.Equal(expected, MyAccountLoginVerificationForm.CanReadPortalIdentity(new Uri(url), completed, error));

    [Fact]
    public void LegacyVerificationCacheRequiresFreshSignIn()
    {
        var settings = new AppSettings { HasValidated21VAccount = true, LastValidatedAccount = "user@example.com" };
        Assert.False(settings.HasCurrentVerification);
        settings.VerificationVersion = AppSettings.CurrentVerificationVersion;
        Assert.True(settings.HasCurrentVerification);
        settings.LastValidatedAccount = "";
        settings.LastValidatedDisplayName = "Display name only";
        Assert.False(settings.HasCurrentVerification);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void DefaultInstallationPreservesExistingMsiProducts(bool project, bool visio)
    {
        var config = new InstallationConfig { IncludeProject = project, IncludeVisio = visio };
        var xml = OfficeDeploymentXmlBuilder.Build(config, Array.Empty<string>());
        Assert.Null(xml.Root!.Element("RemoveMSI"));
    }

    [Theory]
    [InlineData(UpdateChannel.Monthly, "20527.20112")]
    [InlineData(UpdateChannel.Broad, "20430.20178")]
    [InlineData(UpdateChannel.MonthlyEnterprise, "20430.20180")]
    public void VersionParserReadsActualChannelColumns(UpdateChannel channel, string latestBuild)
    {
        string html = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", "office-history-20261029.html"));
        var result = (IReadOnlyList<OfficeVersionOption>)InvokeStatic(typeof(OfficeChannelService), "ParseVersionOptions", html, channel, 36)!;
        Assert.NotEmpty(result);
        Assert.Equal(latestBuild, result[0].Build);
        Assert.True(result[0].IsLatest);
    }

    [Fact]
    public void VersionHistoryUsesHeadersWhenPreviewColumnIsPresentOrColumnsMove()
    {
        const string html = "<table><tr><th>Year</th><th>Release Date</th><th>Semi-Annual Enterprise Channel (Preview)</th><th>Semi-Annual Enterprise Channel</th><th>Current Channel</th></tr>" +
            "<tr><td>2026</td><td>October 13</td><td>Version 2610 (Build 20527.20000)</td><td>Version 2609 (Build 20430.20178)</td><td>Version 2610 (Build 20527.20112)</td></tr></table>";
        var result = (IReadOnlyList<OfficeVersionOption>)InvokeStatic(typeof(OfficeChannelService), "ParseVersionOptions", html, UpdateChannel.Broad, 36)!;
        Assert.Equal("20430.20178", Assert.Single(result).Build);
    }

    [Theory]
    [InlineData("01: An Offline scan was performed because Outlook is either not running or it is running elevated.", 1)]
    [InlineData("02: A Full scan was performed. See report.xml", 2)]
    public void OfficialOutlookSuccessMessagesAreStillRecognized(string output, int expected)
    {
        int? code = (int?)InvokeStatic(typeof(OutlookDiagnosticsService), "DetectOutlookScanCode", output, "", 0);
        Assert.Equal(expected, code);
        Assert.True((bool)InvokeStatic(typeof(OutlookDiagnosticsService), "DetermineSuccess",
            new OutlookDiagnosticResult { ExitCode = 0, DetectedScenarioCode = code, StandardOutput = output }, "OutlookScan")!);
    }

    [Theory]
    [InlineData("2026-10-29 12:01:04 Step 2: checking prerequisites")]
    [InlineData("Step 1: checking prerequisites")]
    public void UnrelatedNumbersDoNotProvideScanEvidence(string output) =>
        Assert.Null(InvokeStatic(typeof(OutlookDiagnosticsService), "DetectOutlookScanCode", output, "", 0));

    [Fact]
    public void OldHarvestedSuccessCannotReplaceCurrentFailure()
    {
        string directory = Path.Combine(Path.GetTempPath(), "M365Tool-ScanResult-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            string path = Path.Combine(directory, "old-success.log");
            DateTime started = new DateTime(2026, 10, 29, 10, 0, 0, DateTimeKind.Utc);
            File.WriteAllText(path, "01: An Offline scan was performed.");
            File.SetLastWriteTimeUtc(path, started.AddHours(-1));
            Assert.Null(InvokeStatic(typeof(OutlookDiagnosticsService), "DetectOutlookScanCodeFromHarvestedLogs", directory, started));
            File.WriteAllText(Path.Combine(directory, "current.log"), "05: An error occurred while performing a scan of Outlook.");
            File.SetLastWriteTimeUtc(Path.Combine(directory, "current.log"), started.AddSeconds(1));
            Assert.Equal(5, (int?)InvokeStatic(typeof(OutlookDiagnosticsService), "DetectOutlookScanCodeFromHarvestedLogs", directory, started));
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("2026-10-29 12:01:04 ExecutionResult: Failed")]
    [InlineData("2026-10-29 12:02:04 ExecutionResult: Failed")]
    [InlineData("Step 1: checking prerequisites.\n05: scan failed")]
    public void OutlookFailureCannotBecomeSuccessThroughTimestampOrStepNumber(string output)
    {
        var code = (int?)InvokeStatic(typeof(OutlookDiagnosticsService), "DetectOutlookScanCode", output, "", 5);
        var result = new OutlookDiagnosticResult { ExitCode = 5, DetectedScenarioCode = code, StandardOutput = output };
        Assert.False((bool)InvokeStatic(typeof(OutlookDiagnosticsService), "DetermineSuccess", result, "OutlookScan")!);
    }

    [Fact]
    public void UnsupportedOfflineScanDoesNotLaunchDefaultScan()
    {
        var error = Assert.Throws<TargetInvocationException>(() => InvokeStatic(typeof(OutlookDiagnosticsService), "BuildArguments",
            "-S ExpertExperienceAdminTask -AcceptEula -LogFolder \"{LOG}\"", @"C:\Logs", true, new GetHelpCapabilities()));
        Assert.IsType<NotSupportedException>(error.InnerException);
    }

    [Fact]
    public void SupportedOfflineScanUsesExplicitSwitch()
    {
        string args = (string)InvokeStatic(typeof(OutlookDiagnosticsService), "BuildArguments", "-S ExpertExperienceAdminTask",
            @"C:\Logs", true, new GetHelpCapabilities { SupportsOfflineScanSwitch = true })!;
        Assert.Contains(" -OfflineScan", args);
    }

    [Fact]
    public async Task InteractiveScriptPreservesUnicodePathsAndExitCode()
    {
        string directory = Path.Combine(Path.GetTempPath(), "M365Tool-中文 路径-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            string script = Path.Combine(directory, "启动.cmd");
            string exitCode = Path.Combine(directory, "退出码.txt");
            string cmd = Path.Combine(Environment.SystemDirectory, "cmd.exe");
            OutlookDiagnosticsService.WriteInteractiveScript(script, directory, cmd, "/d /c exit /b 7", exitCode);
            var result = await new ScriptRunner().RunAsync(new ScriptExecutionRequest
            {
                FileName = cmd, Arguments = $"/d /c \"\"{script}\"\"", WorkingDirectory = directory,
                Timeout = TimeSpan.FromSeconds(10), KillOnCancellation = true
            });
            Assert.Equal(7, result.ExitCode);
            Assert.Equal("7", File.ReadAllText(exitCode).Trim());
            Assert.DoesNotContain("??", File.ReadAllText(script));
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void FailedNavigationWithoutAuthenticatedReturnDoesNotGrantAccess() => RunSta(() =>
    {
        using var dialog = new MyAccountLoginVerificationForm(UiLanguage.English);
        var type = dialog.GetType();
        type.GetField("_navigationCompleted", BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(dialog, false);
        type.GetField("_lastNavigationErrorStatus", BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(dialog, CoreWebView2WebErrorStatus.CannotConnect);
        var task = (Task)type.GetMethod("TryMarkVerifiedAsync", BindingFlags.Instance | BindingFlags.NonPublic)!
            .Invoke(dialog, new object[] { new Uri("https://myaccount.windowsazure.cn/") })!;
        task.GetAwaiter().GetResult();
        Assert.False(dialog.IsVerified);
    });

    private static object? InvokeStatic(Type type, string method, params object[] args) =>
        type.GetMethod(method, BindingFlags.Static | BindingFlags.NonPublic)!.Invoke(null, args);

    private static void RunSta(Action test)
    {
        Exception? error = null;
        var thread = new Thread(() =>
        {
            try { test(); }
            catch (Exception ex) { error = ex; }
        }) { IsBackground = true };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        Assert.True(thread.Join(TimeSpan.FromSeconds(10)), "Login regression test timed out.");
        if (error != null) ExceptionDispatchInfo.Capture(error).Throw();
    }
}
