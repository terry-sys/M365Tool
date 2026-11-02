using Office365CleanupTool.Services;
using Xunit;

namespace M365Tool.Application.Tests;

public class MaintenanceRegressionTests
{
    [Theory]
    [InlineData("20326.20142", "16.0.20326.20142")]
    [InlineData("16.0.20326.20142", "16.0.20326.20142")]
    [InlineData(" 20326.20142 ", "16.0.20326.20142")]
    public void OfficeBuildUsesFullVersion(string input, string expected)
    {
        Assert.Equal(expected, OfficeChannelService.NormalizeBuildVersion(input));
    }

    [Theory]
    [InlineData("")]
    [InlineData("invalid")]
    [InlineData("16.0.20326")]
    [InlineData("15.0.20326.20142")]
    [InlineData("20326.20142 extra")]
    public void InvalidBuildIsRejectedBeforeUpdatingRegistry(string input)
    {
        Assert.Throws<ArgumentException>(() => OfficeChannelService.NormalizeBuildVersion(input));
    }

    [Fact]
    public void TeamsAddinRepairIncludesMandatoryOutlookSwitch()
    {
        string arguments = TeamsMaintenanceService.BuildAddinRepairArguments(@"C:\Test Logs");

        Assert.Contains("-S TeamsAddinScenario", arguments);
        Assert.Contains("-AcceptEula", arguments);
        Assert.Contains("-CloseOutlook", arguments);
        Assert.Contains("-LogFolder \"C:\\Test Logs\"", arguments);
    }

    [Fact]
    public void FailedProxyImportPropagatesFailureWithoutSuccessEvent()
    {
        var service = new NetworkRepairService(new FakeScriptRunner(1));
        var statuses = new List<NetworkRepairStatus>();
        service.ProgressChanged += (_, e) => statuses.Add(e.Status);

        var error = Assert.Throws<InvalidOperationException>(
            () => service.ImportProxyFromInternetSettings(UiLanguage.English));

        Assert.Contains("exit code 1", error.Message);
        Assert.DoesNotContain(NetworkRepairStatus.Success, statuses);
    }

    [Fact]
    public void SuccessfulProxyImportStillReportsSuccess()
    {
        var service = new NetworkRepairService(new FakeScriptRunner(0));
        var statuses = new List<NetworkRepairStatus>();
        service.ProgressChanged += (_, e) => statuses.Add(e.Status);

        service.ImportProxyFromInternetSettings(UiLanguage.English);

        Assert.Contains(NetworkRepairStatus.Success, statuses);
    }

    private sealed class FakeScriptRunner(int exitCode) : IScriptRunner
    {
        public Task<ScriptExecutionResult> RunAsync(ScriptExecutionRequest request) =>
            Task.FromResult(new ScriptExecutionResult { ExitCode = exitCode, StandardError = "test output" });

        public void StartDetached(ScriptExecutionRequest request) =>
            throw new InvalidOperationException("External process execution is disabled in tests.");
    }
}
