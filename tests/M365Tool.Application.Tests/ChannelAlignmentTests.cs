using System.Reflection;
using System.Xml.Linq;
using Office365CleanupTool.Models;
using Office365CleanupTool.Services;
using Xunit;

namespace M365Tool.Application.Tests;

public class ChannelAlignmentTests
{
    [Fact]
    public void SupportedChannelsIncludeDistinctMonthlyEnterpriseCdn()
    {
        var channels = new OfficeChannelService().GetSupportedChannels();
        Assert.Equal(3, channels.Count);
        Assert.Contains(channels, channel => channel.Channel == UpdateChannel.MonthlyEnterprise
            && channel.CdnBaseUrl.EndsWith("55336b82-a18d-4dd6-b5f6-9e5095c314a6"));
    }

    [Fact]
    public void MonthlyEnterpriseInstallationUsesItsOwnOdtChannel()
    {
        var xml = OfficeDeploymentXmlBuilder.Build(new InstallationConfig { Channel = UpdateChannel.MonthlyEnterprise }, Array.Empty<string>());
        Assert.Equal("MonthlyEnterprise", xml.Root!.Element("Add")!.Attribute("Channel")!.Value);
    }

    [Fact]
    public void ImportedMonthlyEnterpriseConfigurationKeepsItsChannel()
    {
        var config = OfficeDeploymentXmlBuilder.Parse(XDocument.Parse("<Configuration><Add Channel='MonthlyEnterprise' /></Configuration>"));
        Assert.Equal(UpdateChannel.MonthlyEnterprise, config.Channel);
    }

    [Fact]
    public void MonthlyEnterpriseVersionListUsesItsOwnBuilds()
    {
        string html = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", "office-history-20261029.html"));
        var method = typeof(OfficeChannelService).GetMethod("ParseVersionOptions", BindingFlags.Static | BindingFlags.NonPublic)!;
        var versions = (IReadOnlyList<OfficeVersionOption>)method.Invoke(null, new object[] { html, UpdateChannel.MonthlyEnterprise, 36 })!;
        Assert.Equal("20430.20180", versions[0].Build);
        Assert.DoesNotContain(versions, version => version.Build == "20430.20178");
    }
    [Theory]
    [InlineData(UpdateChannel.Monthly, "Current")]
    [InlineData(UpdateChannel.MonthlyEnterprise, "MonthlyEnterprise")]
    [InlineData(UpdateChannel.Broad, "SemiAnnual")]
    public void InstallationAndUpdateConfigurationUseTheSameChannel(UpdateChannel channel, string name)
    {
        var install = OfficeDeploymentXmlBuilder.Build(new InstallationConfig { Channel = channel }, Array.Empty<string>());
        Assert.Equal(name, install.Root!.Element("Add")!.Attribute("Channel")!.Value);
        Assert.Equal(channel, OfficeDeploymentXmlBuilder.Parse(install).Channel);
        var update = OfficeChannelService.BuildChannelConfiguration(channel, false, null);
        Assert.Equal(name, update.Root!.Element("Updates")!.Attribute("Channel")!.Value);
        Assert.Null(update.Root.Element("Add"));
        Assert.Null(update.Root.Element("Remove"));
    }

    [Theory]
    [InlineData("Broad", UpdateChannel.Broad)]
    [InlineData("Monthly", UpdateChannel.Monthly)]
    [InlineData(" monthlyenterprise ", UpdateChannel.MonthlyEnterprise)]
    [InlineData(null, UpdateChannel.Monthly)]
    public void ImportPreservesOldAliasesAndDefault(string? name, UpdateChannel expected)
    {
        var add = new XElement("Add");
        if (name != null) add.SetAttributeValue("Channel", name);
        Assert.Equal(expected, OfficeDeploymentXmlBuilder.Parse(new XDocument(new XElement("Configuration", add))).Channel);
    }

    [Theory]
    [InlineData("CurrentPreview")]
    [InlineData("SemiAnnualPreview")]
    [InlineData("PerpetualVL2024")]
    [InlineData("Unknown")]
    public void ImportDoesNotSilentlyChangeUnsupportedChannels(string name) =>
        Assert.Throws<InvalidOperationException>(() => OfficeDeploymentXmlBuilder.Parse(
            XDocument.Parse($"<Configuration><Add Channel='{name}' /></Configuration>")));

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LatestClearsPinAndAutomaticUpdatesRemainAnIndependentChoice(bool disableUpdates)
    {
        foreach (OfficeVersionOption? target in new OfficeVersionOption?[] { null, new() { IsLatest = true, Build = "20430.20180" } })
        {
            XElement updates = OfficeChannelService.BuildChannelConfiguration(UpdateChannel.MonthlyEnterprise, disableUpdates, target).Root!.Element("Updates")!;
            Assert.Equal("", updates.Attribute("TargetVersion")!.Value);
            Assert.Equal(disableUpdates ? "FALSE" : "TRUE", updates.Attribute("Enabled")!.Value);
        }
    }

    [Theory]
    [InlineData("20430.20180")]
    [InlineData("16.0.20430.20180")]
    public void PinnedBuildUsesTheSameFullVersionInOdtAndUpdateCommand(string build)
    {
        var target = new OfficeVersionOption { Build = build };
        var xml = OfficeChannelService.BuildChannelConfiguration(UpdateChannel.MonthlyEnterprise, false, target);
        Assert.Equal("16.0.20430.20180", xml.Root!.Element("Updates")!.Attribute("TargetVersion")!.Value);
        Assert.Contains("updatetoversion=16.0.20430.20180", OfficeChannelService.BuildUpdateArguments(build));
    }

    [Theory]
    [InlineData(UpdateChannel.Monthly, "492350f6-3a01-4f97-b9c0-c7c6ddf67d60")]
    [InlineData(UpdateChannel.MonthlyEnterprise, "55336b82-a18d-4dd6-b5f6-9e5095c314a6")]
    [InlineData(UpdateChannel.Broad, "7ffbc6bf-bc32-4f92-8982-f9dd17fd3114")]
    public void CdnAndLegacyEnumValuesRemainDistinct(UpdateChannel channel, string guid)
    {
        Assert.Equal(0, (int)UpdateChannel.Monthly);
        Assert.Equal(1, (int)UpdateChannel.Broad);
        var option = Assert.Single(new OfficeChannelService().GetSupportedChannels(), option => option.Channel == channel);
        Assert.EndsWith(guid, option.CdnBaseUrl);
    }

    [Theory]
    [InlineData("16.0.20430.20178", "O365ProPlusRetail", "月度企业频道（原半年配置）")]
    [InlineData("16.0.20131.20001", "O365BusinessRetail", "月度企业频道（原半年配置）")]
    [InlineData("16.0.20131.20000", "O365ProPlusRetail", "半年企业频道")]
    [InlineData("16.0.19127.20752", "O365ProPlusRetail", "半年企业频道")]
    [InlineData("", "O365ProPlusRetail", "半年企业频道")]
    [InlineData("invalid", "O365ProPlusRetail", "半年企业频道")]
    [InlineData("16.0.20430.20178", "ProPlus2024Volume", "半年企业频道")]
    [InlineData("16.0.20430.20178", "ProjectProRetail", "半年企业频道")]
    public void UnifiedDisplayRequiresAnActualM365Apps2606OrLaterBuild(string version, string product, string expected)
    {
        Assert.Equal(expected, OfficeInventoryService.ResolveChannelDisplayName("SemiAnnual", null, version, product));
        Assert.Equal(expected, OfficeInventoryService.ResolveChannelDisplayName(null,
            "http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114", version, product));
        Assert.Equal("当前频道", OfficeInventoryService.ResolveChannelDisplayName("Current", "SemiAnnual", version, product));
    }

    [Theory]
    [InlineData(UpdateChannel.Monthly, "20527.20112", "20430.20180", "20430.20178")]
    [InlineData(UpdateChannel.MonthlyEnterprise, "20430.20180", "20527.20112", "20430.20178")]
    [InlineData(UpdateChannel.Broad, "20430.20178", "20527.20112", "20430.20180")]
    public void UnificationDoesNotMixBuildsAcrossChannelColumns(UpdateChannel channel, string expected, string other1, string other2)
    {
        string html = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", "office-history-20261029.html"));
        var method = typeof(OfficeChannelService).GetMethod("ParseVersionOptions", BindingFlags.Static | BindingFlags.NonPublic)!;
        var versions = (IReadOnlyList<OfficeVersionOption>)method.Invoke(null, new object[] { html, channel, 36 })!;
        Assert.Equal(expected, versions[0].Build);
        Assert.DoesNotContain(versions, version => version.Build == other1 || version.Build == other2);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task OdtConfigurationCompletesBeforeTheManualUpdateStarts(bool disableUpdates)
    {
        string directory = NewTestDirectory();
        try
        {
            var runner = new RecordingRunner();
            var service = new OfficeChannelService(runner);
            var target = new OfficeVersionOption { Build = "20430.20180" };
            XDocument xml = OfficeChannelService.BuildChannelConfiguration(UpdateChannel.MonthlyEnterprise, disableUpdates, target);
            await service.ApplyChannelConfigurationAsync("setup.exe", "OfficeC2RClient.exe", xml, directory, target, UiLanguage.English);
            Assert.Equal(2, runner.Requests.Count);
            Assert.Equal("setup.exe", runner.Requests[0].FileName);
            Assert.Equal($"/configure \"{Path.Combine(directory, "OfficeChannel.xml")}\"", runner.Requests[0].Arguments);
            Assert.Equal(directory, runner.Requests[0].WorkingDirectory);
            var saved = XDocument.Load(Path.Combine(directory, "OfficeChannel.xml"));
            Assert.Equal("MonthlyEnterprise", saved.Root!.Element("Updates")!.Attribute("Channel")!.Value);
            Assert.Equal(disableUpdates ? "FALSE" : "TRUE", saved.Root.Element("Updates")!.Attribute("Enabled")!.Value);
            Assert.Equal("OfficeC2RClient.exe", runner.Requests[1].FileName);
            Assert.Contains("updatetoversion=16.0.20430.20180", runner.Requests[1].Arguments);
            Assert.Contains("forceappshutdown=false", runner.Requests[1].Arguments);
            Assert.All(runner.Requests, request => Assert.False(request.KillOnCancellation));
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData(1, false, false)]
    [InlineData(-1, true, true)]
    [InlineData(-1, true, false)]
    public async Task FailedOrPendingOdtDoesNotStartUpdateOnTheOldChannel(int exitCode, bool timedOut, bool running)
    {
        string directory = NewTestDirectory();
        try
        {
            var runner = new RecordingRunner { FirstResult = new() { ExitCode = exitCode, TimedOut = timedOut, IsStillRunning = running } };
            var service = new OfficeChannelService(runner);
            await Assert.ThrowsAsync<InvalidOperationException>(() => service.ApplyChannelConfigurationAsync(
                "setup.exe", "OfficeC2RClient.exe", OfficeChannelService.BuildChannelConfiguration(UpdateChannel.MonthlyEnterprise, false, null),
                directory, null, UiLanguage.English));
            Assert.Single(runner.Requests);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public async Task FailedUpdateClientIsReportedAfterOdtSuccess()
    {
        string directory = NewTestDirectory();
        try
        {
            var runner = new RecordingRunner { SecondResult = new() { ExitCode = 5 } };
            await Assert.ThrowsAsync<InvalidOperationException>(() => new OfficeChannelService(runner).ApplyChannelConfigurationAsync(
                "setup.exe", "OfficeC2RClient.exe", OfficeChannelService.BuildChannelConfiguration(UpdateChannel.MonthlyEnterprise, false, null),
                directory, null, UiLanguage.English));
            Assert.Equal(2, runner.Requests.Count);
        }
        finally { Directory.Delete(directory, true); }
    }

    private static string NewTestDirectory()
    {
        string path = Path.Combine(Path.GetTempPath(), "M365Tool-channel-tests", Guid.NewGuid().ToString("N"), "频道配置 with spaces");
        Directory.CreateDirectory(path);
        return path;
    }

    private sealed class RecordingRunner : IScriptRunner
    {
        public List<ScriptExecutionRequest> Requests { get; } = new();
        public ScriptExecutionResult FirstResult { get; init; } = new();
        public ScriptExecutionResult SecondResult { get; init; } = new();
        public Task<ScriptExecutionResult> RunAsync(ScriptExecutionRequest request)
        {
            Requests.Add(request);
            return Task.FromResult(Requests.Count == 1 ? FirstResult : SecondResult);
        }
        public void StartDetached(ScriptExecutionRequest request) => throw new InvalidOperationException("No actual maintenance in tests.");
    }

}
