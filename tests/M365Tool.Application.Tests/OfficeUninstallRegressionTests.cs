using Office365CleanupTool.Services;
using Office365CleanupTool.Models;
using Xunit;

namespace M365Tool.Application.Tests;

public class OfficeUninstallRegressionTests
{
    [Theory]
    [InlineData("LibreOffice 24.2", null, false)]
    [InlineData("LibreOffice 24.2", "The Document Foundation", false)]
    [InlineData("OfficeSuite", "MobiSystems", false)]
    [InlineData("ProjectLibre", null, false)]
    [InlineData("Microsoft Office Professional Plus 2016", "Microsoft Corporation", true)]
    [InlineData("Microsoft 365 Apps for enterprise - en-us", "Microsoft Corporation", true)]
    [InlineData("Microsoft Visio Professional 2021", null, true)]
    [InlineData("Microsoft Project Professional 2019", "Microsoft", true)]
    public void InventoryOnlyMatchesMicrosoftOfficeProducts(string name, string? publisher, bool expected)
    {
        Assert.Equal(expected, OfficeInventoryService.LooksLikeOfficeProduct(name, publisher));
    }

    [Theory]
    [InlineData("Microsoft Office Professional Plus 2016", "", "16.0.12345.1234", "2016")]
    [InlineData("Microsoft 365 Apps for enterprise", "", "16.0.12345.1234", "M365")]
    [InlineData("", "O365ProPlusRetail", "16.0.12345.1234", "M365")]
    [InlineData("", "ProPlus2021Volume", "16.0.12345.1234", "2021")]
    [InlineData("", "VisioPro2019Volume", "16.0.12345.1234", "2019")]
    [InlineData("", "ProjectPro2024Retail", "16.0.12345.1234", "2024")]
    [InlineData("Microsoft Office", "", "15.0.12345.1234", "2013")]
    [InlineData("Microsoft Office", "", "14.0.12345.1234", "2010")]
    [InlineData("Microsoft Office", "", "12.0.12345.1234", "2007")]
    [InlineData("Microsoft Office", "", "16.0.12345.1234", "")]
    public void InventoryRetainsVersionIdentityWithoutGuessingFromMajor16(
        string name, string id, string build, string expected)
    {
        Assert.Equal(expected, OfficeInventoryService.ResolveOfficeVersion(name, id, build));
    }

    [Fact]
    public void KeepingOtherVersionsDoesNotBlockSelectedVersionUninstall()
    {
        var target = new OfficeUninstallTargetOption { SaraOfficeVersion = "2016" };
        var remaining = new[] { Product("M365"), Product("2021") };

        Assert.All(remaining, item => Assert.False(OfficeUninstallService.IsUninstallTarget(item, target)));
        Assert.False(OfficeUninstallService.HasRemainingTargetOrUnknownVersion(remaining, target));
    }

    [Theory]
    [InlineData("2016")]
    [InlineData("M365")]
    [InlineData("All")]
    [InlineData(null)]
    public void RemainingTargetPreventsCompletion(string? version)
    {
        var target = new OfficeUninstallTargetOption { SaraOfficeVersion = version };
        var remaining = new[] { Product(version is null or "All" ? "2021" : version) };

        Assert.True(OfficeUninstallService.IsUninstallTarget(remaining[0], target));
        Assert.True(OfficeUninstallService.HasRemainingTargetOrUnknownVersion(remaining, target));
    }

    [Fact]
    public void UnknownVersionRequiresAttentionWithoutWaitingForUnrelatedTarget()
    {
        var target = new OfficeUninstallTargetOption { SaraOfficeVersion = "2016" };
        var remaining = new[] { Product("") };

        Assert.False(OfficeUninstallService.IsUninstallTarget(remaining[0], target));
        Assert.True(OfficeUninstallService.HasRemainingTargetOrUnknownVersion(remaining, target));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ExitZeroWithRemainingTargetDoesNotReportCompletion(bool background)
    {
        string summary = OfficeUninstallService.BuildSummary(0, background, true, UiLanguage.English);

        Assert.Contains("Completion cannot be confirmed", summary);
        Assert.DoesNotContain("completed successfully", summary);
    }

    [Theory]
    [InlineData("All")]
    [InlineData("2016")]
    [InlineData(null)]
    public void EmptyInventoryDoesNotRequireAttention(string? version)
    {
        Assert.False(OfficeUninstallService.HasRemainingTargetOrUnknownVersion(
            Array.Empty<InstalledOfficeInfo>(), new OfficeUninstallTargetOption { SaraOfficeVersion = version }));
    }

    private static InstalledOfficeInfo Product(string version) => new() { OfficeVersion = version };
}
