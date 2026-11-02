using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Windows.Forms;
using Office365CleanupTool;
using Office365CleanupTool.Models;
using Office365CleanupTool.Services;
using Xunit;

namespace M365Tool.Application.Tests;

public class OfficeInventoryDisplayTests
{
    [Theory]
    [InlineData("O365ProPlusRetail", UiLanguage.Chinese, "Microsoft 365 企业应用版")]
    [InlineData("o365proplusretail", UiLanguage.English, "Microsoft 365 Apps for enterprise")]
    [InlineData("O365BusinessRetail", UiLanguage.Chinese, "Microsoft 365 商业应用版")]
    [InlineData("O365BusinessRetail", UiLanguage.English, "Microsoft 365 Apps for business")]
    public void ProductNamesDistinguishEnterpriseAndBusiness(string id, UiLanguage language, string expected) =>
        Assert.Equal(expected, OfficeInventoryService.GetClickToRunDisplayName(id, language));

    [Theory]
    [InlineData("O365UnknownRetail")]
    [InlineData("M365FutureRetail")]
    public void UnknownSubscriptionEditionRetainsItsProductId(string id)
    {
        string name = OfficeInventoryService.GetClickToRunDisplayName(id);
        Assert.Contains(id, name);
        Assert.DoesNotContain("企业应用版", name);
        Assert.DoesNotContain("商业应用版", name);
    }

    [Theory]
    [InlineData("492350f6-3a01-4f97-b9c0-c7c6ddf67d60", "当前频道")]
    [InlineData("7ffbc6bf-bc32-4f92-8982-f9dd17fd3114", "半年企业频道")]
    [InlineData("55336b82-a18d-4dd6-b5f6-9e5095c314a6", "月度企业频道")]
    [InlineData("64256afe-f5d9-4f86-8936-8840a6a4f5be", "当前频道（预览）")]
    [InlineData("b8f9b850-328d-4355-9145-c59439a0c4cf", "半年企业频道（预览）")]
    [InlineData("5440fd1f-7ecb-4221-8110-145efaa6372f", "Beta 频道")]
    public void ChannelDetectionUnderstandsRegistryCdnUrls(string guid, string expected)
    {
        Assert.Equal(expected, OfficeInventoryService.ResolveChannelDisplayName($"http://officecdn.microsoft.com/pr/{guid}"));
        Assert.Equal(expected, OfficeInventoryService.ResolveChannelDisplayName($"https://officecdn.microsoft.com/pr/{guid.ToUpperInvariant()}/?source=test"));
    }

    [Theory]
    [InlineData("Current", "当前频道")]
    [InlineData("Broad", "半年企业频道")]
    [InlineData("SemiAnnual", "半年企业频道")]
    [InlineData("MonthlyEnterprise", "月度企业频道")]
    [InlineData("", "")]
    [InlineData("http://officecdn.microsoft.com/pr/unknown", "其他频道")]
    public void ChannelAliasesAndUnknownValuesStayExplicit(string value, string expected) =>
        Assert.Equal(expected, OfficeInventoryService.ResolveChannelDisplayName(value));

    [Theory]
    [InlineData("http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60", "当前频道")]
    [InlineData(null, "半年企业频道")]
    [InlineData("", "半年企业频道")]
    [InlineData("   ", "半年企业频道")]
    [InlineData("http://officecdn.microsoft.com/pr/unknown", "其他频道")]
    public void CurrentUpdateChannelTakesPriorityOverInstallTimeCdn(string? current, string expected) =>
        Assert.Equal(expected, OfficeInventoryService.ResolveChannelDisplayName(current,
            "http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114"));

    [Theory]
    [InlineData(UiLanguage.Chinese, "Microsoft 365 企业应用版", "半年企业频道")]
    [InlineData(UiLanguage.English, "Microsoft 365 Apps for enterprise", "Semi-Annual Enterprise Channel")]
    public void HomeShowsTheDetectedEditionAndChannelInItsLanguage(UiLanguage language, string name, string channel) => RunSta(() =>
    {
        using var home = new WorkbenchHomeControl(new AppSettingsService(), new AppSettings(), language);
        var office = new InstalledOfficeProduct
        {
            ProductId = "O365ProPlusRetail", DisplayName = "Microsoft 365 企业应用版", ProductFamily = "Office",
            Source = "ClickToRun", Version = "16.0.20430.20178", Architecture = "x64 位", Channel = "半年企业频道"
        };
        SetInventory(home, new[] { new InstalledOfficeProduct { ProductFamily = "Visio", Source = "ClickToRun", DisplayName = "Visio" }, office });
        home.ApplyLanguage(language);
        Label[] values = Field<Label[]>(home, "_environmentValues");
        Assert.Equal(name, values[0].Text);
        Assert.Equal("16.0.20430.20178", values[1].Text);
        Assert.Equal(channel, values[2].Text);
    });

    [Theory]
    [InlineData(UiLanguage.Chinese, "月度企业频道（原半年配置）")]
    [InlineData(UiLanguage.English, "Monthly Enterprise (legacy SAEC)")]
    public void HomeExplainsUnifiedExperienceWithoutHidingTheLegacyConfiguration(UiLanguage language, string expected) => RunSta(() =>
    {
        using var home = new WorkbenchHomeControl(new AppSettingsService(), new AppSettings(), language);
        SetInventory(home, new[] { new InstalledOfficeProduct
        {
            ProductId = "O365ProPlusRetail", ProductFamily = "Office", Source = "ClickToRun", Version = "16.0.20430.20178",
            Channel = OfficeInventoryService.ResolveChannelDisplayName("SemiAnnual", null, "16.0.20430.20178", "O365ProPlusRetail")
        } });
        home.ApplyLanguage(language);
        Assert.Equal(expected, Field<Label[]>(home, "_environmentValues")[2].Text);
    });

    [Fact]
    public void HomeDoesNotReplaceAnOfficeSuiteWithClickToRunVisio() => RunSta(() =>
    {
        using var home = new WorkbenchHomeControl(new AppSettingsService(), new AppSettings(), UiLanguage.Chinese);
        SetInventory(home, new[]
        {
            new InstalledOfficeProduct { ProductFamily = "Visio", Source = "ClickToRun", DisplayName = "Visio", Version = "16.0" },
            new InstalledOfficeProduct { ProductFamily = "Office", Source = "卸载注册表(64 位)", DisplayName = "Microsoft Office Professional Plus 2016", Version = "16.1" }
        });
        home.ApplyLanguage(UiLanguage.Chinese);
        Label[] values = Field<Label[]>(home, "_environmentValues");
        Assert.Equal("Microsoft Office Professional Plus 2016", values[0].Text);
        Assert.Equal("16.1", values[1].Text);
    });

    private static void SetInventory(WorkbenchHomeControl home, IReadOnlyList<InstalledOfficeProduct> products) =>
        typeof(WorkbenchHomeControl).GetField("_products", BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(home, products);
    private static T Field<T>(object target, string name) =>
        (T)target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(target)!;
    private static void RunSta(Action action)
    {
        Exception? error = null;
        var thread = new Thread(() => { try { action(); } catch (Exception ex) { error = ex; } }) { IsBackground = true };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        Assert.True(thread.Join(TimeSpan.FromSeconds(15)), "UI test timed out.");
        if (error != null) ExceptionDispatchInfo.Capture(error).Throw();
    }
}
