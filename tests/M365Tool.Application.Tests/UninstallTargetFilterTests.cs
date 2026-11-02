using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Windows.Forms;
using Office365CleanupTool;
using Office365CleanupTool.Models;
using Office365CleanupTool.Services;
using Xunit;

namespace M365Tool.Application.Tests;

public class UninstallTargetFilterTests
{
    [Fact]
    public void StandardModeDoesNotOfferOffice2007() => RunSta(() =>
    {
        using var control = new WorkbenchUninstallControl(new FixtureService { Products = new[] { Product("M365", "O365ProPlusRetail"), Product("2007", "msi-2007", "卸载注册表(32 位)") } });
        Assert.Equal(0, Field<ComboBox>(control, "_cboEngine").SelectedIndex);
        Assert.DoesNotContain(Field<ComboBox>(control, "_cboTarget").Items.Cast<object>(), item => item.ToString() == "Office 2007");
    });

    [Fact]
    public void StandardTargetsOnlyContainDetectedC2rVersions()
    {
        var products = new[]
        {
            Product("M365", "O365ProPlusRetail"),
            Product("2021", "ProPlus2021Volume"),
            Product("2007", "msi-2007", "卸载注册表(32 位)"),
            Product("2016", "msi-2016", "卸载注册表(64 位)")
        };
        var service = new FixtureService();
        var targets = service.GetAvailableTargets(OfficeUninstallMode.Standard, products);
        Assert.Equal(new[] { "M365", "2021", "All" }, targets.Select(target => target.SaraOfficeVersion));
        foreach (var target in targets)
        {
            var xml = OfficeUninstallService.BuildStandardRemovalXml(target, products);
            Assert.All(xml.Descendants("Product"), product => Assert.Contains(product.Attribute("ID")!.Value, new[] { "O365ProPlusRetail", "ProPlus2021Volume" }));
        }
        var allXml = OfficeUninstallService.BuildStandardRemovalXml(targets.Single(target => target.SaraOfficeVersion == "All"), products);
        Assert.Equal(2, allXml.Descendants("Product").Count());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void EmptyMsiOnlyAndC2rWithoutProductIdHaveNoStandardTargets(int scenario)
    {
        var products = scenario switch
        {
            0 => Array.Empty<InstalledOfficeInfo>(),
            1 => new[] { Product("2007", "msi-2007", "卸载注册表(32 位)") },
            _ => new[] { Product("M365", " ") }
        };
        Assert.Empty(new FixtureService().GetAvailableTargets(OfficeUninstallMode.Standard, products));
    }

    [Fact]
    public void UnknownC2rVersionStillHasAnExplicitC2rOnlyAllOption()
    {
        var products = new[] { Product("", "ProjectProRetail"), Product("2007", "msi-2007", "卸载注册表(32 位)") };
        var target = Assert.Single(new FixtureService().GetAvailableTargets(OfficeUninstallMode.Standard, products));
        Assert.Equal("All", target.SaraOfficeVersion);
        var xml = OfficeUninstallService.BuildStandardRemovalXml(target, products);
        Assert.Equal("ProjectProRetail", Assert.Single(xml.Descendants("Product")).Attribute("ID")!.Value);
    }

    [Fact]
    public void DeepCleanupKeepsItsFullSupportedListWithoutAnInventoryRequirement()
    {
        var service = new FixtureService();
        var targets = service.GetAvailableTargets(OfficeUninstallMode.DeepCleanup, Array.Empty<InstalledOfficeInfo>());
        Assert.Equal(service.GetSupportedTargets().Select(target => target.SaraOfficeVersion), targets.Select(target => target.SaraOfficeVersion));
        Assert.Contains(targets, target => target.UsesDetectedInstalledVersion);
        Assert.Contains(targets, target => target.SaraOfficeVersion == "2007");
    }

    [Theory]
    [InlineData(UiLanguage.Chinese)]
    [InlineData(UiLanguage.English)]
    public void SwitchingEnginesRemovesAnIncompatibleTarget(UiLanguage language) => RunSta(() =>
    {
        using var control = new WorkbenchUninstallControl(new FixtureService { Products = new[] { Product("M365", "O365ProPlusRetail") } });
        control.ApplyLanguage(language);
        var engine = Field<ComboBox>(control, "_cboEngine");
        var target = Field<ComboBox>(control, "_cboTarget");
        Assert.Equal(new[] { "Microsoft 365 Apps", language == UiLanguage.Chinese ? "全部已检测到的 C2R 产品" : "All detected C2R products" }, Captions(target));
        engine.SelectedIndex = 1;
        target.SelectedIndex = Array.IndexOf(Captions(target), "Office 2007");
        Assert.Equal("Office 2007", target.SelectedItem!.ToString());
        engine.SelectedIndex = 0;
        Assert.DoesNotContain("Office 2007", Captions(target));
        Assert.Equal("Microsoft 365 Apps", target.SelectedItem!.ToString());
        Assert.True(Field<Button>(control, "_btnStartUninstall").Enabled);
    });

    [Theory]
    [InlineData(UiLanguage.Chinese)]
    [InlineData(UiLanguage.English)]
    public void LanguageChangeKeepsEngineAndValidSelection(UiLanguage language) => RunSta(() =>
    {
        using var control = new WorkbenchUninstallControl(new FixtureService { Products = new[] { Product("M365", "O365ProPlusRetail") } });
        control.ApplyLanguage(language);
        var engine = Field<ComboBox>(control, "_cboEngine");
        var target = Field<ComboBox>(control, "_cboTarget");
        target.SelectedIndex = 1;
        control.ApplyLanguage(language == UiLanguage.Chinese ? UiLanguage.English : UiLanguage.Chinese);
        Assert.Equal(0, engine.SelectedIndex);
        Assert.Equal(1, target.SelectedIndex);
        Assert.DoesNotContain("Office 2007", Captions(target));
        engine.SelectedIndex = 1;
        target.SelectedIndex = Array.IndexOf(Captions(target), "Office 2007");
        control.ApplyLanguage(language);
        Assert.Equal(1, engine.SelectedIndex);
        Assert.Equal("Office 2007", target.SelectedItem!.ToString());
    });

    [Theory]
    [InlineData(UiLanguage.Chinese)]
    [InlineData(UiLanguage.English)]
    public void MsiOnlyInventoryDisablesStandardStartButAllowsDeepCleanup(UiLanguage language) => RunSta(() =>
    {
        using var control = new WorkbenchUninstallControl(new FixtureService { Products = new[] { Product("2007", "msi-2007", "卸载注册表(32 位)") } });
        control.ApplyLanguage(language);
        var engine = Field<ComboBox>(control, "_cboEngine");
        var target = Field<ComboBox>(control, "_cboTarget");
        Assert.Empty(target.Items.Cast<object>());
        Assert.Null(target.SelectedItem);
        Assert.False(target.Enabled);
        Assert.False(Field<Button>(control, "_btnStartUninstall").Enabled);
        Assert.False(Field<CheckBox>(control, "_chkForceClose").Enabled);
        Assert.True(engine.Enabled);
        Assert.Contains(language == UiLanguage.Chinese ? "未检测到" : "No C2R products", Field<Label>(control, "_lblDescription").Text);
        engine.SelectedIndex = 1;
        Assert.Contains("Office 2007", Captions(target));
        Assert.True(Field<Button>(control, "_btnStartUninstall").Enabled);
        engine.SelectedIndex = 0;
        Assert.Empty(target.Items.Cast<object>());
        Assert.False(Field<Button>(control, "_btnStartUninstall").Enabled);
    });

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RevisitingPageRefreshesAvailableTargets(bool added) => RunSta(() =>
    {
        var service = new FixtureService();
        var detected = new[] { Product("2024", "ProPlus2024Volume") };
        service.Products = added ? Array.Empty<InstalledOfficeInfo>() : detected;
        using var control = new WorkbenchUninstallControl(service);
        service.Products = added ? detected : Array.Empty<InstalledOfficeInfo>();
        control.PrepareForDisplay();
        var target = Field<ComboBox>(control, "_cboTarget");
        Assert.Equal(added, Field<Button>(control, "_btnStartUninstall").Enabled);
        if (added) Assert.Contains("Office 2024", Captions(target));
        else Assert.Empty(target.Items.Cast<object>());
    });

    [Fact]
    public void EmptyTargetsDoNotDisableAnActiveCancelAction() => RunSta(() =>
    {
        using var control = new WorkbenchUninstallControl(new FixtureService());
        using var cancellation = new CancellationTokenSource();
        SetField(control, "_isRunning", true);
        SetField(control, "_cancellation", cancellation);
        control.ApplyLanguage(UiLanguage.English);
        var button = Field<Button>(control, "_btnStartUninstall");
        Assert.Contains("Stop Waiting", button.Text);
        Assert.True(button.Enabled);
        cancellation.Cancel();
        typeof(WorkbenchUninstallControl).GetMethod("UpdateUiState", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(control, null);
        Assert.False(button.Enabled);
    });

    [Fact]
    public void RefreshKeepsExecutionResultText() => RunSta(() =>
    {
        using var control = new WorkbenchUninstallControl(new FixtureService());
        var result = Field<TextBox>(control, "_txtResult");
        result.Text = "Fixture uninstall failure details";
        control.PrepareForDisplay();
        control.ApplyLanguage(UiLanguage.English);
        Assert.Equal("Fixture uninstall failure details", result.Text);
    });

    private static string[] Captions(ComboBox target) => target.Items.Cast<object>().Select(item => item.ToString()!).ToArray();
    private static InstalledOfficeInfo Product(string version, string id, string source = "ClickToRun") =>
        new() { OfficeVersion = version, ProductId = id, Source = source, DisplayName = id };
    private static void SetField(object target, string name, object value) =>
        target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(target, value);
    private sealed class FixtureService : OfficeUninstallService
    {
        public IReadOnlyList<InstalledOfficeInfo> Products { get; set; } = Array.Empty<InstalledOfficeInfo>();
        public FixtureService() : base(new NeverRun()) { }
        public override IReadOnlyList<InstalledOfficeInfo> DetectInstalledOffice() => Products;
    }

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
    private sealed class NeverRun : IScriptRunner
    {
        public Task<ScriptExecutionResult> RunAsync(ScriptExecutionRequest request) => throw new InvalidOperationException("No actual uninstall in tests.");
        public void StartDetached(ScriptExecutionRequest request) => throw new InvalidOperationException("No actual uninstall in tests.");
    }
}
