using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Windows.Forms;
using Office365CleanupTool;
using Office365CleanupTool.Models;
using Office365CleanupTool.Services;
using Xunit;

namespace M365Tool.Application.Tests;

public class UiRegressionTests
{
    [Theory]
    [InlineData(700, UiLanguage.Chinese)]
    [InlineData(1100, UiLanguage.Chinese)]
    [InlineData(1600, UiLanguage.Chinese)]
    [InlineData(700, UiLanguage.English)]
    [InlineData(1100, UiLanguage.English)]
    [InlineData(1600, UiLanguage.English)]
    public void PrecheckTitleDoesNotCoverRecheckButton(int width, UiLanguage language) => RunSta(() =>
    {
        using var control = new WorkbenchInstallationControl(language) { Size = new Size(width, 650) };
        Invoke(control, "LayoutControls");
        var title = Field<Label>(control, "_lblPrecheckTitle");
        var button = Field<Button>(control, "_btnRefreshChecks");
        Assert.False(title.Bounds.IntersectsWith(button.Bounds),
            $"The title {title.Bounds} covers the recheck button {button.Bounds}.");
        var center = new Point(button.Left + button.Width / 2, button.Top + button.Height / 2);
        Assert.Same(button, button.Parent!.GetChildAtPoint(center));
    });

    [Theory]
    [InlineData("zh-CN", "zh-CN")]
    [InlineData("zh-HK", "zh-TW")]
    [InlineData("zh-Hant", "zh-TW")]
    [InlineData("en-AU", "en-US")]
    [InlineData("en-GB", "en-GB")]
    [InlineData("fr-CA", "fr-FR")]
    [InlineData("is-IS", "en-US")]
    public void SystemLanguageResolvesToOneSupportedOfficeLanguage(string system, string expected) =>
        Assert.Equal(expected, WorkbenchInstallationControl.ResolveSupportedSystemLanguage(system));

    [Fact]
    public void OnlyOneLanguageIsMarkedAsSystemDefault() => RunSta(() =>
    {
        using var control = new WorkbenchInstallationControl(UiLanguage.English);
        var list = Field<CheckedListBox>(control, "_clbLanguages");
        Assert.Single(list.Items.Cast<object>(), item => item.ToString()!.Contains("System default"));
        Assert.Single(list.CheckedItems.Cast<object>());
    });

    [Fact]
    public void FilteringLanguagesDoesNotDiscardHiddenSelections() => RunSta(() =>
    {
        using var control = new WorkbenchInstallationControl(UiLanguage.English);
        var list = Field<CheckedListBox>(control, "_clbLanguages");
        var expected = list.CheckedItems.Cast<object>().Select(Code).Append("en-US").ToHashSet(StringComparer.OrdinalIgnoreCase);
        list.SetItemChecked(list.Items.Cast<object>().Select(Code).ToList().IndexOf("en-US"), true);
        Invoke(control, "UpdateConfiguration");
        Field<TextBox>(control, "_txtLanguageSearch").Text = "en-US";
        // A normal option change while the list is filtered updates the configuration.
        Field<ComboBox>(control, "_cboEdition").SelectedIndex = 1;
        Field<TextBox>(control, "_txtLanguageSearch").Clear();
        Assert.True(expected.SetEquals(list.CheckedItems.Cast<object>().Select(Code)));
        Assert.True(expected.SetEquals(Field<InstallationConfig>(control, "_currentConfig").LanguageIDs));
    });

    [Fact]
    public void ClearingLanguageSelectionDoesNotRecheckDefaultsOnSearch() => RunSta(() =>
    {
        using var control = new WorkbenchInstallationControl(UiLanguage.English);
        var list = Field<CheckedListBox>(control, "_clbLanguages");
        foreach (int index in list.CheckedIndices.Cast<int>().ToArray()) list.SetItemChecked(index, false);
        Invoke(control, "UpdateConfiguration");
        Field<TextBox>(control, "_txtLanguageSearch").Text = "en";
        Field<TextBox>(control, "_txtLanguageSearch").Clear();
        Assert.Empty(list.CheckedItems.Cast<object>());
    });

    [Fact]
    public void ChangingLanguagePreservesInstallFailureStatus() => RunSta(() =>
    {
        using var control = new WorkbenchInstallationControl(UiLanguage.Chinese);
        Field<Label>(control, "_lblResult").Text = "安装失败";
        Field<TextBox>(control, "_txtResultDetails").Text = "Failure details";
        control.ApplyLanguage(UiLanguage.English);
        Assert.Equal("Install failed", Field<Label>(control, "_lblResult").Text);
        Assert.Equal("Failure details", Field<TextBox>(control, "_txtResultDetails").Text);
    });

    [Fact]
    public void ChangingLanguagePreservesCancelActionDuringInstall() => RunSta(() =>
    {
        using var control = new WorkbenchInstallationControl(UiLanguage.Chinese);
        SetField(control, "_isInstalling", true);
        Invoke(control, "UpdateUiState");
        control.ApplyLanguage(UiLanguage.English);
        Assert.Contains("Stop Waiting", Field<Button>(control, "_btnInstall").Text);
    });

    [Fact]
    public void ChangingLanguagePreservesCancelActionDuringUninstall() => RunSta(() =>
    {
        using var control = new WorkbenchUninstallControl(new OfficeUninstallService(new NeverRun()));
        SetField(control, "_isRunning", true);
        Invoke(control, "UpdateUiState");
        control.ApplyLanguage(UiLanguage.English);
        Assert.Contains("Stop Waiting", Field<Button>(control, "_btnStartUninstall").Text);
    });

    [Fact]
    public void OlderChannelResponseCannotReplaceNewerSelection() => RunSta(() =>
    {
        var service = new PendingChannels();
        using var control = new WorkbenchChannelControl(service);
        var first = (Task)Invoke(control, "LoadVersionsAsync")!;
        Field<ComboBox>(control, "_cboChannel").SelectedIndex = 2;
        service.Broad.SetResult(Versions("semi-build"));
        PumpUntil(() => Field<List<OfficeVersionOption>>(control, "_loadedVersions").FirstOrDefault()?.Build == "semi-build");
        service.Monthly.SetResult(Versions("current-build"));
        PumpUntil(() => first.IsCompleted);
        first.GetAwaiter().GetResult();
        Assert.Equal("semi-build", Assert.Single(Field<List<OfficeVersionOption>>(control, "_loadedVersions")).Build);
    });

    [Fact]
    public void LanguageChangeCannotEnableStaleVersionsDuringLoading() => RunSta(() =>
    {
        var service = new PendingChannels();
        using var control = new WorkbenchChannelControl(service);
        SetField(control, "_loadedVersions", Versions("stale-build").ToList());
        var loading = (Task)Invoke(control, "LoadVersionsAsync")!;
        control.ApplyLanguage(UiLanguage.English);
        bool enabledDuringLoading = Field<ComboBox>(control, "_cboVersion").Enabled;
        service.Monthly.SetResult(Versions("fresh-build"));
        PumpUntil(() => loading.IsCompleted);
        loading.GetAwaiter().GetResult();
        Assert.False(enabledDuringLoading);
    });

    [Theory]
    [InlineData(UiLanguage.Chinese)]
    [InlineData(UiLanguage.English)]
    public void InstallationPreservesAllThreeChannelsAcrossLanguageChanges(UiLanguage language) => RunSta(() =>
    {
        using var control = new WorkbenchInstallationControl(language, new PendingValidation());
        var combo = Field<ComboBox>(control, "_cboChannel");
        Assert.Equal(3, combo.Items.Count);
        Assert.Equal(UpdateChannel.Monthly, Field<InstallationConfig>(control, "_currentConfig").Channel);
        foreach (var choice in new[] { (1, UpdateChannel.MonthlyEnterprise), (2, UpdateChannel.Broad), (0, UpdateChannel.Monthly) })
        {
            combo.SelectedIndex = choice.Item1;
            Assert.Equal(choice.Item2, Field<InstallationConfig>(control, "_currentConfig").Channel);
            control.ApplyLanguage(language == UiLanguage.Chinese ? UiLanguage.English : UiLanguage.Chinese);
            Assert.Equal(choice.Item1, combo.SelectedIndex);
            Assert.Equal(choice.Item2, Field<InstallationConfig>(control, "_currentConfig").Channel);
        }
    });

    [Fact]
    public void LegacyChannelResponseCannotReplaceMonthlyEnterpriseBuilds() => RunSta(() =>
    {
        var service = new PendingChannels();
        using var control = new WorkbenchChannelControl(service);
        var current = (Task)Invoke(control, "LoadVersionsAsync")!;
        var combo = Field<ComboBox>(control, "_cboChannel");
        combo.SelectedIndex = 2;
        combo.SelectedIndex = 1;
        control.ApplyLanguage(UiLanguage.English);
        Assert.Equal(1, combo.SelectedIndex);
        service.Enterprise.SetResult(Versions("20430.20180"));
        PumpUntil(() => Field<List<OfficeVersionOption>>(control, "_loadedVersions").FirstOrDefault()?.Build == "20430.20180");
        service.Broad.SetResult(Versions("20430.20178"));
        service.Monthly.SetResult(Versions("20527.20112"));
        PumpUntil(() => current.IsCompleted);
        current.GetAwaiter().GetResult();
        Assert.Equal("20430.20180", Assert.Single(Field<List<OfficeVersionOption>>(control, "_loadedVersions")).Build);
        Assert.Contains("Monthly Enterprise", combo.SelectedItem!.ToString());
    });

    private static IReadOnlyList<OfficeVersionOption> Versions(string build) =>
        new[] { new OfficeVersionOption { Build = build, IsLatest = true } };

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OlderPrecheckCannotOverwriteTheLatestReport(bool olderFails) => RunSta(() =>
    {
        var service = new PendingValidation();
        using var control = new WorkbenchInstallationControl(UiLanguage.English, service);
        var first = (Task)Invoke(control, "RefreshPrecheckAsync")!;
        var second = (Task)Invoke(control, "RefreshPrecheckAsync")!;
        service.Second.SetResult(Report("new report"));
        PumpUntil(() => second.IsCompleted);
        if (olderFails) service.First.SetException(new InvalidOperationException("old failure"));
        else service.First.SetResult(Report("old report"));
        PumpUntil(() => first.IsCompleted);
        first.GetAwaiter().GetResult();
        Assert.Contains("new report", Field<TextBox>(control, "_txtPrecheckResult").Text);
    });

    [Fact]
    public void OlderPrecheckCompletionCannotEnableRecheckDuringANewerCheck() => RunSta(() =>
    {
        var service = new PendingValidation();
        using var control = new WorkbenchInstallationControl(UiLanguage.English, service);
        var first = (Task)Invoke(control, "RefreshPrecheckAsync")!;
        var second = (Task)Invoke(control, "RefreshPrecheckAsync")!;
        service.First.SetResult(Report("old report"));
        PumpUntil(() => first.IsCompleted);
        bool enabledBeforeLatestCompleted = Field<Button>(control, "_btnRefreshChecks").Enabled;
        service.Second.SetResult(Report("new report"));
        PumpUntil(() => second.IsCompleted);
        Assert.False(enabledBeforeLatestCompleted);
        Assert.True(Field<Button>(control, "_btnRefreshChecks").Enabled);
    });

    private static ValidationReport Report(string detail) => new(new[] { new ValidationCheckResult("check", true, detail) });

    [Fact]
    public void LanguageChangeDoesNotEnableRecheckWhileAPrecheckIsRunning() => RunSta(() =>
    {
        var service = new PendingValidation();
        using var control = new WorkbenchInstallationControl(UiLanguage.English, service);
        var checking = (Task)Invoke(control, "RefreshPrecheckAsync")!;
        control.ApplyLanguage(UiLanguage.Chinese);
        bool enabledDuringCheck = Field<Button>(control, "_btnRefreshChecks").Enabled;
        service.First.SetResult(Report("report"));
        PumpUntil(() => checking.IsCompleted);
        Assert.False(enabledDuringCheck);
    });

    private sealed class PendingValidation : ValidationService
    {
        private int _count;
        public TaskCompletionSource<ValidationReport> First { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public TaskCompletionSource<ValidationReport> Second { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public override Task<ValidationReport> GetValidationReportAsync(UiLanguage language = UiLanguage.Chinese) =>
            ++_count == 1 ? First.Task : Second.Task;
    }

    private static string Code(object item) => item.ToString()!.Split(" - ")[0];
    private static T Field<T>(object target, string name) => (T)target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(target)!;
    private static void SetField(object target, string name, object value) => target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(target, value);
    private static object? Invoke(object target, string name) => target.GetType().GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(target, null);

    private static void RunSta(Action test)
    {
        Exception? error = null;
        var thread = new Thread(() =>
        {
            try
            {
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
                test();
            }
            catch (Exception ex) { error = ex; }
        }) { IsBackground = true };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        Assert.True(thread.Join(TimeSpan.FromSeconds(15)), "UI test timed out.");
        if (error != null) ExceptionDispatchInfo.Capture(error).Throw();
    }

    private static void PumpUntil(Func<bool> completed)
    {
        var timer = Stopwatch.StartNew();
        while (!completed() && timer.Elapsed < TimeSpan.FromSeconds(5)) { System.Windows.Forms.Application.DoEvents(); Thread.Sleep(1); }
        Assert.True(completed(), "UI continuation did not complete.");
    }

    private sealed class PendingChannels : OfficeChannelService
    {
        public TaskCompletionSource<IReadOnlyList<OfficeVersionOption>> Monthly { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public TaskCompletionSource<IReadOnlyList<OfficeVersionOption>> Enterprise { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public TaskCompletionSource<IReadOnlyList<OfficeVersionOption>> Broad { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public override Task<IReadOnlyList<OfficeVersionOption>> GetAvailableVersionsAsync(UpdateChannel channel, int maxResults = 36) =>
            channel switch
            {
                UpdateChannel.Monthly => Monthly.Task,
                UpdateChannel.MonthlyEnterprise => Enterprise.Task,
                UpdateChannel.Broad => Broad.Task,
                _ => throw new ArgumentOutOfRangeException(nameof(channel))
            };
    }
    private sealed class NeverRun : IScriptRunner
    {
        public Task<ScriptExecutionResult> RunAsync(ScriptExecutionRequest request) => throw new InvalidOperationException("Maintenance is disabled in UI tests.");
        public void StartDetached(ScriptExecutionRequest request) => throw new InvalidOperationException("Maintenance is disabled in UI tests.");
    }
}
