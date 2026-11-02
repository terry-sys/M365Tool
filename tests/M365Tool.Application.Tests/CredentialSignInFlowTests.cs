using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Office365CleanupTool;
using Office365CleanupTool.Services;
using Xunit;

namespace M365Tool.Application.Tests;

public class CredentialSignInFlowTests
{
    [Fact]
    public void AuthenticatedReturnEntersToolWithoutRenderingProfile() => RunSta(() =>
    {
        using var form = ReadyForm();
        Assert.Null(Field<WebView2>(form, "_webView").CoreWebView2);
        Assert.True(form.TryCompleteCredentialVerification(new Uri("https://myaccount.windowsazure.cn/"), true));
        Assert.True(form.IsVerified);
        Assert.Equal(DialogResult.OK, form.DialogResult);
        Assert.Equal("user@tenant.onmschina.cn", form.VerifiedAccount);
        Assert.Null(form.VerifiedAvatarBytes);
        Assert.False(form.HasVerifiedProfileDisplayName);
    });

    [Fact]
    public void ClosingLoginDialogKeepsResourcesForDeferredSyncUntilDisposal() => RunSta(() =>
    {
        using var form = ReadyForm();
        var timer = Field<System.Windows.Forms.Timer>(form, "_verificationTimer");
        bool disposed = false;
        timer.Disposed += (_, _) => disposed = true;
        form.TryCompleteCredentialVerification(new Uri("https://myaccount.windowsazure.cn/"), true);
        typeof(MyAccountLoginVerificationForm).GetMethod("OnFormClosed", BindingFlags.Instance | BindingFlags.NonPublic)!
            .Invoke(form, new object[] { new FormClosedEventArgs(CloseReason.UserClosing) });
        Assert.False(disposed);
        form.Dispose();
        Assert.True(disposed);
    });

    [Fact]
    public void DeferredProfileWindowDoesNotTakeFocusFromTheTool() => RunSta(() =>
    {
        using var form = ReadyForm();
        SetField(form, "_deferredProfileSyncStarted", true);
        Assert.True((bool)typeof(MyAccountLoginVerificationForm)
            .GetProperty("ShowWithoutActivation", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(form)!);
    });

    [Theory]
    [InlineData(false, true, true, "user@tenant.onmschina.cn")]
    [InlineData(true, false, true, "user@tenant.onmschina.cn")]
    [InlineData(true, true, false, "user@tenant.onmschina.cn")]
    [InlineData(true, true, true, "")]
    public void IncompleteOrFailedAuthenticationCannotEnter(bool session, bool visitedSignIn, bool authenticatedReturn, string account) => RunSta(() =>
    {
        using var form = ReadyForm();
        SetField(form, "_hasVisitedChinaSignIn", visitedSignIn);
        SetField(form, "_hasReturnedFromChinaSignIn", authenticatedReturn);
        SetField(form, "_candidateAccount", account);
        Assert.False(form.TryCompleteCredentialVerification(new Uri("https://myaccount.windowsazure.cn/"), session));
        Assert.False(form.IsVerified);
    });

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AuthenticatedReturnDoesNotWaitForTheProfilePage(bool profileFailed) => RunSta(() =>
    {
        using var form = ReadyForm();
        SetField(form, "_navigationCompleted", false);
        SetField(form, "_lastNavigationErrorStatus", profileFailed ? CoreWebView2WebErrorStatus.CannotConnect : null);
        Assert.True(form.TryCompleteCredentialVerification(new Uri("https://myaccount.windowsazure.cn/"), true));
    });

    [Theory]
    [InlineData("https://myaccount.windowsazure.cn/", false)]
    [InlineData("https://login.partner.microsoftonline.cn/", true)]
    [InlineData("https://login.chinacloudapi.cn/", true)]
    public void ProfilePageIsConcealedWhileSignInPageRemainsUsable(string uri, bool expected) =>
        Assert.Equal(expected, MyAccountLoginVerificationForm.ShouldShowSignInContent(new Uri(uri)));

    [Theory]
    [InlineData("https://login.partner.microsoftonline.cn/", true)]
    [InlineData("https://login.chinacloudapi.cn/", true)]
    [InlineData("https://login.microsoftonline.cn/", true)]
    [InlineData("https://login.microsoftonline.com/", false)]
    [InlineData("http://login.partner.microsoftonline.cn/", false)]
    [InlineData("https://login.partner.microsoftonline.cn.example.com/", false)]
    public void AccountCaptureOnlyAcceptsChinaMicrosoftSignInSources(string uri, bool expected) =>
        Assert.Equal(expected, MyAccountLoginVerificationForm.IsChinaSignInHost(new Uri(uri)));

    [Theory]
    [InlineData("ESTSAUTH", true)]
    [InlineData("ESTSAUTHPERSISTENT", true)]
    [InlineData("estsStateTransient", true)]
    [InlineData("estsStatePersistent", true)]
    [InlineData("ESTSAUTHLIGHT", false)]
    [InlineData("fpc", false)]
    [InlineData("esctx", false)]
    public void SignInProofUsesSessionCookiesRatherThanTelemetry(string cookie, bool expected) =>
        Assert.Equal(expected, MyAccountLoginVerificationForm.IsAuthenticationSessionCookie(cookie));

    [Theory]
    [InlineData("about:blank", false)]
    [InlineData("http://login.partner.microsoftonline.cn/", false)]
    [InlineData("https://myaccount.windowsazure.cn/", false)]
    [InlineData("/relative", false)]
    [InlineData("https://login.partner.microsoftonline.cn/", true)]
    public void EarlySurfaceRequiresHttpsAndNeverRevealsTheAccountPortal(string value, bool expected) =>
        Assert.Equal(expected, MyAccountLoginVerificationForm.ShouldShowSignInContent(new Uri(value, UriKind.RelativeOrAbsolute)));

    [Fact]
    public void PreparationSharesAnAlreadyPendingInitialization() => RunSta(() =>
    {
        using var form = new MyAccountLoginVerificationForm(UiLanguage.English);
        var pending = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        SetField(form, "_browserInitializationTask", pending.Task);
        Assert.Same(pending.Task, form.PrepareAsync());
        Assert.Same(pending.Task, form.PrepareAsync());
        Assert.False(Field<bool>(form, "_initialized"));
        Assert.False(Field<bool>(form, "_navigationStarted"));
        pending.SetResult();
    });

    [Fact]
    public void AVisibleLoginSurfaceDoesNotGrantAccess() => RunSta(() =>
    {
        using var form = new MyAccountLoginVerificationForm(UiLanguage.English);
        SetField(form, "_signInSurfaceReady", true);
        SetField(form, "_hasVisitedChinaSignIn", true);
        SetField(form, "_candidateAccount", "user@tenant.onmschina.cn");
        Assert.False(form.TryCompleteCredentialVerification(new Uri("https://myaccount.windowsazure.cn/"), false));
        Assert.False(form.IsVerified);
    });

    private static MyAccountLoginVerificationForm ReadyForm()
    {
        var form = new MyAccountLoginVerificationForm(UiLanguage.English);
        SetField(form, "_hasVisitedChinaSignIn", true);
        SetField(form, "_hasReturnedFromChinaSignIn", true);
        SetField(form, "_navigationCompleted", true);
        SetField(form, "_candidateAccount", "user@tenant.onmschina.cn");
        return form;
    }

    private static T Field<T>(object instance, string name) => (T)instance.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(instance)!;
    private static void SetField(object instance, string name, object? value) => instance.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(instance, value);
    private static void RunSta(Action action)
    {
        Exception? error = null;
        var thread = new Thread(() => { try { action(); } catch (Exception ex) { error = ex; } }) { IsBackground = true };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        Assert.True(thread.Join(TimeSpan.FromSeconds(10)), "Sign-in flow test timed out.");
        if (error != null) ExceptionDispatchInfo.Capture(error).Throw();
    }
}
