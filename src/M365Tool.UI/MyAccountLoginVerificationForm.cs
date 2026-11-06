using System;
using System.Drawing;
using System.Diagnostics;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Office365CleanupTool.Services;

namespace Office365CleanupTool
{
    public sealed class MyAccountLoginVerificationForm : Form
    {
        private const string MyAccountUrl = "https://myaccount.windowsazure.cn/";
        private const int SignInPageLoadTimeoutMs = 60000;
        private const int SignInPageRenderTimeoutMs = 45000;
        private const double SignInPageZoomFactor = 0.85d;

        private static readonly Regex NumericLikeRegex = new(
            @"^\d+(?:[.,]\d+)?$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);
        private static readonly Regex BlockedNameRegex = new(
            @"^(Microsoft 365|Office|Files|Create|Apps|Search|Outlook|Teams|Word|Excel|PowerPoint|OneNote|OneDrive|SharePoint|Admin|管理|创建|应用|搜索|文件|最近|已共享|收藏夹|类型|取消固定|更多|设置及其他)$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        private readonly UiLanguage _language;
        private readonly WebView2 _webView;
        private readonly Label _statusLabel;
        private readonly System.Windows.Forms.Timer _verificationTimer;
        private readonly System.Windows.Forms.Timer _navigationWatchdogTimer;
        private string _userDataFolder = string.Empty;
        private Task? _browserInitializationTask;
        private bool _signInSurfaceReady;
        private bool _isCheckingSurface;
        private readonly Stopwatch _loadClock = Stopwatch.StartNew();
        private readonly List<string> _loadTimings = new();
        private readonly HashSet<string> _timedStages = new();
        private readonly string _loadTraceId = Guid.NewGuid().ToString("N");
        private static readonly object LoadTraceLock = new();
        private long _clickStartedAt = -1;
        private int _flushedTimingCount;
        private bool _initialized;
        private bool _navigationStarted;
        private bool _navigationCompleted;
        private bool _isCompletingVerification;
        private bool _isCheckingVerification;
        private int _verificationNavigationId;
        private bool _hasVisitedChinaSignIn;
        private bool _hasReturnedFromChinaSignIn;
        private Uri? _lastNavigationTarget;
        private string _candidateAccount = string.Empty;
        private bool _isProfileSyncing;
        private bool _deferredProfileSyncStarted;
        private bool _resourcesDisposed;
        private CoreWebView2WebErrorStatus? _lastNavigationErrorStatus;
        private string _lastAvatarSource = string.Empty;
        private string _lastAvatarRect = string.Empty;
        private string _lastProfileCandidates = string.Empty;
        private string _lastProfileTraceReason = string.Empty;

        public bool IsVerified { get; private set; }

        public string VerifiedAccount { get; private set; } = string.Empty;

        public string VerifiedUserName { get; private set; } = string.Empty;

        public byte[]? VerifiedAvatarBytes { get; private set; }

        public bool HasVerifiedProfileDisplayName { get; private set; }

        protected override bool ShowWithoutActivation => _deferredProfileSyncStarted;

        public MyAccountLoginVerificationForm(UiLanguage language)
        {
            _language = language;
            Text = string.Empty;
            ShowIcon = false;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            ShowInTaskbar = false;
            MinimizeBox = false;
            MaximizeBox = false;
            ClientSize = new System.Drawing.Size(1160, 820);
            MinimumSize = Size;
            MaximumSize = Size;
            BackColor = System.Drawing.Color.White;

            _webView = new WebView2
            {
                Dock = DockStyle.Fill
            };

            _statusLabel = new Label
            {
                Dock = DockStyle.Fill,
                Text = T("正在初始化登录页面...", "Initializing sign-in page..."),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = WorkbenchUi.CreateUiFont(11F),
                ForeColor = Color.FromArgb(86, 105, 130),
                BackColor = Color.White
            };

            _verificationTimer = new System.Windows.Forms.Timer
            {
                Interval = 250
            };
            _verificationTimer.Tick += VerificationTimer_Tick;

            _navigationWatchdogTimer = new System.Windows.Forms.Timer
            {
                Interval = SignInPageLoadTimeoutMs
            };
            _navigationWatchdogTimer.Tick += NavigationWatchdogTimer_Tick;

            Controls.Add(_webView);
            Controls.Add(_statusLabel);
            _statusLabel.BringToFront();
            Shown += async (_, _) => await StartMandatorySignInFlowAsync();
        }

        protected override void OnLoad(EventArgs e)
        {
            FitToWorkingArea(Screen.FromControl(Owner ?? this).WorkingArea);
            base.OnLoad(e);
        }

        internal void FitToWorkingArea(Rectangle workingArea)
        {
            var fittedSize = new Size(Math.Min(Width, workingArea.Width), Math.Min(Height, workingArea.Height));
            MinimumSize = Size.Empty;
            MaximumSize = Size.Empty;
            Size = fittedSize;
            MinimumSize = Size;
            MaximumSize = Size;
        }

        internal Task PrepareAsync()
        {
            if (IsDisposed || Disposing) return Task.FromException(new ObjectDisposedException(nameof(MyAccountLoginVerificationForm)));
            return _browserInitializationTask ??= InitializeBrowserAsync();
        }

        private async Task InitializeBrowserAsync()
        {
            try
            {
                _userDataFolder = Path.Combine(
                    Path.GetTempPath(),
                    "M365Tool",
                    "AuthWebView2",
                    Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(_userDataFolder);

                var options = new CoreWebView2EnvironmentOptions
                {
                    AllowSingleSignOnUsingOSPrimaryAccount = false
                };
                RecordLoadTiming("environment-start");
                CoreWebView2Environment env = await CoreWebView2Environment.CreateAsync(null, _userDataFolder, options);
                RecordLoadTiming("environment-ready");
                if (IsDisposed || Disposing) return;
                _ = _webView.Handle;
                await _webView.EnsureCoreWebView2Async(env);
                RecordLoadTiming("webview-ready");
                if (IsDisposed || Disposing) return;
                _webView.ZoomFactor = SignInPageZoomFactor;

                await ConfigureWebViewAsync();
                RecordLoadTiming("browser-prepared");
                if (IsDisposed || Disposing) return;
            }
            catch (Exception) when (IsDisposed || Disposing)
            {
                // Closing the owner can cancel native controller creation during warm-up.
            }
            finally
            {
                if (IsDisposed || Disposing) TryDeleteUserDataFolder();
            }
        }

        private async Task StartMandatorySignInFlowAsync()
        {
            if (_initialized)
            {
                return;
            }

            _initialized = true;
            _clickStartedAt = _loadClock.ElapsedMilliseconds;
            RecordLoadTiming(_browserInitializationTask?.IsCompletedSuccessfully == true ? "click-prepared" : "click-cold");
            _navigationStarted = false;
            _navigationCompleted = false;

            try
            {
                SetStatus(T("正在初始化登录页面...", "Initializing sign-in page..."));
                StartNavigationWatchdog(SignInPageLoadTimeoutMs);
                await PrepareAsync();
                if (IsDisposed || Disposing) return;

                SetStatus(T("正在打开登录页面...", "Opening sign-in page..."));
                RecordLoadTiming("navigation-requested");
                _webView.CoreWebView2.Navigate(MyAccountUrl);
                StartNavigationWatchdog(SignInPageLoadTimeoutMs);
                _verificationTimer.Start();
            }
            catch (WebView2RuntimeNotFoundException ex)
            {
                if (IsDisposed || Disposing) return;
                MessageBox.Show(
                    T(
                        "未找到 WebView2 Runtime，请安装或修复 Microsoft Edge WebView2 Runtime 后重试。\r\n\r\n详细信息：" + ex.Message,
                        "WebView2 Runtime was not found. Install/repair Microsoft Edge WebView2 Runtime and retry.\r\n\r\nDetails: " + ex.Message),
                    T("缺少运行时", "Runtime Missing"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                DialogResult = DialogResult.Cancel;
                Close();
            }
            catch (Exception ex)
            {
                if (IsDisposed || Disposing) return;
                MessageBox.Show(
                    T("无法打开登录页面：" + ex.Message, "Unable to open sign-in page: " + ex.Message),
                    T("打开失败", "Open Failed"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                DialogResult = DialogResult.Cancel;
                Close();
            }
        }

        private async Task ConfigureWebViewAsync()
        {
            CoreWebView2 core = _webView.CoreWebView2;
            core.Settings.IsScriptEnabled = true;
            core.Settings.IsStatusBarEnabled = false;
            _webView.ZoomFactor = SignInPageZoomFactor;
            core.NavigationStarting += (_, e) =>
            {
                _verificationNavigationId++;
                _signInSurfaceReady = false;
                _verificationTimer.Interval = 250;
                RecordLoadTiming("navigation-started");
                _navigationStarted = true;
                _navigationCompleted = false;
                _lastNavigationErrorStatus = null;
                Uri.TryCreate(e.Uri, UriKind.Absolute, out Uri? target);
                if (IsChinaSignInHost(target))
                {
                    _hasVisitedChinaSignIn = true;
                    RecordLoadTiming("login-host-reached");
                    _hasReturnedFromChinaSignIn = false;
                }
                else if (IsVerificationHost(target) && _hasVisitedChinaSignIn
                    && IsChinaSignInHost(_lastNavigationTarget ?? TryGetCurrentUri()))
                    _hasReturnedFromChinaSignIn = true;
                _lastNavigationTarget = target;
                // Keep the renderer active behind the opaque loading label so page scripts are not background-throttled.
                if (!IsVerified && !_isProfileSyncing) _webView.Visible = ShouldShowSignInContent(target);
                StartNavigationWatchdog(SignInPageLoadTimeoutMs);
                SetStatus(IsVerificationHost(target) && _hasReturnedFromChinaSignIn
                    ? T("正在确认登录结果...", "Confirming sign-in...")
                    : T("正在连接登录验证服务...", "Connecting to the sign-in service..."));
                if (_hasReturnedFromChinaSignIn && IsVerificationHost(target))
                    this.BeginInvokeWhenReady(() => _ = TryMarkVerifiedAsync(target));
            };
            core.SourceChanged += (_, _) =>
            {
                if (!IsVerified && !_isProfileSyncing && IsVerificationHost(TryGetCurrentUri()))
                {
                    _webView.Visible = false;
                    SetStatus(T("正在确认登录结果...", "Confirming sign-in..."));
                    _ = TryMarkVerifiedAsync(TryGetCurrentUri());
                }
            };
            core.ContentLoading += (_, _) =>
            {
                if (!_isProfileSyncing && !IsVerified && !_signInSurfaceReady)
                {
                    SetStatus(T("正在加载登录表单...", "Loading the sign-in form..."));
                }
            };
            core.NewWindowRequested += (_, e) =>
            {
                e.Handled = true;
                if (!string.IsNullOrWhiteSpace(e.Uri) &&
                    Uri.TryCreate(e.Uri, UriKind.Absolute, out Uri? uri) &&
                    IsHttpOrHttps(uri))
                {
                    _navigationStarted = true;
                    _navigationCompleted = false;
                    _lastNavigationErrorStatus = null;
                    if (IsVerificationHost(uri) && IsVerified)
                    {
                        _isProfileSyncing = true;
                        SetStatus(T(
                            "正在初始化中...",
                            "Initializing..."));
                    }
                    else
                    {
                        SetStatus(T("登录页面加载中...", "Loading sign-in page..."));
                    }

                    StartNavigationWatchdog(SignInPageLoadTimeoutMs);
                    core.Navigate(e.Uri);
                }
            };
            core.DOMContentLoaded += WebView_DOMContentLoaded;
            core.NavigationCompleted += WebView_NavigationCompleted;
            core.WebMessageReceived += SignInAccount_WebMessageReceived;
            await core.AddScriptToExecuteOnDocumentCreatedAsync(SignInAccountCaptureScript);
        }

        private void StartNavigationWatchdog(int interval)
        {
            _navigationWatchdogTimer.Stop();
            _navigationWatchdogTimer.Interval = interval;
            _navigationWatchdogTimer.Start();
        }

        private void SetStatus(string text)
        {
            if (_statusLabel.IsDisposed)
            {
                return;
            }

            _statusLabel.Text = text;
            _statusLabel.Visible = true;
            _statusLabel.BringToFront();
        }

        private void HideStatus()
        {
            if (_statusLabel.IsDisposed)
            {
                return;
            }

            if (!ShouldShowSignInContent(TryGetCurrentUri()))
            {
                _webView.Visible = false;
                return;
            }

            _statusLabel.Visible = false;
            _webView.Visible = true;
            _webView.BringToFront();
        }

        private async void NavigationWatchdogTimer_Tick(object? sender, EventArgs e)
        {
            _navigationWatchdogTimer.Stop();
            if (_signInSurfaceReady || IsVerified || _isProfileSyncing)
            {
                return;
            }

            RecordLoadTiming("load-timeout");
            FlushLoadTimings();
            string detail = _navigationStarted
                ? T("登录页已开始加载，但长时间没有完成。", "The sign-in page started loading but did not finish.")
                : T("WebView2 初始化或登录页加载没有开始。", "WebView2 initialization or sign-in page loading did not start.");

            MessageBox.Show(
                T(
                    $"登录页面暂时未加载出来。\r\n\r\n{detail}\r\n\r\n诊断信息：\r\n{await CollectWebViewDiagnosticsAsync()}",
                    $"The sign-in page did not load.\r\n\r\n{detail}\r\n\r\nDiagnostics:\r\n{await CollectWebViewDiagnosticsAsync()}"),
                T("登录页加载超时", "Sign-in Page Timeout"),
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        private async void WebView_NavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            int navigationId = _verificationNavigationId;
            RecordLoadTiming("navigation-completed");
            _navigationStarted = true;
            _navigationCompleted = e.IsSuccess;
            _lastNavigationErrorStatus = e.IsSuccess ? null : e.WebErrorStatus;
            Uri? uri = TryGetCurrentUri();
            if (IsVerificationHost(uri) && !_isProfileSyncing)
            {
                SetStatus(e.IsSuccess
                    ? T("正在确认登录结果...", "Confirming sign-in...")
                    : T("登录结果加载失败，请关闭登录窗口后重试。", "Unable to confirm sign-in. Close this window and retry."));
                if (e.IsSuccess) await TryMarkVerifiedAsync(uri);
                return;
            }
            if (_signInSurfaceReady && !_isProfileSyncing) return;
            if (e.IsSuccess)
            {
                if (_isProfileSyncing)
                {
                    SetStatus(T(
                        "正在初始化中...",
                        "Initializing..."));
                }
                else
                {
                    await TryRevealSignInSurfaceAsync();
                    if (IsDisposed || Disposing || _signInSurfaceReady || navigationId != _verificationNavigationId) return;
                    // Completed error/help pages can be displayed without mistaking them for authentication.
                    bool rendered = await HasRenderedSignInSurfaceAsync();
                    if (IsDisposed || Disposing || navigationId != _verificationNavigationId || uri != TryGetCurrentUri()) return;
                    if (rendered) { MarkSignInSurfaceReady(); return; }
                    SetStatus(T("正在等待登录表单...", "Waiting for the sign-in form..."));
                    StartNavigationWatchdog(SignInPageRenderTimeoutMs);
                }
            }
            if (!e.IsSuccess || _webView.CoreWebView2 == null)
            {
                if (!e.IsSuccess)
                {
                    SetStatus(T(
                        "登录页加载失败，请关闭当前登录窗口后重试。",
                        "The sign-in page failed to load. Close this sign-in window and retry."));
                }

                return;
            }

            await TryMarkVerifiedAsync(uri);
        }

        private async void WebView_DOMContentLoaded(object? sender, CoreWebView2DOMContentLoadedEventArgs e)
        {
            RecordLoadTiming("dom-content-loaded");
            await TryRevealSignInSurfaceAsync();
        }

        private async Task<bool> HasRenderedSignInSurfaceAsync()
        {
            string state = await TryEvalStringAsync(
                "(() => { try {" +
                "const body=document.body;if(!body)return '';" +
                "const clean=v=>String(v||'').replace(/\\s+/g,' ').trim();" +
                "const text=clean(body.innerText||body.textContent||'');" +
                "const html=String(body.innerHTML||'').trim();" +
                "const visible=el=>{const r=el.getBoundingClientRect();const s=getComputedStyle(el);return r.width>1&&r.height>1&&r.left<innerWidth&&r.top<innerHeight&&r.right>0&&r.bottom>0&&s.visibility!=='hidden'&&s.display!=='none';};" +
                "const interactive=Array.from(document.querySelectorAll('input,button,a,[role=\"button\"],[aria-label]')).filter(visible).length;" +
                "const inputs=document.querySelectorAll('input[type=\"email\"],input[type=\"password\"],#i0116,#i0118,#idSIButton9').length;" +
                "const authText=/(sign in|login|password|microsoft|21vianet|my account|登录|登入|密码|帐户|账户|验证)/i.test(text);" +
                "const hasRealSurface=authText||inputs>0||interactive>=2||text.length>40;" +
                "const shellOnly=html.length>0&&text.length===0&&interactive===0&&inputs===0;" +
                "return document.readyState!=='loading'&&hasRealSurface&&!shellOnly?'1':'';" +
                "} catch(e) { return ''; } })();");
            return state == "1";
        }

        private async Task TryRevealSignInSurfaceAsync()
        {
            Uri? uri = TryGetCurrentUri();
            if (_isCheckingSurface || _signInSurfaceReady || IsVerified || _isProfileSyncing
                || IsDisposed || Disposing || _webView.CoreWebView2 == null || !ShouldShowSignInContent(uri)) return;
            int navigationId = _verificationNavigationId;
            _isCheckingSurface = true;
            try
            {
                bool ready = await TryEvalStringAsync(SignInSurfaceReadyScript) == "1";
                if (ready && !IsDisposed && !Disposing && !IsVerified && !_isProfileSyncing
                    && navigationId == _verificationNavigationId && uri == TryGetCurrentUri())
                    MarkSignInSurfaceReady();
            }
            finally { _isCheckingSurface = false; }
        }

        private void MarkSignInSurfaceReady()
        {
            if (!ShouldShowSignInContent(TryGetCurrentUri()) || IsVerified || _isProfileSyncing) return;
            _signInSurfaceReady = true;
            _navigationWatchdogTimer.Stop();
            _verificationTimer.Interval = 1000;
            HideStatus();
            RecordLoadTiming("surface-ready");
            FlushLoadTimings();
        }

        private void RecordLoadTiming(string stage)
        {
            if (!_timedStages.Add(stage)) return;
            long elapsed = _loadClock.ElapsedMilliseconds;
            string sinceClick = _clickStartedAt < 0 ? "-" : (elapsed - _clickStartedAt).ToString();
            _loadTimings.Add($"{DateTime.UtcNow:O}\tSession={_loadTraceId}\tStage={stage}\tElapsedMs={elapsed}\tSinceClickMs={sinceClick}");
        }

        private void FlushLoadTimings()
        {
            if (_clickStartedAt < 0 || _flushedTimingCount == _loadTimings.Count) return;
            string[] lines = _loadTimings.Skip(_flushedTimingCount).ToArray();
            _flushedTimingCount = _loadTimings.Count;
            _ = Task.Run(() =>
            {
                try
                {
                    string directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "M365Tool", "Logs");
                    lock (LoadTraceLock)
                    {
                        Directory.CreateDirectory(directory);
                        File.AppendAllLines(Path.Combine(directory, "AuthLoadTiming.log"), lines);
                    }
                }
                catch { /* Diagnostics must not delay or interrupt sign-in. */ }
            });
        }

        private async Task<string> CollectWebViewDiagnosticsAsync()
        {
            try
            {
                string runtimeVersion = CoreWebView2Environment.GetAvailableBrowserVersionString();
                string source = _webView.CoreWebView2?.Source ?? string.Empty;
                string scriptEnabled = _webView.CoreWebView2?.Settings.IsScriptEnabled == true ? "true" : "false";
                string navigationError = _lastNavigationErrorStatus?.ToString() ?? "none";
                string documentState = await TryEvalStringAsync(
                    "(() => { try {" +
                    "const body=document.body;" +
                    "const visible=el=>{const r=el.getBoundingClientRect();const s=getComputedStyle(el);return r.width>1&&r.height>1&&r.left<innerWidth&&r.top<innerHeight&&r.right>0&&r.bottom>0&&s.visibility!=='hidden'&&s.display!=='none';};" +
                    "const text=body?String(body.innerText||body.textContent||'').trim():'';" +
                    "const html=body?String(body.innerHTML||'').trim():'';" +
                    "const interactive=Array.from(document.querySelectorAll('input,button,a,[role=\"button\"],[aria-label]')).filter(visible).length;" +
                    "return JSON.stringify({readyState:document.readyState,title:document.title,href:location.href,textLength:text.length,htmlLength:html.length,interactive,inputs:document.querySelectorAll('input').length,buttons:document.querySelectorAll('button').length,iframes:document.querySelectorAll('iframe').length});" +
                    "} catch(e) { return JSON.stringify({error:String(e&&e.message||e)}); } })();");

                return $"Runtime={runtimeVersion}\r\nSource={source}\r\nScriptEnabled={scriptEnabled}\r\nNavigationStarted={_navigationStarted}\r\nNavigationCompleted={_navigationCompleted}\r\nWebError={navigationError}\r\nUserDataFolder={_userDataFolder}\r\nDocument={documentState}";
            }
            catch (Exception ex)
            {
                return "Unable to collect diagnostics: " + ex.Message;
            }
        }

        private async Task TryMarkVerifiedAsync(Uri? uri)
        {
            if (_isCheckingVerification || !CanCheckSignInReturn(uri))
            {
                return;
            }

            int navigationId = _verificationNavigationId;
            _isCheckingVerification = true;
            try
            {
                bool authenticatedSession = await HasMicrosoftSignInSessionAsync();
                if (navigationId != _verificationNavigationId || !CanCheckSignInReturn(uri)) return;
                // Profile rendering is independent of authentication. Only use it as an
                // identity fallback when the sign-in page did not supply an email address.
                if (authenticatedSession && !IsValidAccount(_candidateAccount)
                    && CanReadPortalIdentity(TryGetCurrentUri(), _navigationCompleted, _lastNavigationErrorStatus.HasValue))
                    _candidateAccount = await ReadSignedInAccountAsync();
                if (navigationId != _verificationNavigationId || !CanCheckSignInReturn(uri)) return;
                TryCompleteCredentialVerification(uri, authenticatedSession);
            }
            catch (Exception) when (IsDisposed || Disposing)
            {
                // Closing the sign-in window may complete while its session check is awaiting.
            }
            finally
            {
                _isCheckingVerification = false;
            }
        }

        private bool CanCheckSignInReturn(Uri? uri) =>
            !IsDisposed && !Disposing && !IsVerified && !_isCompletingVerification
            && _webView.CoreWebView2 != null
            && _hasReturnedFromChinaSignIn && IsVerificationHost(uri);

        internal static bool CanReadPortalIdentity(Uri? uri, bool navigationCompleted, bool hasNavigationError) =>
            navigationCompleted && !hasNavigationError && IsVerificationHost(uri);

        internal static bool ShouldShowSignInContent(Uri? uri) =>
            uri is { IsAbsoluteUri: true } && uri.Scheme == Uri.UriSchemeHttps && !IsVerificationHost(uri);

        internal bool TryCompleteCredentialVerification(Uri? uri, bool authenticatedSession)
        {
            if (IsDisposed || Disposing || IsVerified || !_hasVisitedChinaSignIn || !_hasReturnedFromChinaSignIn
                || !authenticatedSession || !IsVerificationHost(uri)
                || !IsValidAccount(_candidateAccount)) return false;
            CompleteVerification(_candidateAccount.Trim(), ExtractUserNameFromAccount(_candidateAccount));
            return true;
        }

        private async Task<bool> HasMicrosoftSignInSessionAsync()
        {
            if (!_hasVisitedChinaSignIn || _webView.CoreWebView2 == null) return false;
            foreach (string url in new[] { "https://login.partner.microsoftonline.cn/", "https://login.chinacloudapi.cn/", "https://login.microsoftonline.cn/" })
            {
                var cookies = await _webView.CoreWebView2.CookieManager.GetCookiesAsync(url);
                foreach (var cookie in cookies)
                    if (cookie.IsHttpOnly && IsAuthenticationSessionCookie(cookie.Name)) return true;
            }
            return false;
        }

        internal static bool IsAuthenticationSessionCookie(string name) =>
            name.Equals("ESTSAUTH", StringComparison.OrdinalIgnoreCase)
            || name.Equals("ESTSAUTHPERSISTENT", StringComparison.OrdinalIgnoreCase)
            || name.Equals("estsStateTransient", StringComparison.OrdinalIgnoreCase)
            || name.Equals("estsStatePersistent", StringComparison.OrdinalIgnoreCase);

        internal static bool IsChinaSignInHost(Uri? uri) => uri is { IsAbsoluteUri: true }
            && uri.Scheme == Uri.UriSchemeHttps
            && (uri.Host.Equals("login.partner.microsoftonline.cn", StringComparison.OrdinalIgnoreCase)
                || uri.Host.Equals("login.chinacloudapi.cn", StringComparison.OrdinalIgnoreCase)
                || uri.Host.Equals("login.microsoftonline.cn", StringComparison.OrdinalIgnoreCase));

        private void SignInAccount_WebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            if (IsDisposed || Disposing || IsVerified || !Uri.TryCreate(e.Source, UriKind.Absolute, out Uri? source)
                || !IsChinaSignInHost(source)) return;
            try
            {
                using JsonDocument message = JsonDocument.Parse(e.WebMessageAsJson);
                JsonElement root = message.RootElement;
                if (root.GetProperty("kind").GetString() != "m365tool-signin-account") return;
                string account = root.GetProperty("account").GetString() ?? string.Empty;
                if (IsValidAccount(account)) _candidateAccount = account.Trim();
            }
            catch (JsonException) { }
            catch (KeyNotFoundException) { }
            catch (InvalidOperationException) { }
        }

        internal const string SignInSurfaceReadyScript =
            "(() => { try {" +
            "const visible=el=>{const r=el.getBoundingClientRect();const s=getComputedStyle(el);" +
            "return r.width>1&&r.height>1&&r.left<innerWidth&&r.top<innerHeight&&r.right>0&&r.bottom>0" +
            "&&s.display!=='none'&&s.visibility!=='hidden'&&Number(s.opacity)!==0;};" +
            "return Array.from(document.querySelectorAll('input:not([type=hidden]),button,[role=button]'))" +
            ".some(el=>visible(el)&&!el.disabled&&el.getAttribute('aria-disabled')!=='true')?'1':'';" +
            "}catch(e){return '';}})();";

        internal const string SignInAccountCaptureScript =
            "(() => {" +
            "if(location.protocol!=='https:'||!['login.partner.microsoftonline.cn','login.chinacloudapi.cn','login.microsoftonline.cn'].includes(location.hostname))return;" +
            "const emit=()=>{const input=document.querySelector('input[type=\"email\"],input[name=\"loginfmt\"],input[name=\"UserName\"],#i0116');" +
            "const identity=document.querySelector('#displayName,#userName,.identity');" +
            "const text=(input&&input.value)||identity&&(identity.innerText||identity.textContent)||'';" +
            "const match=String(text).match(/[A-Za-z0-9._%+\\-]+@[A-Za-z0-9.\\-]+\\.[A-Za-z]{2,}/);" +
            "if(match&&window.chrome&&window.chrome.webview)window.chrome.webview.postMessage({kind:'m365tool-signin-account',account:match[0]});};" +
            "for(const name of ['input','change','click','submit','DOMContentLoaded'])document.addEventListener(name,emit,true);emit();})();";

        private Task<string> ReadSignedInAccountAsync() => TryEvalStringAsync(SignedInAccountScript);

        internal const string SignedInAccountScript =
            "(() => { try {" +
            "if(document.readyState!=='complete'||location.protocol!=='https:'||location.hostname!=='myaccount.windowsazure.cn')return '';" +
            "const visible=el=>{const r=el.getBoundingClientRect();const s=getComputedStyle(el);return r.width>0&&r.height>0&&s.visibility!=='hidden'&&s.display!=='none';};" +
            "const clean=t=>String(t||'').replace(/\\s+/g,' ').trim();" +
            "const email=el=>{if(!el||!visible(el))return '';const t=(el.innerText||el.textContent||'')+' '+(el.getAttribute('href')||'');const m=t.match(/[A-Za-z0-9._%+\\-]+@[A-Za-z0-9.\\-]+\\.[A-Za-z]{2,}/);return m?m[0]:'';};" +
            "const photo=document.querySelector('#profilePhoto');" +
            "if(photo&&visible(photo)){" +
            "const card=photo.closest('[class*=\"ms-Card\"]')||photo.closest('article,section')||photo.parentElement;" +
            "if(card&&visible(card)){" +
            "const signedInActions=Array.from(card.querySelectorAll('a,button,[role=\"button\"]')).some(el=>visible(el)&&/^(从所有位置注销|从所有位置退出|Sign out everywhere|Sign out of all sessions)$/i.test(clean(el.innerText||el.textContent)));" +
            "const account=email(card.querySelector('a[href^=\"mailto:\" i]'))||email(card);" +
            "if(signedInActions&&account)return account;}}" +
            "const logout=Array.from(document.querySelectorAll('#mectrl_body_signOut,a[href*=\"/logout\" i],a[href*=\"signout\" i],[aria-label=\"Sign out\"],[aria-label=\"注销\"],[aria-label=\"退出登录\"]')).find(visible);" +
            "if(!logout||!visible(logout)){" +
            "const trigger=document.querySelector('#mectrl_main_trigger,[aria-label*=\"Account manager\" i],#mectrl_headerPicture');" +
            "if(trigger&&visible(trigger)&&trigger.getAttribute('aria-expanded')!=='true')trigger.click();return '';}" +
            "for(const el of document.querySelectorAll('#mectrl_currentAccount_secondary,#mectrl_currentAccount_primary')){" +
            "const account=email(el);if(account)return account;}" +
            "return ''; } catch(e) { return ''; } })();";

        private void CompleteVerification(string account, string userName)
        {
            RecordLoadTiming("authenticated");
            FlushLoadTimings();
            VerifiedAccount = account;
            VerifiedUserName = userName;
            VerifiedAvatarBytes = null;
            HasVerifiedProfileDisplayName = false;
            _isCompletingVerification = true;
            _verificationTimer.Stop();
            _navigationWatchdogTimer.Stop();
            SetStatus(T("验证成功，正在进入工具...", "Verification succeeded. Opening the tool..."));
            IsVerified = true;
            DialogResult = DialogResult.OK;
        }

        public bool BeginDeferredProfileSync(
            IWin32Window owner,
            Action<string, string, byte[]?, bool> onCompleted)
        {
            if (_deferredProfileSyncStarted ||
                IsDisposed ||
                !IsVerified ||
                !IsValidAccount(VerifiedAccount) ||
                _webView.CoreWebView2 == null)
            {
                return false;
            }

            _deferredProfileSyncStarted = true;
            _isProfileSyncing = true;
            _ = RunDeferredProfileSyncAsync(owner, onCompleted);
            return true;
        }

        private async Task RunDeferredProfileSyncAsync(
            IWin32Window owner,
            Action<string, string, byte[]?, bool> onCompleted)
        {
            try
            {
                PrepareDeferredProfileSyncWindow(owner);
                byte[]? avatarBytes = await TryFetchMyAccountProfileAsync(VerifiedAccount);
                onCompleted(VerifiedAccount, VerifiedUserName, avatarBytes, HasVerifiedProfileDisplayName);
            }
            catch (Exception ex)
            {
                _lastProfileTraceReason = "profile-exception-" + ex.GetType().Name;
                await WriteProfileTraceAsync(VerifiedAccount, "myaccount-profile-exception", 0);
            }
            finally
            {
                TryCloseDeferredProfileSyncWindow();
            }
        }

        private void PrepareDeferredProfileSyncWindow(IWin32Window owner)
        {
            StartPosition = FormStartPosition.Manual;
            Location = new Point(-32000, -32000);
            Opacity = 0.01d;
            ShowInTaskbar = false;
            _webView.Visible = true;
            if (!Visible)
            {
                Show(owner);
            }
        }

        private void TryCloseDeferredProfileSyncWindow()
        {
            try
            {
                if (!IsDisposed)
                {
                    Close();
                    Dispose();
                }
            }
            catch
            {
                // Ignore cleanup failures.
            }
        }

        private async Task<byte[]?> TryFetchMyAccountProfileAsync(string account)
        {
            if (!IsValidAccount(account) || _webView.CoreWebView2 == null)
            {
                _lastProfileTraceReason = "invalid-account-or-webview";
                await WriteProfileTraceAsync(account, "myaccount-profile-skipped", 0);
                return null;
            }

            _isProfileSyncing = true;
            SetStatus(T(
                "正在初始化中...",
                "Initializing..."));

            _lastAvatarSource = string.Empty;
            _lastAvatarRect = string.Empty;
            _lastProfileCandidates = string.Empty;
            _lastProfileTraceReason = string.Empty;

            // Credential handoff can happen before the returning portal navigation has completed.
            // Let that navigation render instead of canceling it with an identical Navigate call.
            if (!IsVerificationHost(TryGetCurrentUri()) && !IsVerificationHost(_lastNavigationTarget))
                await TryNavigateAndWaitAsync(MyAccountUrl, 15000);

            bool profileReady = await WaitForMyAccountProfileShellAsync();
            if (!profileReady)
            {
                // A closed modal window or a canceled return may need one explicit retry.
                await TryNavigateAndWaitAsync(MyAccountUrl, 15000);
                profileReady = await WaitForMyAccountProfileShellAsync();
            }
            if (!profileReady)
            {
                _isProfileSyncing = false;
                _lastProfileTraceReason = "profile-shell-timeout";
                await WriteProfileTraceAsync(account, "myaccount-profile-shell-timeout", 0);
                return null;
            }

            byte[]? avatarBytes = null;
            DateTime deadline = DateTime.UtcNow.AddSeconds(10);
            while (DateTime.UtcNow < deadline)
            {
                await TryOpenMyAccountProfileMenuAsync();

                MyAccountProfileResult profile = await TryExtractMyAccountProfileAsync(account);
                string displayName = !string.IsNullOrWhiteSpace(profile.DisplayName)
                    ? profile.DisplayName
                    : await TryExtractMyAccountDisplayNameAsync(account);
                if (!string.IsNullOrWhiteSpace(displayName))
                {
                    VerifiedUserName = displayName;
                    HasVerifiedProfileDisplayName = HasSyncedProfileDisplayName(account);
                }

                _lastAvatarSource = profile.AvatarSource;
                _lastAvatarRect = profile.AvatarRect?.ToTraceString() ?? string.Empty;
                _lastProfileCandidates = profile.Candidates;

                if (profile.AvatarRect is not null)
                {
                    avatarBytes = await TryCaptureMyAccountAvatarAsync(profile.AvatarRect);
                    if (avatarBytes is null)
                    {
                        _lastProfileTraceReason = "avatar-capture-empty";
                    }
                }
                else
                {
                    _lastProfileTraceReason = "avatar-rect-missing";
                }

                if (HasSyncedProfileDisplayName(account) && avatarBytes is not null)
                {
                    _isProfileSyncing = false;
                    HasVerifiedProfileDisplayName = true;
                    _lastProfileTraceReason = "ok";
                    await WriteProfileTraceAsync(account, "myaccount-success", avatarBytes?.Length ?? 0);
                    return avatarBytes;
                }

                await Task.Delay(300);
            }

            _isProfileSyncing = false;
            if (string.IsNullOrWhiteSpace(_lastProfileTraceReason))
            {
                _lastProfileTraceReason = avatarBytes is null ? "profile-timeout-no-avatar" : "profile-timeout-display-not-synced";
            }
            await WriteProfileTraceAsync(account, "myaccount-timeout", avatarBytes?.Length ?? 0);
            return avatarBytes;
        }

        private async Task<MyAccountProfileResult> TryExtractMyAccountProfileAsync(string account)
        {
            if (_webView.CoreWebView2 == null)
            {
                return MyAccountProfileResult.Empty;
            }

            string accountJson = JsonSerializer.Serialize(account.Trim());
            string raw = await TryEvalStringAsync(
                "(() => { try {" +
                $"const expectedAccount={accountJson}.toLowerCase();" +
                "const clean=v=>String(v||'').replace(/\\s+/g,' ').trim();" +
                "const email=/[A-Za-z0-9._%+\\-]+@[A-Za-z0-9.\\-]+\\.[A-Za-z]{2,}/;" +
                "const rect=el=>{const r=el.getBoundingClientRect();return {left:Math.round(r.left),top:Math.round(r.top),width:Math.round(r.width),height:Math.round(r.height)};};" +
                "const visible=el=>{const r=el.getBoundingClientRect();return r.width>=16&&r.height>=16&&r.left<innerWidth&&r.top<innerHeight&&r.right>0&&r.bottom>0;};" +
                "const validName=t=>{t=clean(t);return t.length>=2&&t.length<=80&&!email.test(t)&&/[A-Za-z\\u4e00-\\u9fff]{2,}/.test(t)&&!/^(My Account|Security info|Devices|Change password|Organizations|Settings & Privacy|Recent activity|My Apps|Why can't I edit|Sign out everywhere|View account|My Microsoft 365 profile|Third party notice)$/i.test(t);};" +
                "const describe=(selector,source)=>{const el=document.querySelector(selector);if(!el)return {selector,source,found:false};const r=rect(el);const styles=getComputedStyle(el);const image=String(el.currentSrc||el.src||styles.backgroundImage||'');return {selector,source,found:true,visible:visible(el)&&(el.tagName!=='IMG'||(el.complete&&el.naturalWidth>=16)),tag:el.tagName,id:el.id||'',alt:el.getAttribute('alt')||'',role:el.getAttribute('role')||'',cls:String(el.className||'').slice(0,120),srcLen:image.length,rect:r};};" +
                "const candidates=[" +
                "describe('#mectrl_currentAccount_picture .mectrl_profilepic','mectrl_currentAccount_picture')," +
                "describe('#mectrl_currentAccount_picture [role=\"img\"]','mectrl_currentAccount_picture')," +
                "describe('#mectrl_headerPicture','mectrl_headerPicture')," +
                "describe('#profilePhoto img','profilePhoto')," +
                "describe('#profilePhoto canvas','profilePhoto')," +
                "describe('#profilePhoto [role=\"img\"]','profilePhoto')" +
                "];" +
                "const avatarCandidates=candidates.filter(x=>x.found&&x.visible);" +
                "let avatarSource='';let avatarRect=null;" +
                "if(avatarCandidates.length){avatarSource=avatarCandidates[0].source;avatarRect=avatarCandidates[0].rect;}" +
                "const primary=document.querySelector('#mectrl_currentAccount_primary');" +
                "let displayName=clean(primary&&(primary.innerText||primary.textContent)||'');" +
                "if(!validName(displayName))displayName='';" +
                "const triggerEl=document.querySelector('#mectrl_main_trigger');" +
                "const trigger=clean(triggerEl&&triggerEl.getAttribute('aria-label')||'');" +
                "if(!displayName){const match=/Account manager for\\s+(.+)$/i.exec(trigger);if(match&&validName(match[1]))displayName=clean(match[1]);}" +
                "const secondaryEl=document.querySelector('#mectrl_currentAccount_secondary');" +
                "const secondary=clean(secondaryEl&&(secondaryEl.innerText||secondaryEl.textContent)||'');" +
                "const diagnostics={primary:clean(primary&&(primary.innerText||primary.textContent)||''),secondary,trigger,hasCurrentPicture:!!document.querySelector('#mectrl_currentAccount_picture'),url:location.href,title:document.title};" +
                "return JSON.stringify({account:secondary||expectedAccount,displayName,avatarSource,avatarRect,candidates,diagnostics});" +
                "} catch(e) { return JSON.stringify({error:String(e&&e.message||e),stack:String(e&&e.stack||''),url:location.href,title:document.title}); } })();");

            return MyAccountProfileResult.FromJson(raw);
        }

        private async Task<string> TryExtractMyAccountDisplayNameAsync(string account)
        {
            if (_webView.CoreWebView2 == null)
            {
                return string.Empty;
            }

            string accountJson = JsonSerializer.Serialize(account.Trim());
            string candidate = await TryEvalStringAsync(
                "(() => { try {" +
                $"const account={accountJson}.toLowerCase();" +
                "const clean=v=>String(v||'').replace(/\\s+/g,' ').trim();" +
                "const email=/[A-Za-z0-9._%+\\-]+@[A-Za-z0-9.\\-]+\\.[A-Za-z]{2,}/;" +
                "const blocked=/^(profile overview collapsible item|change your profile photo|skipToMainContent|security info|devices|change password|organizations|privacy|sign out|account manager)$/i;" +
                "const valid=t=>t.length>=2&&t.length<=80&&!email.test(t)&&!blocked.test(t)&&/[A-Za-z\\u4e00-\\u9fff]{2,}/.test(t);" +
                "const firstValid=list=>{for(const item of list){const text=clean(item);if(valid(text))return text;}return '';};" +
                "const meControl=document.querySelector('#mectrl_currentAccount_primary');" +
                "const meName=clean(meControl&&(meControl.getAttribute('aria-label')||meControl.innerText||meControl.textContent)||'');" +
                "if(valid(meName))return meName;" +
                "const profile=document.querySelector('#profilePhoto');" +
                "if(profile){" +
                "const card=profile.closest('[class*=\"ms-Card\"], [class*=\"ms-Collapsible\"], [role=\"button\"]')||profile.parentElement;" +
                "if(card){const names=Array.from(card.querySelectorAll('[class*=\"ms-tileTitle\"], [class*=\"ms-TitleBar--label\"], div, span')).map(el=>clean(el.innerText||el.textContent||el.getAttribute('aria-label')||''));const picked=firstValid(names);if(picked)return picked;}" +
                "}" +
                "const explicit=firstValid(Array.from(document.querySelectorAll('[class*=\"ms-tileTitle\"], [class*=\"ms-TitleBar--label\"]')).map(el=>el.innerText||el.textContent||el.getAttribute('aria-label')||''));" +
                "if(explicit)return explicit;" +
                "const trigger=document.querySelector('#mectrl_main_trigger,[aria-label*=\"Account manager\" i]');" +
                "const triggerText=clean(trigger&&(trigger.getAttribute('aria-label')||trigger.innerText||trigger.textContent)||'');" +
                "const match=/Account manager for\\s+(.+)$/i.exec(triggerText);" +
                "if(match&&valid(match[1]))return clean(match[1]);" +
                "const body=clean(document.body&&document.body.innerText||'');" +
                "const bodyMatch=/Account manager for\\s+([^\\n\\r]+?)(?:\\s+(?:Security info|Devices|Change password)|$)/i.exec(body);" +
                "if(bodyMatch&&valid(bodyMatch[1]))return clean(bodyMatch[1]);" +
                "return ''; } catch(e) { return ''; } })();");

            candidate = CleanNameCandidate(candidate);
            if (string.IsNullOrWhiteSpace(candidate) ||
                candidate.Contains('@', StringComparison.Ordinal))
            {
                return string.Empty;
            }

            return candidate;
        }

        private async Task TryOpenMyAccountProfileMenuAsync()
        {
            await TryEvalStringAsync(
                "(() => { try {" +
                "const visible=el=>{const r=el.getBoundingClientRect();return r.width>=16&&r.height>=16&&r.left<innerWidth&&r.top<innerHeight&&r.right>0&&r.bottom>0;};" +
                "const picture=document.querySelector('#mectrl_currentAccount_picture .mectrl_profilepic,#mectrl_currentAccount_picture [role=\"img\"]');" +
                "if(picture&&visible(picture))return '1';" +
                "const trigger=document.querySelector('#mectrl_main_trigger,[aria-label*=\"Account manager\" i],#mectrl_headerPicture');" +
                "if(trigger&&typeof trigger.click==='function'){trigger.click();return '1';}" +
                "const profile=document.querySelector('[role=\"button\"][aria-label*=\"Profile overview\" i], .ms-CollapsibleHeader[role=\"button\"]');" +
                "if(profile&&typeof profile.click==='function'){profile.click();return '1';}" +
                "return ''; } catch(e) { return ''; } })();");
        }

        private async Task<byte[]?> TryCaptureMyAccountAvatarAsync(AvatarRect? avatarRect)
        {
            Rectangle? bounds = avatarRect?.ToRectangle();
            if (bounds is null || _webView.CoreWebView2 == null || _webView.ClientSize.Width <= 0 || _webView.ClientSize.Height <= 0)
            {
                return null;
            }

            try
            {
                using var stream = new MemoryStream();
                await _webView.CoreWebView2.CapturePreviewAsync(CoreWebView2CapturePreviewImageFormat.Png, stream);
                stream.Position = 0;
                using var screenshot = new Bitmap(stream);
                float scaleX = screenshot.Width / (float)_webView.ClientSize.Width * (float)_webView.ZoomFactor;
                float scaleY = screenshot.Height / (float)_webView.ClientSize.Height * (float)_webView.ZoomFactor;
                Rectangle crop = Rectangle.Intersect(
                    new Rectangle(
                        (int)Math.Floor(bounds.Value.Left * scaleX),
                        (int)Math.Floor(bounds.Value.Top * scaleY),
                        (int)Math.Ceiling(bounds.Value.Width * scaleX),
                        (int)Math.Ceiling(bounds.Value.Height * scaleY)),
                    new Rectangle(Point.Empty, screenshot.Size));

                if (crop.Width < 16 || crop.Height < 16)
                {
                    return null;
                }

                using Bitmap avatar = screenshot.Clone(crop, screenshot.PixelFormat);
                using var output = new MemoryStream();
                avatar.Save(output, ImageFormat.Png);
                return TryNormalizeAvatarImageBytes(output.ToArray());
            }
            catch
            {
                return null;
            }
        }

        private async Task WriteProfileTraceAsync(string account, string result, int avatarLength)
        {
            try
            {
                string directory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "M365Tool",
                    "Logs");
                Directory.CreateDirectory(directory);
                string candidates = !string.IsNullOrWhiteSpace(_lastProfileCandidates)
                    ? _lastProfileCandidates
                    : await TryCollectMyAccountProfileCandidatesAsync();
                string source = _webView.CoreWebView2?.Source ?? string.Empty;
                string line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss}\t{result}\tAccount={account}\tDisplay={VerifiedUserName}\tProfileDisplay={HasVerifiedProfileDisplayName}\tAvatarBytes={avatarLength}\tAvatarSource={_lastAvatarSource}\tAvatarRect={_lastAvatarRect}\tReason={_lastProfileTraceReason}\tUrl={source}\tCandidates={candidates}{Environment.NewLine}";
                File.AppendAllText(Path.Combine(directory, "AuthProfileTrace.log"), line);
            }
            catch
            {
                // Profile trace is best-effort only.
            }
        }

        private async Task<string> TryCollectMyAccountProfileCandidatesAsync()
        {
            return await TryEvalStringAsync(
                "(() => { try {" +
                "const clean=v=>String(v||'').replace(/\\s+/g,' ').trim();" +
                "const visible=el=>{const r=el.getBoundingClientRect();return r.width>1&&r.height>1&&r.left<innerWidth&&r.top<innerHeight&&r.right>0&&r.bottom>0;};" +
                "const texts=Array.from(document.querySelectorAll('#mectrl_currentAccount_primary,#mectrl_currentAccount_secondary,#mectrl_main_trigger,#profilePhoto,[aria-label],button,div,span')).map(el=>{const r=el.getBoundingClientRect();const role=(el.getAttribute('role')||'');const text=clean(el.getAttribute('aria-label')||el.innerText||el.textContent||'');return {id:el.id||'',text,role,tag:el.tagName,left:Math.round(r.left),top:Math.round(r.top),width:Math.round(r.width),height:Math.round(r.height)};})" +
                ".filter(x=>x.text&&x.top<innerHeight).slice(0,40);" +
                "const media=Array.from(document.querySelectorAll('#mectrl_currentAccount_picture,#profilePhoto,img,canvas,[style*=\"background\"]')).map(el=>{const r=el.getBoundingClientRect();const src=String(el.currentSrc||el.src||getComputedStyle(el).backgroundImage||'');return {id:el.id||'',tag:el.tagName,alt:el.alt||'',cls:String(el.className||''),left:Math.round(r.left),top:Math.round(r.top),width:Math.round(r.width),height:Math.round(r.height),data:src.includes('data:image/'),len:src.length};})" +
                ".filter(x=>x.width>0&&x.height>0).slice(0,20);" +
                "return JSON.stringify({title:document.title,href:location.href,texts,media});" +
                "} catch(e) { return String(e); } })();");
        }

        private bool HasSyncedProfileDisplayName(string account)
        {
            if (string.IsNullOrWhiteSpace(VerifiedUserName))
            {
                return false;
            }

            string accountPrefix = ExtractUserNameFromAccount(account);
            return !string.Equals(VerifiedUserName.Trim(), accountPrefix, StringComparison.OrdinalIgnoreCase);
        }

        private async Task<bool> WaitForMyAccountProfileShellAsync()
        {
            DateTime deadline = DateTime.UtcNow.AddSeconds(20);
            while (DateTime.UtcNow < deadline)
            {
                if (!IsVerificationHost(TryGetCurrentUri()))
                {
                    await Task.Delay(300);
                    continue;
                }
                string ready = await TryEvalStringAsync(
                    "(() => { try {" +
                    "const hasProfile=!!document.querySelector('#mectrl_currentAccount_primary,#mectrl_currentAccount_picture,#mectrl_headerPicture,#profilePhoto img,#profilePhoto canvas,#profilePhoto [role=\"img\"]');" +
                    "const body=(document.body&&(document.body.innerText||document.body.textContent)||'').trim();" +
                    "return hasProfile || body.length > 80 ? '1' : ''; " +
                    "} catch(e) { return ''; } })();");
                if (ready == "1")
                {
                    return true;
                }

                await Task.Delay(300);
            }

            return false;
        }

        private async Task<bool> TryNavigateAndWaitAsync(string url, int timeoutMs)
        {
            if (_webView.CoreWebView2 == null)
            {
                return false;
            }

            CoreWebView2 core = _webView.CoreWebView2;
            var completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            ulong? requestedNavigationId = null;
            void Starting(object? sender, CoreWebView2NavigationStartingEventArgs e)
            {
                if (requestedNavigationId == null && string.Equals(e.Uri, url, StringComparison.OrdinalIgnoreCase))
                    requestedNavigationId = e.NavigationId;
            }
            void Completed(object? sender, CoreWebView2NavigationCompletedEventArgs e)
            {
                if (requestedNavigationId == e.NavigationId) completion.TrySetResult(e.IsSuccess);
            }

            core.NavigationStarting += Starting;
            core.NavigationCompleted += Completed;
            try
            {
                using var cancellation = new CancellationTokenSource(timeoutMs);
                using CancellationTokenRegistration registration = cancellation.Token.Register(() => completion.TrySetResult(false));
                core.Navigate(url);
                return await completion.Task;
            }
            catch
            {
                return false;
            }
            finally
            {
                try
                {
                    core.NavigationStarting -= Starting;
                    core.NavigationCompleted -= Completed;
                }
                catch (Exception) when (IsDisposed || Disposing) { }
            }
        }

        private static byte[]? TryNormalizeAvatarImageBytes(byte[] raw)
        {
            if (raw.Length < 128 || raw.Length > 2 * 1024 * 1024)
            {
                return null;
            }

            try
            {
                using var input = new MemoryStream(raw);
                using var image = Image.FromStream(input);
                if (image.Width < 16 || image.Height < 16)
                {
                    return null;
                }

                using var normalized = new Bitmap(image);
                using var output = new MemoryStream();
                normalized.Save(output, ImageFormat.Png);
                byte[] bytes = output.ToArray();
                return bytes.Length >= 128 ? bytes : null;
            }
            catch
            {
                return null;
            }
        }

        private async Task<string> TryEvalStringAsync(string script)
        {
            if (_webView.CoreWebView2 == null)
            {
                return string.Empty;
            }

            try
            {
                string raw = await _webView.CoreWebView2.ExecuteScriptAsync(script);
                if (string.IsNullOrWhiteSpace(raw))
                {
                    return string.Empty;
                }

                using JsonDocument json = JsonDocument.Parse(raw);
                if (json.RootElement.ValueKind == JsonValueKind.String)
                {
                    return json.RootElement.GetString() ?? string.Empty;
                }

                return json.RootElement.ToString();
            }
            catch
            {
                return string.Empty;
            }
        }

        private Uri? TryGetCurrentUri()
        {
            string source = _webView.CoreWebView2?.Source ?? string.Empty;
            return Uri.TryCreate(source, UriKind.Absolute, out Uri? uri) ? uri : null;
        }

        private static bool IsHttpOrHttps(Uri uri)
        {
            return uri.Scheme.Equals(Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) ||
                   uri.Scheme.Equals(Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsVerificationHost(Uri? uri)
        {
            if (uri == null || !uri.IsAbsoluteUri || uri.Scheme != Uri.UriSchemeHttps)
            {
                return false;
            }

            string host = uri.Host.ToLowerInvariant();
            if (host.StartsWith("login.", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return host.Equals("myaccount.windowsazure.cn", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsValidAccount(string? account)
        {
            return !string.IsNullOrWhiteSpace(account) &&
                   Regex.IsMatch(account.Trim(), @"^[^@\s]+@[^@\s]+\.[^@\s]+$", RegexOptions.CultureInvariant);
        }

        private static string ExtractUserNameFromAccount(string account)
        {
            if (string.IsNullOrWhiteSpace(account))
            {
                return string.Empty;
            }

            string trimmed = account.Trim();
            int at = trimmed.IndexOf('@');
            return at > 0 ? trimmed[..at] : trimmed;
        }

        private static string CleanNameCandidate(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            string candidate = value
                .Replace("\t", " ")
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Trim();

            candidate = Regex.Replace(candidate, @"[\uE000-\uF8FF]+", " ", RegexOptions.CultureInvariant);
            candidate = Regex.Replace(candidate, @"^Account manager for\s+", string.Empty, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            candidate = Regex.Replace(candidate, @"\s+", " ", RegexOptions.CultureInvariant).Trim();

            while (candidate.Contains("  ", StringComparison.Ordinal))
            {
                candidate = candidate.Replace("  ", " ", StringComparison.Ordinal);
            }

            if (candidate.Length < 2 || candidate.Length > 80)
            {
                return string.Empty;
            }

            if (Regex.IsMatch(candidate, @"^[\{\}\[\]\(\),.:;\-_=+]+$", RegexOptions.CultureInvariant))
            {
                return string.Empty;
            }

            if (BlockedNameRegex.IsMatch(candidate))
            {
                return string.Empty;
            }

            if (NumericLikeRegex.IsMatch(candidate))
            {
                return string.Empty;
            }

            return candidate;
        }

        private async void VerificationTimer_Tick(object? sender, EventArgs e)
        {
            if (IsVerified)
            {
                _verificationTimer.Stop();
                return;
            }

            await TryRevealSignInSurfaceAsync();
            Uri? uri = TryGetCurrentUri();
            await TryMarkVerifiedAsync(uri);
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _verificationTimer.Stop();
            _navigationWatchdogTimer.Stop();
            RecordLoadTiming("dialog-closed");
            FlushLoadTimings();

            // ShowDialog may close before the same WebView is reused for background
            // profile sync. Release its session only when the form is disposed.
            base.OnFormClosed(e);
        }

        protected override void Dispose(bool disposing)
        {
            bool releaseSession = disposing && !_resourcesDisposed;
            if (releaseSession)
            {
                _resourcesDisposed = true;
                _verificationTimer?.Stop();
                _navigationWatchdogTimer?.Stop();
                _verificationTimer?.Dispose();
                _navigationWatchdogTimer?.Dispose();
                try
                {
                    if (_webView?.CoreWebView2 != null)
                    {
                        _webView.CoreWebView2.DOMContentLoaded -= WebView_DOMContentLoaded;
                        _webView.CoreWebView2.NavigationCompleted -= WebView_NavigationCompleted;
                        _webView.CoreWebView2.WebMessageReceived -= SignInAccount_WebMessageReceived;
                    }
                }
                catch { /* The browser may already be shutting down. */ }
            }
            base.Dispose(disposing);
            if (releaseSession) TryDeleteUserDataFolder();
        }

        private void TryDeleteUserDataFolder()
        {
            if (string.IsNullOrWhiteSpace(_userDataFolder))
            {
                return;
            }

            string target = _userDataFolder;
            Task.Run(async () =>
            {
                for (int i = 0; i < 5; i++)
                {
                    try
                    {
                        if (Directory.Exists(target))
                        {
                            Directory.Delete(target, true);
                        }
                        return;
                    }
                    catch
                    {
                        await Task.Delay(200);
                    }
                }
            });
        }

        private sealed class MyAccountProfileResult
        {
            public static MyAccountProfileResult Empty { get; } = new();

            public string Account { get; private init; } = string.Empty;

            public string DisplayName { get; private init; } = string.Empty;

            public string AvatarSource { get; private init; } = string.Empty;

            public AvatarRect? AvatarRect { get; private init; }

            public string Candidates { get; private init; } = string.Empty;

            public static MyAccountProfileResult FromJson(string json)
            {
                if (string.IsNullOrWhiteSpace(json))
                {
                    return Empty;
                }

                try
                {
                    using JsonDocument document = JsonDocument.Parse(json);
                    JsonElement root = document.RootElement;
                    if (root.TryGetProperty("error", out JsonElement error) &&
                        error.ValueKind == JsonValueKind.String)
                    {
                        return new MyAccountProfileResult
                        {
                            Candidates = root.GetRawText()
                        };
                    }

                    return new MyAccountProfileResult
                    {
                        Account = GetString(root, "account"),
                        DisplayName = CleanNameCandidate(GetString(root, "displayName")),
                        AvatarSource = GetString(root, "avatarSource"),
                        AvatarRect = AvatarRect.FromJson(root, "avatarRect"),
                        Candidates = root.TryGetProperty("candidates", out JsonElement candidates)
                            ? candidates.GetRawText()
                            : string.Empty
                    };
                }
                catch
                {
                    return Empty;
                }
            }

            private static string GetString(JsonElement root, string name)
            {
                return root.TryGetProperty(name, out JsonElement value) &&
                       value.ValueKind == JsonValueKind.String
                    ? value.GetString() ?? string.Empty
                    : string.Empty;
            }
        }

        private sealed class AvatarRect
        {
            [JsonPropertyName("left")]
            public int Left { get; init; }

            [JsonPropertyName("top")]
            public int Top { get; init; }

            [JsonPropertyName("width")]
            public int Width { get; init; }

            [JsonPropertyName("height")]
            public int Height { get; init; }

            public Rectangle ToRectangle()
            {
                return new Rectangle(Left, Top, Width, Height);
            }

            public string ToTraceString()
            {
                return $"{Left},{Top},{Width},{Height}";
            }

            public static AvatarRect? FromJson(JsonElement root, string name)
            {
                if (!root.TryGetProperty(name, out JsonElement value) ||
                    value.ValueKind != JsonValueKind.Object)
                {
                    return null;
                }

                int left = GetInt(value, "left");
                int top = GetInt(value, "top");
                int width = GetInt(value, "width");
                int height = GetInt(value, "height");
                return width >= 16 && height >= 16
                    ? new AvatarRect { Left = left, Top = top, Width = width, Height = height }
                    : null;
            }

            private static int GetInt(JsonElement root, string name)
            {
                return root.TryGetProperty(name, out JsonElement value) &&
                       value.ValueKind == JsonValueKind.Number &&
                       value.TryGetInt32(out int result)
                    ? result
                    : 0;
            }
        }

        private string T(string zh, string en) => LocalizationService.Localize(_language, zh, en);
    }
}
