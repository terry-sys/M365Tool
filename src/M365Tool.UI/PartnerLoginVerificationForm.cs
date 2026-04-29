using System;
using System.IO;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Office365CleanupTool.Services;

namespace Office365CleanupTool
{
    public sealed class PartnerLoginVerificationForm : Form
    {
        private const string PartnerPortalUrl = "https://portal.partner.microsoftonline.cn/Home";
        private const int MaxVerificationPollCount = 45;

        private static readonly Regex EmailRegex = new(
            @"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);
        private static readonly Regex NumericLikeRegex = new(
            @"^\d+(?:[.,]\d+)?$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private readonly UiLanguage _language;
        private readonly WebView2 _webView;
        private readonly System.Windows.Forms.Timer _verificationTimer;
        private string _userDataFolder = string.Empty;
        private bool _initialized;
        private string _candidateAccount = string.Empty;
        private int _verificationPollCount;

        public bool IsVerified { get; private set; }

        public string VerifiedAccount { get; private set; } = string.Empty;

        public string VerifiedUserName { get; private set; } = string.Empty;

        public PartnerLoginVerificationForm(UiLanguage language)
        {
            _language = language;
            Text = "21V Sign-in Verification";
            Icon = AppIconProvider.GetAppIcon() ?? Icon;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            ShowInTaskbar = false;
            MinimizeBox = false;
            MaximizeBox = false;
            MinimumSize = new System.Drawing.Size(880, 640);
            Size = new System.Drawing.Size(1160, 820);
            BackColor = System.Drawing.Color.White;

            _webView = new WebView2
            {
                Dock = DockStyle.Fill
            };

            _verificationTimer = new System.Windows.Forms.Timer
            {
                Interval = 1000
            };
            _verificationTimer.Tick += VerificationTimer_Tick;

            Controls.Add(_webView);
            Shown += async (_, _) => await StartMandatorySignInFlowAsync();
        }

        private async Task StartMandatorySignInFlowAsync()
        {
            if (_initialized)
            {
                return;
            }

            _initialized = true;
            _verificationPollCount = 0;

            try
            {
                _userDataFolder = Path.Combine(
                    Path.GetTempPath(),
                    "21V M365助手",
                    "AuthWebView2",
                    Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(_userDataFolder);

                var options = new CoreWebView2EnvironmentOptions
                {
                    AllowSingleSignOnUsingOSPrimaryAccount = false
                };
                CoreWebView2Environment env = await CoreWebView2Environment.CreateAsync(null, _userDataFolder, options);
                await _webView.EnsureCoreWebView2Async(env);

                ConfigureWebView();
                await ClearSessionAsync();

                _webView.CoreWebView2.Navigate(PartnerPortalUrl);
                _verificationTimer.Start();
            }
            catch (WebView2RuntimeNotFoundException ex)
            {
                MessageBox.Show(
                    T(
                        "WebView2 Runtime was not found. Install/repair Microsoft Edge WebView2 Runtime and retry.\\r\\n\\r\\nDetails: " + ex.Message,
                        "WebView2 Runtime was not found. Install/repair Microsoft Edge WebView2 Runtime and retry.\\r\\n\\r\\nDetails: " + ex.Message),
                    T("Runtime Missing", "Runtime Missing"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                DialogResult = DialogResult.Cancel;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    T("Unable to open 21V sign-in page: " + ex.Message, "Unable to open 21V sign-in page: " + ex.Message),
                    T("Open Failed", "Open Failed"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                DialogResult = DialogResult.Cancel;
                Close();
            }
        }

        private void ConfigureWebView()
        {
            CoreWebView2 core = _webView.CoreWebView2;
            core.Settings.IsStatusBarEnabled = false;
            core.NavigationCompleted += WebView_NavigationCompleted;
        }

        private async Task ClearSessionAsync()
        {
            if (_webView.CoreWebView2 == null)
            {
                return;
            }

            try
            {
                _webView.CoreWebView2.CookieManager.DeleteAllCookies();
            }
            catch
            {
                // Ignore cookie clear failures.
            }

            try
            {
                await _webView.CoreWebView2.Profile.ClearBrowsingDataAsync();
            }
            catch
            {
                // Ignore profile clear failures.
            }
        }

        private async void WebView_NavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            if (!e.IsSuccess || _webView.CoreWebView2 == null)
            {
                return;
            }

            Uri? uri = TryGetCurrentUri();
            if (IsLoginHost(uri))
            {
                string typedAccount = await TryGetEnteredAccountAsync();
                if (!string.IsNullOrWhiteSpace(typedAccount))
                {
                    _candidateAccount = typedAccount.Trim();
                }
            }

            await TryMarkVerifiedAsync(uri);
        }

        private async Task TryMarkVerifiedAsync(Uri? uri)
        {
            if (IsVerified || !IsVerificationHost(uri))
            {
                return;
            }

            string account = await ExtractAccountAsync();
            if (string.IsNullOrWhiteSpace(account))
            {
                account = _candidateAccount;
                if (string.IsNullOrWhiteSpace(account))
                {
                    account = await ExtractDisplayNameAsync();
                    if (string.IsNullOrWhiteSpace(account))
                    {
                        return;
                    }
                }
            }

            account = account.Trim();
            if (!IsValidAccount(account))
            {
                // Fallback display name is allowed for UI, but not as account identity.
                string displayName = CleanNameCandidate(account);
                if (string.IsNullOrWhiteSpace(displayName))
                {
                    return;
                }

                VerifiedAccount = string.Empty;
                VerifiedUserName = displayName;
                IsVerified = true;
                _verificationTimer.Stop();
                DialogResult = DialogResult.OK;
                CloseWhenReady();
                return;
            }

            VerifiedAccount = account;
            VerifiedUserName = ExtractUserNameFromAccount(account);
            IsVerified = true;
            _verificationTimer.Stop();
            DialogResult = DialogResult.OK;
            CloseWhenReady();
        }

        private void CloseWhenReady()
        {
            this.BeginInvokeWhenReady(Close);
        }

        private async Task<string> ExtractAccountAsync()
        {
            string stateJson = await TryEvalStringAsync(
                "(() => { try { return JSON.stringify(window.__INITIAL_STATE__ || window.initialState || window.__PRELOADED_STATE__ || window.__NEXT_DATA__ || null); } catch (e) { return ''; } })();");
            string account = ExtractEmail(stateJson);
            if (!string.IsNullOrWhiteSpace(account))
            {
                return account;
            }

            string html = await TryEvalStringAsync(
                "(() => { try { return document && document.documentElement ? document.documentElement.outerHTML : ''; } catch (e) { return ''; } })();");
            account = ExtractEmail(html);
            if (!string.IsNullOrWhiteSpace(account))
            {
                return account;
            }

            string title = await TryEvalStringAsync(
                "(() => { try { return document && document.title ? document.title : ''; } catch (e) { return ''; } })();");
            account = ExtractEmail(title);
            if (!string.IsNullOrWhiteSpace(account))
            {
                return account;
            }

            string body = await TryEvalStringAsync(
                "(() => { try { return document && document.body ? document.body.innerText : ''; } catch (e) { return ''; } })();");
            account = ExtractEmail(body);
            if (!string.IsNullOrWhiteSpace(account))
            {
                return account;
            }

            string source = _webView.CoreWebView2?.Source ?? string.Empty;
            return ExtractEmail(source);
        }

        private async Task<string> ExtractDisplayNameAsync()
        {
            string stateJson = await TryEvalStringAsync(
                "(() => { try { return JSON.stringify(window.__INITIAL_STATE__ || window.initialState || window.__PRELOADED_STATE__ || window.__NEXT_DATA__ || null); } catch (e) { return ''; } })();");
            string fromState = ExtractNameCandidate(stateJson);
            if (!string.IsNullOrWhiteSpace(fromState))
            {
                return fromState;
            }

            string directName = await TryEvalStringAsync(
                "(() => { try {" +
                "const sels=[" +
                "\"[data-automation-id*='me' i]\"," +
                "\"[data-testid*='persona' i]\"," +
                "\"[aria-label*='account' i]\"," +
                "\"[aria-label*='闂傚倸鍊搁崐宄懊归崶褏鏆﹂柛顭戝亝閸欏繘鏌ｉ姀銏╃劸缂佲偓婢跺本鍠愰柡鍌涱儥濞兼牕霉閻樺樊鍎忕紒鈧€ｎ偁浜滈柡宥冨妿椤ｅ弶淇? i]\"," +
                "\"[aria-label*='闂傚倸鍊搁崐鐑芥倿閿曞倹鍎戠憸鐗堝笒閺勩儵鏌涢弴銊ョ仩闁搞劌鍊垮娲敆閳ь剛绮旂€靛摜鐭嗗鑸靛姈閻撴稓鈧箍鍎辨鎼佺嵁閺嶎厽鐓? i]\"," +
                "\"#mectrl_currentAccount_primary\"," +
                "\"#mectrl_currentAccount_secondary\"" +
                "];" +
                "for(const s of sels){const el=document.querySelector(s);if(!el)continue;const t=(el.innerText||el.textContent||el.getAttribute('aria-label')||'').trim();if(t)return t;}" +
                "return ''; } catch (e) { return ''; } })();");
            string fromDirect = CleanNameCandidate(directName);
            if (!string.IsNullOrWhiteSpace(fromDirect))
            {
                return fromDirect;
            }

            string body = await TryEvalStringAsync(
                "(() => { try { return document && document.body ? document.body.innerText : ''; } catch (e) { return ''; } })();");
            return PickBestLineAsName(body);
        }

        private async Task<string> TryGetEnteredAccountAsync()
        {
            string typed = await TryEvalStringAsync(
                "(() => { try { const el = document.querySelector('input[type=\"email\"],input[name*=\"user\" i],input[name*=\"login\" i],input[id*=\"user\" i],input[id*=\"login\" i]'); return el && el.value ? String(el.value) : ''; } catch (e) { return ''; } })();");
            string email = ExtractEmail(typed);
            return IsValidAccount(email) ? email : string.Empty;
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

        private static bool IsLoginHost(Uri? uri)
        {
            return uri != null &&
                   uri.Host.StartsWith("login.", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsVerificationHost(Uri? uri)
        {
            if (uri == null)
            {
                return false;
            }

            string host = uri.Host.ToLowerInvariant();
            if (host.StartsWith("login.", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return host.Equals("portal.partner.microsoftonline.cn", StringComparison.OrdinalIgnoreCase) ||
                   host.Equals("microsoft365.microsoftonline.cn", StringComparison.OrdinalIgnoreCase) ||
                   host.EndsWith(".microsoftonline.cn", StringComparison.OrdinalIgnoreCase) ||
                   host.EndsWith(".office365.cn", StringComparison.OrdinalIgnoreCase);
        }

        private static string ExtractEmail(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            Match match = EmailRegex.Match(text);
            return match.Success ? match.Value.Trim() : string.Empty;
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

        private static string ExtractNameCandidate(string source)
        {
            if (string.IsNullOrWhiteSpace(source))
            {
                return string.Empty;
            }

            string[] patterns =
            {
                "\"displayName\"\\s*:\\s*\"(?<v>[^\"]{2,80})\"",
                "\"userName\"\\s*:\\s*\"(?<v>[^\"]{2,80})\"",
                "\"name\"\\s*:\\s*\"(?<v>[^\"]{2,80})\"",
                "\"fullName\"\\s*:\\s*\"(?<v>[^\"]{2,80})\""
            };

            foreach (string pattern in patterns)
            {
                Match match = Regex.Match(source, pattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                if (!match.Success)
                {
                    continue;
                }

                string candidate = CleanNameCandidate(Regex.Unescape(match.Groups["v"].Value));
                if (!string.IsNullOrWhiteSpace(candidate))
                {
                    return candidate;
                }
            }

            return string.Empty;
        }

        private static string PickBestLineAsName(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            string[] blocked =
            {
                "home", "search", "create", "apps", "recent", "shared", "folder",
                "settings", "help", "microsoft", "portal", "partner"
            };

            string best = string.Empty;
            int bestScore = int.MinValue;
            string[] lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string raw in lines)
            {
                string candidate = CleanNameCandidate(raw);
                if (string.IsNullOrWhiteSpace(candidate))
                {
                    continue;
                }

                if (ExtractEmail(candidate).Length > 0)
                {
                    return ExtractEmail(candidate);
                }

                int score = 0;
                if (candidate.Length is >= 2 and <= 30)
                {
                    score += 3;
                }
                else if (candidate.Length <= 45)
                {
                    score += 1;
                }
                else
                {
                    score -= 3;
                }

                if (candidate.Any(char.IsLetter))
                {
                    score += 2;
                }

                if (Regex.IsMatch(candidate, @"[\u4e00-\u9fff]", RegexOptions.CultureInvariant))
                {
                    score += 2;
                }

                if (blocked.Any(x => candidate.Contains(x, StringComparison.OrdinalIgnoreCase)))
                {
                    score -= 6;
                }

                if (score > bestScore)
                {
                    bestScore = score;
                    best = candidate;
                }
            }

            return bestScore >= 2 ? best : string.Empty;
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

            while (candidate.Contains("  ", StringComparison.Ordinal))
            {
                candidate = candidate.Replace("  ", " ", StringComparison.Ordinal);
            }

            if (candidate.Length < 2 || candidate.Length > 80)
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

            _verificationPollCount++;
            if (_verificationPollCount > MaxVerificationPollCount)
            {
                _verificationTimer.Stop();
                return;
            }

            await TryMarkVerifiedAsync(TryGetCurrentUri());
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _verificationTimer.Stop();

            try
            {
                if (_webView.CoreWebView2 != null)
                {
                    _webView.CoreWebView2.NavigationCompleted -= WebView_NavigationCompleted;
                }
            }
            catch
            {
                // Ignore unhook failures.
            }

            base.OnFormClosed(e);
            TryDeleteUserDataFolder();
            _verificationTimer.Dispose();
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

        private string T(string zh, string en) => LocalizationService.Localize(_language, zh, en);
    }
}
