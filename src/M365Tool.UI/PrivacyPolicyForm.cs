using System;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Office365CleanupTool.Services;

namespace Office365CleanupTool
{
    public sealed class PrivacyPolicyForm : Form
    {
        private readonly UiLanguage _language;
        private readonly string _html;
        private readonly WebView2 _webView;
        private readonly Label _statusLabel;
        private string _userDataFolder = string.Empty;
        private bool _initialized;

        public PrivacyPolicyForm(UiLanguage language, string html)
        {
            _language = language;
            _html = html ?? string.Empty;

            Text = T("隐私政策", "Privacy Policy");
            Icon = AppIconProvider.GetAppIcon() ?? Icon;
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(900, 640);
            Size = new Size(1080, 760);
            BackColor = Color.White;

            _webView = new WebView2
            {
                Dock = DockStyle.Fill
            };

            _statusLabel = new Label
            {
                Dock = DockStyle.Fill,
                Text = T("正在加载隐私政策...", "Loading privacy policy..."),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = WorkbenchUi.CreateUiFont(11F),
                ForeColor = WorkbenchUi.SecondaryTextColor,
                BackColor = Color.White
            };

            Controls.Add(_webView);
            Controls.Add(_statusLabel);
            _statusLabel.BringToFront();

            Shown += async (_, _) => await InitializeAsync();
        }

        private async Task InitializeAsync()
        {
            if (_initialized)
            {
                return;
            }

            _initialized = true;

            try
            {
                _userDataFolder = Path.Combine(
                    Path.GetTempPath(),
                    "M365Tool",
                    "PrivacyWebView2",
                    Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(_userDataFolder);

                CoreWebView2Environment env = await CoreWebView2Environment.CreateAsync(null, _userDataFolder);
                await _webView.EnsureCoreWebView2Async(env);
                _webView.CoreWebView2.Settings.IsStatusBarEnabled = false;
                _webView.CoreWebView2.NavigateToString(_html);
                _statusLabel.Visible = false;
                _webView.BringToFront();
            }
            catch (WebView2RuntimeNotFoundException)
            {
                ShowFallbackText(T(
                    "未找到 Microsoft Edge WebView2 Runtime，暂时无法在工具内展示 HTML 隐私政策。",
                    "Microsoft Edge WebView2 Runtime was not found. The HTML privacy policy cannot be displayed in the tool."));
            }
            catch (Exception ex)
            {
                ShowFallbackText(T(
                    $"隐私政策加载失败：{ex.Message}",
                    $"Failed to load privacy policy: {ex.Message}"));
            }
        }

        private void ShowFallbackText(string text)
        {
            _statusLabel.Text = text;
            _statusLabel.Visible = true;
            _statusLabel.BringToFront();
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);
            TryDeleteUserDataFolder();
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
                            Directory.Delete(target, recursive: true);
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
