using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Office365CleanupTool.Services;

namespace Office365CleanupTool
{
    public class WorkbenchSettingsControl : UserControl
    {
        private readonly AppSettingsService _settingsService;
        private readonly AppSettings _settings;
        private UiLanguage _language;

        private readonly Panel _scrollHost;
        private readonly Panel _contentPanel;
        private readonly Label _lblTitle;
        private readonly Label _lblDescription;
        private readonly Panel _settingsCard;
        private readonly Label _lblLanguageTitle;
        private readonly Label _lblLanguageMode;
        private readonly Label _lblLanguageHint;
        private readonly ComboBox _cboLanguageMode;
        private readonly Label _lblCurrentLanguage;
        private readonly Label _lblValidationState;
        private readonly Button _btnSave;
        private readonly Panel _supportCard;
        private readonly Label _lblSupportTitle;
        private readonly Label _lblSupportTag;
        private readonly Label _lblSupportSummary;
        private readonly Panel _supportActionsPanel;
        private readonly Label _lblSupportActionsCaption;
        private readonly Label _lblSupportActions;
        private readonly Label _lblSupportHint;
        private readonly Panel _qrPanel;
        private readonly PictureBox _picSupportQr;
        private readonly Button _btnOpenQr;
        private readonly Panel _privacyCard;
        private readonly Label _lblPrivacyTitle;
        private readonly Label _lblPrivacySummary;
        private readonly Label _lblPrivacyContent;
        private readonly Button _btnOpenPrivacyPolicy;

        public event EventHandler<string>? LanguageModeChanged;

        private sealed class LanguageModeOption
        {
            public string Value { get; set; } = "Auto";

            public string Display { get; set; } = string.Empty;

            public override string ToString() => Display;
        }

        public WorkbenchSettingsControl(AppSettingsService settingsService, AppSettings settings, UiLanguage language)
        {
            AutoScaleMode = AutoScaleMode.None;
            _settingsService = settingsService;
            _settings = settings;
            _language = language;

            Dock = DockStyle.Fill;
            BackColor = WorkbenchUi.CanvasColor;
            AutoScroll = false;

            _scrollHost = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = WorkbenchUi.CanvasColor,
                AutoScroll = true
            };

            _contentPanel = new Panel
            {
                Location = new Point(0, 0),
                BackColor = WorkbenchUi.CanvasColor,
                Dock = DockStyle.Top
            };

            _lblTitle = new Label
            {
                Font = WorkbenchUi.CreateUiFont(17F, FontStyle.Bold),
                ForeColor = WorkbenchUi.PrimaryTextColor,
                BackColor = Color.Transparent
            };

            _lblDescription = new Label
            {
                Font = WorkbenchUi.CreateUiFont(10.25F),
                ForeColor = WorkbenchUi.SecondaryTextColor,
                BackColor = Color.Transparent
            };

            _settingsCard = WorkbenchUi.CreateCard(Point.Empty, Size.Empty);
            _lblLanguageTitle = WorkbenchUi.CreateSectionTitle("语言", new Point(20, 16), 180);
            _lblLanguageMode = WorkbenchUi.CreateFieldLabel("显示语言", new Point(20, 68), 180);
            _lblLanguageHint = new Label
            {
                Font = WorkbenchUi.CreateUiFont(9F),
                ForeColor = WorkbenchUi.SecondaryTextColor,
                BackColor = Color.Transparent,
                Location = new Point(20, 44),
                Size = new Size(460, 18)
            };
            _cboLanguageMode = WorkbenchUi.CreateComboBox(new Point(20, 96), new Size(280, 32));
            _lblCurrentLanguage = new Label
            {
                Font = WorkbenchUi.CreateUiFont(9.5F),
                ForeColor = WorkbenchUi.SecondaryTextColor,
                BackColor = Color.Transparent,
                Location = new Point(20, 138),
                Size = new Size(760, 24)
            };
            _lblValidationState = new Label
            {
                Font = WorkbenchUi.CreateUiFont(9F, FontStyle.Bold),
                ForeColor = WorkbenchUi.SecondaryTextColor,
                Location = new Point(20, 170),
                Size = new Size(188, 28)
            };
            WorkbenchUi.ApplyStatusStyle(_lblValidationState, WorkbenchTone.Default);
            _lblValidationState.Visible = false;

            _btnSave = WorkbenchUi.CreatePrimaryButton("保存设置", new Point(20, 194), new Size(152, 42));
            _btnSave.Click += btnSave_Click;

            _settingsCard.Controls.AddRange(new Control[]
            {
                _lblLanguageTitle, _lblLanguageHint, _lblLanguageMode, _cboLanguageMode,
                _lblCurrentLanguage, _lblValidationState,
                _btnSave
            });

            _supportCard = WorkbenchUi.CreateCard(Point.Empty, Size.Empty);
            _lblSupportTitle = WorkbenchUi.CreateSectionTitle("反馈与支持", new Point(20, 16), 220);
            _lblSupportTag = WorkbenchUi.CreateTagLabel("公众号支持", Point.Empty, 104, WorkbenchTone.Info);
            _lblSupportTag.Visible = false;
            _lblSupportSummary = new Label
            {
                Font = WorkbenchUi.CreateUiFont(10F),
                ForeColor = WorkbenchUi.PrimaryTextColor,
                BackColor = Color.Transparent,
                Location = new Point(20, 48),
                Size = new Size(640, 42)
            };
            _supportActionsPanel = WorkbenchUi.CreateInsetPanel(Point.Empty, Size.Empty);
            _lblSupportActionsCaption = WorkbenchUi.CreateFieldLabel("可通过公众号完成", new Point(16, 12), 220);
            _lblSupportActions = new Label
            {
                Font = WorkbenchUi.CreateUiFont(9.5F),
                ForeColor = WorkbenchUi.SecondaryTextColor,
                BackColor = Color.Transparent,
                Location = new Point(16, 40),
                Size = new Size(600, 84)
            };
            _lblSupportHint = new Label
            {
                Font = WorkbenchUi.CreateUiFont(9F),
                ForeColor = WorkbenchUi.SecondaryTextColor,
                BackColor = Color.Transparent,
                Location = new Point(20, 190),
                Size = new Size(640, 40)
            };
            _qrPanel = WorkbenchUi.CreateInsetPanel(Point.Empty, Size.Empty);
            _picSupportQr = new PictureBox
            {
                BackColor = Color.White,
                BorderStyle = BorderStyle.None,
                SizeMode = PictureBoxSizeMode.Zoom
            };
            _btnOpenQr = WorkbenchUi.CreateSecondaryButton("打开二维码", Point.Empty, new Size(132, 38));
            _btnOpenQr.Click += btnOpenQr_Click;
            _qrPanel.Controls.AddRange(new Control[] { _picSupportQr, _btnOpenQr });
            _supportActionsPanel.Controls.AddRange(new Control[] { _lblSupportActionsCaption, _lblSupportActions });

            _supportCard.Controls.AddRange(new Control[]
            {
                _lblSupportTitle,
                _lblSupportTag,
                _lblSupportSummary,
                _supportActionsPanel,
                _lblSupportHint,
                _qrPanel
            });

            _privacyCard = WorkbenchUi.CreateCard(Point.Empty, Size.Empty);
            _lblPrivacyTitle = WorkbenchUi.CreateSectionTitle("隐私政策", new Point(20, 16), 220);
            _lblPrivacySummary = new Label
            {
                Font = WorkbenchUi.CreateUiFont(10F),
                ForeColor = WorkbenchUi.PrimaryTextColor,
                BackColor = Color.Transparent,
                Location = new Point(20, 48),
                Size = new Size(840, 42)
            };
            _lblPrivacyContent = new Label
            {
                Font = WorkbenchUi.CreateUiFont(9.5F),
                ForeColor = WorkbenchUi.SecondaryTextColor,
                BackColor = Color.Transparent,
                Location = new Point(20, 96),
                Size = new Size(840, 168)
            };
            _btnOpenPrivacyPolicy = WorkbenchUi.CreateSecondaryButton("查看完整隐私政策", new Point(20, 96), new Size(220, 44));
            _btnOpenPrivacyPolicy.Font = WorkbenchUi.CreateUiFont(9.5F, FontStyle.Bold);
            _btnOpenPrivacyPolicy.Click += btnOpenPrivacyPolicy_Click;
            _privacyCard.Controls.AddRange(new Control[]
            {
                _lblPrivacyTitle,
                _lblPrivacySummary,
                _lblPrivacyContent,
                _btnOpenPrivacyPolicy
            });

            _contentPanel.Controls.AddRange(new Control[] { _lblTitle, _lblDescription, _settingsCard, _supportCard, _privacyCard });
            _scrollHost.Controls.Add(_contentPanel);
            Controls.Add(_scrollHost);

            Resize += (_, _) => LayoutControls();
            _scrollHost.Resize += (_, _) => LayoutControls();
            _scrollHost.Scroll += (_, _) => _cboLanguageMode.DroppedDown = false;
            _scrollHost.MouseWheel += (_, _) => _cboLanguageMode.DroppedDown = false;
            ParentChanged += (_, _) => this.BeginInvokeWhenReady(PrepareForDisplay);
            VisibleChanged += (_, _) =>
            {
                if (Visible)
                {
                    this.BeginInvokeWhenReady(PrepareForDisplay);
                }
            };

            LoadSupportQrImage();
            BindLanguageModeOptions();
            ApplyLanguage(language);
            LayoutControls();
        }

        public void PrepareForDisplay()
        {
            if (_scrollHost.IsHandleCreated)
            {
                _scrollHost.AutoScrollPosition = new Point(0, 0);
            }

            LayoutControls();
            this.BeginInvokeWhenReady(LayoutControls);
        }

        public void ApplyLanguage(UiLanguage language)
        {
            _language = language;
            _lblTitle.Text = T("关于", "About");
            _lblTitle.Visible = false;
            _lblDescription.Text = string.Empty;
            _lblDescription.Visible = false;
            _lblLanguageTitle.Text = T("界面语言", "Interface Language");
            _lblLanguageHint.Text = T("语言切换后会立即刷新当前窗口显示。", "Language changes are applied immediately to the current window.");
            _lblLanguageMode.Text = T("显示语言", "Display Language");
            _btnSave.Text = T("保存语言设置", "Save Language");
            _lblCurrentLanguage.Text = T(
                $"当前实际界面语言：{GetLanguageDisplay(LocalizationService.ResolveLanguage(_settings.LanguageMode))}",
                $"Current UI language: {GetLanguageDisplay(LocalizationService.ResolveLanguage(_settings.LanguageMode))}");
            _lblValidationState.Text = T(
                _settings.HasValidated21VAccount ? "21V 门户验证：已通过" : "21V 门户验证：未通过",
                _settings.HasValidated21VAccount ? "21V verification: Passed" : "21V verification: Not passed");
            WorkbenchUi.ApplyStatusStyle(
                _lblValidationState,
                _settings.HasValidated21VAccount ? WorkbenchTone.Success : WorkbenchTone.Warning);
            _lblSupportTitle.Text = T("反馈与支持", "Feedback & Support");
            _lblSupportTag.Text = T("公众号支持", "Service Entry");
            _lblSupportSummary.Text = T(
                "扫码关注公众号，可提交使用反馈、联系我们，并开启技术支持工单。",
                "Scan the service account QR code to submit feedback, contact us, and open a technical support ticket.");
            _lblSupportActionsCaption.Text = T("可通过公众号完成", "Available via the service account");
            _lblSupportActions.Text = T(
                "可通过公众号完成：\r\n• 产品问题反馈\r\n• 联系支持团队\r\n• 发起技术支持工单",
                "Use the service account to:\r\n• Submit product feedback\r\n• Contact the support team\r\n• Open a technical support ticket");
            _lblSupportHint.Text = HasSupportQrImage()
                ? T("建议使用微信扫码；如需放大查看，可点击右下角按钮打开二维码图片。", "We recommend scanning with WeChat. Click the button below to open the QR image in full size.")
                : T("当前未找到二维码图片文件，请联系管理员补充公众号二维码资源。", "QR code image file was not found. Please contact the administrator to add the support QR asset.");
            _btnOpenQr.Text = T("打开二维码", "Open QR");
            _lblPrivacyTitle.Text = T("隐私政策", "Privacy Policy");
            _lblPrivacySummary.Text = T(
                "本工具用于 21Vianet Microsoft 365 运维场景，登录验证和本地日志仅用于提供工具功能、排障与支持。",
                "This tool is built for 21Vianet Microsoft 365 operations. Sign-in verification and local logs are used only to provide tool features, troubleshooting, and support.");
            _btnOpenPrivacyPolicy.Text = T("查看完整隐私政策", "View Full Privacy Policy");
            _lblPrivacyContent.Text = T(
                "我们可能处理的信息：\r\n• 21V 账户登录验证状态、显示名、账号标识和头像缓存，用于显示登录状态和限制未登录功能。\r\n• 工具配置、安装/卸载/修复/诊断执行日志和返回码，用于本机排障、结果查看和支持工单。\r\n• 二维码反馈或技术支持中由你主动提交的信息。\r\n\r\n数据使用原则：\r\n• 本工具不出售个人信息，不进行画像或自动化决策。\r\n• 头像、设置和日志优先保存在本机用户目录；除你主动提交反馈或工单外，工具不会主动上传日志内容。\r\n• 你可以通过注销、清理缓存或删除本机日志目录来移除本地保存的账号状态和执行记录。",
                "Information this tool may process:\r\n• 21V account verification state, display name, account identifier, and avatar cache for sign-in status and feature gating.\r\n• Tool settings, install/uninstall/repair/diagnostic logs, and exit codes for local troubleshooting, result review, and support tickets.\r\n• Information you voluntarily submit through QR-code feedback or technical support.\r\n\r\nData use principles:\r\n• This tool does not sell personal information, build profiles, or make automated decisions.\r\n• Avatars, settings, and logs are stored locally in the user profile by default. The tool does not actively upload logs unless you submit feedback or a support ticket.\r\n• You can sign out, clear cache, or delete the local log folder to remove locally stored account state and execution records.");
            BindLanguageModeOptions();
        }

        private void LayoutControls()
        {
            if (_scrollHost.ClientSize.Width <= 0 || _scrollHost.ClientSize.Height <= 0)
            {
                return;
            }

            int viewportWidth = _scrollHost.ClientSize.Width;
            int margin = WorkbenchUi.GetAdaptivePageMargin(viewportWidth);
            int availableWidth = WorkbenchUi.GetAdaptiveContentWidth(
                viewportWidth,
                560,
                WorkbenchUi.ReadingContentMaxWidth,
                margin);
            int contentLeft = WorkbenchUi.GetAdaptiveContentLeft(viewportWidth, availableWidth, margin);

            _settingsCard.Location = new Point(contentLeft, 18);
            _settingsCard.Size = new Size(availableWidth, 246);
            _lblLanguageHint.Size = new Size(Math.Min(availableWidth - 180, 460), 18);
            _btnSave.Location = new Point(20, _settingsCard.Height - _btnSave.Height - 20);

            bool stackSupport = availableWidth < 860;
            int qrSize = stackSupport ? 170 : 172;
            int supportHeight = stackSupport ? 558 : 304;
            _supportCard.Location = new Point(contentLeft, _settingsCard.Bottom + 18);
            _supportCard.Size = new Size(availableWidth, supportHeight);
            _lblSupportTag.Location = new Point(availableWidth - _lblSupportTag.Width - 20, 16);

            if (stackSupport)
            {
                _lblSupportSummary.Size = new Size(availableWidth - 40, 48);
                _supportActionsPanel.Location = new Point(20, 104);
                _supportActionsPanel.Size = new Size(availableWidth - 40, 132);
                _lblSupportActions.Size = new Size(_supportActionsPanel.Width - 32, 84);
                _qrPanel.Location = new Point((availableWidth - 248) / 2, 250);
                _qrPanel.Size = new Size(248, 248);
                _picSupportQr.Location = new Point((_qrPanel.Width - qrSize) / 2, 18);
                _picSupportQr.Size = new Size(qrSize, qrSize);
                _btnOpenQr.Location = new Point((_qrPanel.Width - _btnOpenQr.Width) / 2, _picSupportQr.Bottom + 12);
                _lblSupportHint.Location = new Point(20, _qrPanel.Bottom + 10);
                _lblSupportHint.Size = new Size(availableWidth - 40, 36);
            }
            else
            {
                int qrPanelWidth = 236;
                int textWidth = availableWidth - qrPanelWidth - 76;
                _lblSupportSummary.Size = new Size(textWidth, 48);
                _supportActionsPanel.Location = new Point(20, 104);
                _supportActionsPanel.Size = new Size(textWidth, 132);
                _lblSupportActions.Size = new Size(_supportActionsPanel.Width - 32, 84);
                _lblSupportHint.Location = new Point(20, 246);
                _lblSupportHint.Size = new Size(textWidth, 36);
                _qrPanel.Location = new Point(availableWidth - qrPanelWidth - 20, 28);
                _qrPanel.Size = new Size(qrPanelWidth, 248);
                _picSupportQr.Location = new Point((_qrPanel.Width - qrSize) / 2, 18);
                _picSupportQr.Size = new Size(qrSize, qrSize);
                _btnOpenQr.Location = new Point((_qrPanel.Width - _btnOpenQr.Width) / 2, _picSupportQr.Bottom + 12);
            }

            int privacyHeight = stackSupport ? 420 : 360;
            _privacyCard.Location = new Point(contentLeft, _supportCard.Bottom + 18);
            _privacyCard.Size = new Size(availableWidth, privacyHeight);
            _lblPrivacySummary.Size = new Size(availableWidth - 40, 44);
            SizePrivacyPolicyButton(availableWidth - 40);
            _btnOpenPrivacyPolicy.Location = new Point(20, 106);
            int privacyContentTop = _btnOpenPrivacyPolicy.Bottom + 22;
            _lblPrivacyContent.Location = new Point(20, privacyContentTop);
            _lblPrivacyContent.Size = new Size(availableWidth - 40, privacyHeight - privacyContentTop - 24);

            _contentPanel.Width = Math.Max(10, viewportWidth - 1);
            _contentPanel.Height = _privacyCard.Bottom + 24;
        }

        private void SizePrivacyPolicyButton(int maxWidth)
        {
            int measuredWidth = TextRenderer.MeasureText(_btnOpenPrivacyPolicy.Text, _btnOpenPrivacyPolicy.Font).Width + 44;
            int buttonWidth = Math.Min(Math.Max(220, measuredWidth), Math.Max(160, maxWidth));
            _btnOpenPrivacyPolicy.Size = new Size(buttonWidth, 44);
        }

        private void BindLanguageModeOptions()
        {
            string currentValue = _settings.LanguageMode;
            _cboLanguageMode.Items.Clear();
            _cboLanguageMode.Items.Add(new LanguageModeOption { Value = "Auto", Display = T("跟随系统", "Follow system") });
            _cboLanguageMode.Items.Add(new LanguageModeOption { Value = "zh-CN", Display = "中文" });
            _cboLanguageMode.Items.Add(new LanguageModeOption { Value = "en-US", Display = "English" });

            for (int i = 0; i < _cboLanguageMode.Items.Count; i++)
            {
                if (_cboLanguageMode.Items[i] is LanguageModeOption option &&
                    string.Equals(option.Value, currentValue, StringComparison.OrdinalIgnoreCase))
                {
                    _cboLanguageMode.SelectedIndex = i;
                    return;
                }
            }

            _cboLanguageMode.SelectedIndex = 0;
        }

        private void btnSave_Click(object? sender, EventArgs e)
        {
            if (_cboLanguageMode.SelectedItem is not LanguageModeOption option)
            {
                return;
            }

            _settings.LanguageMode = option.Value;
            _settingsService.Save(_settings);
            LanguageModeChanged?.Invoke(this, option.Value);
            ApplyLanguage(LocalizationService.ResolveLanguage(_settings.LanguageMode));

            MessageBox.Show(
                T("设置已保存。", "Settings saved."),
                T("完成", "Done"),
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void LoadSupportQrImage()
        {
            using Stream? stream = OpenSupportQrStream();
            if (stream is null)
            {
                _picSupportQr.Image = null;
                return;
            }

            using var image = Image.FromStream(stream);
            _picSupportQr.Image?.Dispose();
            _picSupportQr.Image = new Bitmap(image);
        }

        private void btnOpenQr_Click(object? sender, EventArgs e)
        {
            string? qrPath = ResolveSupportQrFilePath();
            if (string.IsNullOrWhiteSpace(qrPath))
            {
                MessageBox.Show(
                    T("未找到二维码图片，请联系管理员补充资源。", "QR image not found. Please contact the administrator."),
                    T("提示", "Notice"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            Process.Start(new ProcessStartInfo
            {
                FileName = qrPath,
                UseShellExecute = true
            });
        }

        private void btnOpenPrivacyPolicy_Click(object? sender, EventArgs e)
        {
            string html = ReadEmbeddedPrivacyPolicyHtml();
            if (string.IsNullOrWhiteSpace(html))
            {
                MessageBox.Show(
                    T("未找到隐私政策 HTML 资源。", "Privacy policy HTML resource was not found."),
                    T("提示", "Notice"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            using var dialog = new PrivacyPolicyForm(_language, html);
            dialog.ShowDialog(this);
        }

        private static string ReadEmbeddedPrivacyPolicyHtml()
        {
            using Stream? stream = typeof(WorkbenchSettingsControl)
                .Assembly
                .GetManifestResourceStream("M365Tool.Assets.privacy.html");
            if (stream is null)
            {
                return string.Empty;
            }

            using var reader = new StreamReader(stream);
            return reader.ReadToEnd();
        }

        private static Stream? OpenSupportQrStream()
        {
            string externalPath = Path.Combine(AppContext.BaseDirectory, "Assets", "support-wechat-qr.png");
            if (File.Exists(externalPath))
            {
                return File.OpenRead(externalPath);
            }

            return typeof(WorkbenchSettingsControl).Assembly.GetManifestResourceStream("M365Tool.Assets.support-wechat-qr.png");
        }

        private static bool HasSupportQrImage()
        {
            using Stream? stream = OpenSupportQrStream();
            return stream is not null;
        }

        private static string? ResolveSupportQrFilePath()
        {
            string externalPath = Path.Combine(AppContext.BaseDirectory, "Assets", "support-wechat-qr.png");
            if (File.Exists(externalPath))
            {
                return externalPath;
            }

            using Stream? stream = OpenSupportQrStream();
            if (stream is null)
            {
                return null;
            }

            string tempDirectory = Path.Combine(Path.GetTempPath(), "M365Tool");
            Directory.CreateDirectory(tempDirectory);
            string tempPath = Path.Combine(tempDirectory, "support-wechat-qr.png");
            using FileStream output = File.Create(tempPath);
            stream.CopyTo(output);
            return tempPath;
        }

        private string T(string zh, string en) => LocalizationService.Localize(_language, zh, en);

        private string GetLanguageDisplay(UiLanguage language)
        {
            return language == UiLanguage.Chinese ? "中文" : "English";
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _picSupportQr.Image?.Dispose();
            }

            base.Dispose(disposing);
        }
    }
}
