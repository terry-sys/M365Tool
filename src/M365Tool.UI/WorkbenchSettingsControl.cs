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

        public event EventHandler<string>? LanguageModeChanged;

        private sealed class LanguageModeOption
        {
            public string Value { get; set; } = "Auto";

            public string Display { get; set; } = string.Empty;

            public override string ToString() => Display;
        }

        public WorkbenchSettingsControl(AppSettingsService settingsService, AppSettings settings, UiLanguage language)
        {
            AutoScaleMode = AutoScaleMode.Dpi;
            _settingsService = settingsService;
            _settings = settings;
            _language = language;

            Dock = DockStyle.Fill;
            BackColor = WorkbenchUi.CanvasColor;
            AutoScroll = true;

            _lblTitle = new Label
            {
                Font = new Font("Microsoft YaHei UI", 17F, FontStyle.Bold),
                ForeColor = WorkbenchUi.PrimaryTextColor,
                BackColor = Color.Transparent
            };

            _lblDescription = new Label
            {
                Font = new Font("Microsoft YaHei UI", 10.25F),
                ForeColor = WorkbenchUi.SecondaryTextColor,
                BackColor = Color.Transparent
            };

            _settingsCard = WorkbenchUi.CreateCard(Point.Empty, Size.Empty);
            _lblLanguageTitle = WorkbenchUi.CreateSectionTitle("语言", new Point(20, 16), 180);
            _lblLanguageMode = WorkbenchUi.CreateFieldLabel("显示语言", new Point(20, 68), 180);
            _lblLanguageHint = new Label
            {
                Font = new Font("Microsoft YaHei UI", 9F),
                ForeColor = WorkbenchUi.SecondaryTextColor,
                BackColor = Color.Transparent,
                Location = new Point(20, 44),
                Size = new Size(460, 18)
            };
            _cboLanguageMode = WorkbenchUi.CreateComboBox(new Point(20, 96), new Size(280, 32));
            _lblCurrentLanguage = new Label
            {
                Font = new Font("Microsoft YaHei UI", 9.5F),
                ForeColor = WorkbenchUi.SecondaryTextColor,
                BackColor = Color.Transparent,
                Location = new Point(20, 138),
                Size = new Size(760, 24)
            };
            _lblValidationState = new Label
            {
                Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold),
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
                Font = new Font("Microsoft YaHei UI", 10F),
                ForeColor = WorkbenchUi.PrimaryTextColor,
                BackColor = Color.Transparent,
                Location = new Point(20, 48),
                Size = new Size(640, 42)
            };
            _supportActionsPanel = WorkbenchUi.CreateInsetPanel(Point.Empty, Size.Empty);
            _lblSupportActionsCaption = WorkbenchUi.CreateFieldLabel("可通过公众号完成", new Point(16, 12), 220);
            _lblSupportActions = new Label
            {
                Font = new Font("Microsoft YaHei UI", 9.5F),
                ForeColor = WorkbenchUi.SecondaryTextColor,
                BackColor = Color.Transparent,
                Location = new Point(16, 40),
                Size = new Size(600, 84)
            };
            _lblSupportHint = new Label
            {
                Font = new Font("Microsoft YaHei UI", 9F),
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

            Controls.AddRange(new Control[] { _lblTitle, _lblDescription, _settingsCard, _supportCard });

            Resize += (_, _) => LayoutControls();
            Scroll += (_, _) => _cboLanguageMode.DroppedDown = false;
            MouseWheel += (_, _) => _cboLanguageMode.DroppedDown = false;

            LoadSupportQrImage();
            BindLanguageModeOptions();
            ApplyLanguage(language);
            LayoutControls();
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
            BindLanguageModeOptions();
        }

        private void LayoutControls()
        {
            const int margin = 24;
            int viewportWidth = ClientSize.Width;
            int availableWidth = Math.Max(560, Math.Min(1160, viewportWidth - margin * 2));
            int contentLeft = Math.Max(margin, (viewportWidth - availableWidth) / 2);

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

            AutoScrollMinSize = new Size(0, _supportCard.Bottom + 24);
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
