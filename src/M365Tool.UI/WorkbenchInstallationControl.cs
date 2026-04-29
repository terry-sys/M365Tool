using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Office365CleanupTool.Models;
using Office365CleanupTool.Services;

namespace Office365CleanupTool
{
    public class WorkbenchInstallationControl : UserControl
    {
        private sealed class BufferedPanel : Panel
        {
            public BufferedPanel()
            {
                DoubleBuffered = true;
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.UserPaint, true);
            }
        }

        private readonly InstallationService _installationService;
        private readonly ValidationService _validationService;
        private readonly InstallationConfig _currentConfig;
        private UiLanguage _language;
        private readonly Panel _scrollHost;
        private readonly Panel _contentPanel;
        private readonly Label _lblTitle;
        private readonly Label _lblDescription;
        private readonly Panel _basicCard;
        private readonly Panel _optionsCard;
        private readonly Panel _excludeAppsCard;
        private readonly Panel _languagesCard;
        private readonly Panel _precheckCard;
        private readonly Panel _resultCard;
        private readonly Panel _excludeAppsHost;
        private readonly Panel _languagesHost;
        private readonly Panel _precheckOutputHost;
        private readonly Panel _resultOutputHost;
        private readonly Label _lblBasicTitle;
        private readonly Label _lblBasicHint;
        private readonly Panel _basicFlowStrip;
        private readonly Label _lblBasicFlowText;
        private readonly Label _lblBasicFlowTag;

        private readonly ComboBox _cboEdition;
        private readonly ComboBox _cboInstallTarget;
        private readonly ComboBox _cboArchitecture;
        private readonly ComboBox _cboChannel;
        private readonly CheckBox _chkSilentInstall;
        private readonly CheckBox _chkPinToTaskbar;
        private readonly CheckedListBox _clbExcludeApps;
        private readonly TextBox _txtLanguageSearch;
        private readonly CheckedListBox _clbLanguages;
        private readonly TextBox _txtPrecheckResult;
        private readonly Button _btnRefreshChecks;
        private readonly Label _lblResult;
        private readonly TextBox _txtResultDetails;
        private readonly ProgressBar _progressBar;
        private readonly Button _btnOpenLogs;
        private readonly Button _btnInstall;
        private readonly Label _lblEdition;
        private readonly Label _lblInstallTarget;
        private readonly Label _lblArchitecture;
        private readonly Label _lblChannel;
        private readonly Label _lblExcludeAppsTitle;
        private readonly Label _lblLanguagesTitle;
        private readonly Label _lblPrecheckStatus;
        private readonly Label _lblOptionsTitle;
        private readonly Label _lblPrecheckTitle;
        private readonly Label _lblResultTitle;

        private bool _isInstalling;
        private bool _hasAppliedLanguage;
        private string _lastLogFilePath = string.Empty;

        private readonly Color _successColor = WorkbenchUi.SuccessColor;
        private readonly Color _normalColor = WorkbenchUi.SecondaryTextColor;

        private static readonly (string Code, string ChineseName, string EnglishName)[] SupportedLanguages =
        {
            ("zh-CN", "简体中文", "Chinese (Simplified)"),
            ("zh-TW", "繁体中文", "Chinese (Traditional)"),
            ("en-US", "英语（美国）", "English (United States)"),
            ("en-GB", "英语（英国）", "English (United Kingdom)"),
            ("ja-JP", "日语", "Japanese"),
            ("ko-KR", "韩语", "Korean"),
            ("fr-FR", "法语", "French"),
            ("de-DE", "德语", "German"),
            ("es-ES", "西班牙语", "Spanish (Spain)"),
            ("es-MX", "西班牙语（墨西哥）", "Spanish (Mexico)"),
            ("pt-BR", "葡萄牙语（巴西）", "Portuguese (Brazil)"),
            ("pt-PT", "葡萄牙语（葡萄牙）", "Portuguese (Portugal)"),
            ("it-IT", "意大利语", "Italian"),
            ("ru-RU", "俄语", "Russian"),
            ("ar-SA", "阿拉伯语", "Arabic (Saudi Arabia)"),
            ("tr-TR", "土耳其语", "Turkish"),
            ("pl-PL", "波兰语", "Polish"),
            ("nl-NL", "荷兰语", "Dutch"),
            ("sv-SE", "瑞典语", "Swedish"),
            ("da-DK", "丹麦语", "Danish"),
            ("fi-FI", "芬兰语", "Finnish"),
            ("nb-NO", "挪威语", "Norwegian Bokmal"),
            ("cs-CZ", "捷克语", "Czech"),
            ("hu-HU", "匈牙利语", "Hungarian"),
            ("ro-RO", "罗马尼亚语", "Romanian"),
            ("uk-UA", "乌克兰语", "Ukrainian"),
            ("el-GR", "希腊语", "Greek"),
            ("he-IL", "希伯来语", "Hebrew"),
            ("th-TH", "泰语", "Thai"),
            ("vi-VN", "越南语", "Vietnamese"),
            ("id-ID", "印度尼西亚语", "Indonesian"),
            ("ms-MY", "马来语", "Malay"),
            ("hi-IN", "印地语", "Hindi")
        };

        private sealed class ExcludeAppItem
        {
            public string Display { get; set; } = string.Empty;
            public string Value { get; set; } = string.Empty;
            public override string ToString() => Display;
        }

        private sealed class LanguageItem
        {
            public string Display { get; set; } = string.Empty;
            public string Value { get; set; } = string.Empty;
            public override string ToString() => Display;
        }

        private sealed class OptionItem
        {
            public string Key { get; set; } = string.Empty;
            public string Display { get; set; } = string.Empty;
            public override string ToString() => Display;
        }

        public WorkbenchInstallationControl(UiLanguage initialLanguage)
        {
            AutoScaleMode = AutoScaleMode.Dpi;
            Dock = DockStyle.Fill;
            BackColor = WorkbenchUi.CanvasColor;
            AutoScroll = false;
            DoubleBuffered = true;
            _language = initialLanguage;

            _installationService = new InstallationService();
            _validationService = new ValidationService();
            _currentConfig = new InstallationConfig();
            _installationService.ProgressChanged += OnInstallationProgressChanged;

            _scrollHost = new BufferedPanel
            {
                Dock = DockStyle.Fill,
                BackColor = WorkbenchUi.PageSurfaceColor,
                AutoScroll = true
            };

            _contentPanel = new BufferedPanel
            {
                Location = new Point(0, 0),
                BackColor = WorkbenchUi.PageSurfaceColor,
                Dock = DockStyle.Top
            };

            _lblTitle = new Label
            {
                Text = "安装 M365 应用",
                Font = new Font("Microsoft YaHei UI", 17F, FontStyle.Bold),
                ForeColor = WorkbenchUi.PrimaryTextColor
            };

            _lblDescription = new Label
            {
                Text = "选择安装产品、位数和频道，然后执行在线安装与环境检查",
                Font = new Font("Microsoft YaHei UI", 10F),
                ForeColor = WorkbenchUi.SecondaryTextColor
            };

            _basicCard = WorkbenchUi.CreateCard(Point.Empty, Size.Empty);
            _lblBasicTitle = WorkbenchUi.CreateSectionTitle("安装参数", new Point(20, 14), 180);
            _lblBasicHint = new Label
            {
                AutoSize = false,
                Font = new Font("Microsoft YaHei UI", 9.5F),
                ForeColor = WorkbenchUi.SecondaryTextColor,
                BackColor = Color.Transparent
            };
            _basicFlowStrip = WorkbenchUi.CreateInsetPanel(Point.Empty, Size.Empty);
            _basicFlowStrip.Visible = false;
            var lblBasicFlowIcon = new Label
            {
                AutoSize = false,
                Text = "\uE8FD",
                Font = new Font("Segoe MDL2 Assets", 13F),
                ForeColor = WorkbenchUi.PrimaryColor,
                BackColor = Color.FromArgb(235, 243, 255),
                TextAlign = ContentAlignment.MiddleCenter,
                Location = new Point(10, 6),
                Size = new Size(28, 28)
            };
            var lblBasicFlowText = new Label
            {
                AutoSize = false,
                Font = new Font("Microsoft YaHei UI", 9.25F),
                ForeColor = WorkbenchUi.SecondaryTextColor
            };
            var lblBasicFlowTag = WorkbenchUi.CreateTagLabel("流程", Point.Empty, 64, WorkbenchTone.Info);
            _basicFlowStrip.Controls.AddRange(new Control[] { lblBasicFlowIcon, lblBasicFlowText, lblBasicFlowTag });
            _lblBasicFlowText = lblBasicFlowText;
            _lblBasicFlowTag = lblBasicFlowTag;
            _lblEdition = WorkbenchUi.CreateFieldLabel("版本类型", Point.Empty, 110);
            _cboEdition = WorkbenchUi.CreateComboBox(Point.Empty, new Size(190, 32));
            _lblInstallTarget = WorkbenchUi.CreateFieldLabel("安装产品", Point.Empty, 110);
            _cboInstallTarget = WorkbenchUi.CreateComboBox(Point.Empty, new Size(170, 32));
            _lblArchitecture = WorkbenchUi.CreateFieldLabel("位数", Point.Empty, 110);
            _cboArchitecture = WorkbenchUi.CreateComboBox(Point.Empty, new Size(140, 32));
            _lblChannel = WorkbenchUi.CreateFieldLabel("频道", Point.Empty, 110);
            _cboChannel = WorkbenchUi.CreateComboBox(Point.Empty, new Size(170, 32));
            _basicCard.Controls.AddRange(new Control[] { _lblBasicTitle, _lblBasicHint, _basicFlowStrip, _lblEdition, _cboEdition, _lblInstallTarget, _cboInstallTarget, _lblArchitecture, _cboArchitecture, _lblChannel, _cboChannel });

            _optionsCard = WorkbenchUi.CreateCard(Point.Empty, Size.Empty);
            _lblOptionsTitle = WorkbenchUi.CreateSectionTitle("安装选项", new Point(20, 14), 200);
            _chkSilentInstall = WorkbenchUi.CreateCheckBox("静默安装", Point.Empty, 120);
            _chkPinToTaskbar = WorkbenchUi.CreateCheckBox("固定到任务栏", Point.Empty, 140);
            _optionsCard.Controls.AddRange(new Control[] { _lblOptionsTitle, _chkSilentInstall, _chkPinToTaskbar });

            _excludeAppsCard = WorkbenchUi.CreateCard(Point.Empty, Size.Empty);
            _lblExcludeAppsTitle = WorkbenchUi.CreateSectionTitle("排除应用", new Point(18, 14), 180);
            _clbExcludeApps = new CheckedListBox { CheckOnClick = true, Font = new Font("Microsoft YaHei UI", 10F), IntegralHeight = false };
            _excludeAppsHost = WorkbenchUi.CreateInsetPanel(Point.Empty, Size.Empty);
            _excludeAppsCard.Controls.AddRange(new Control[] { _lblExcludeAppsTitle, _excludeAppsHost });
            _excludeAppsHost.Controls.Add(_clbExcludeApps);

            _languagesCard = WorkbenchUi.CreateCard(Point.Empty, Size.Empty);
            _lblLanguagesTitle = WorkbenchUi.CreateSectionTitle("语言", new Point(18, 14), 120);
            _txtLanguageSearch = new TextBox { Font = new Font("Microsoft YaHei UI", 9.5F), PlaceholderText = "搜索语言" };
            _clbLanguages = new CheckedListBox { CheckOnClick = true, Font = new Font("Microsoft YaHei UI", 10F), IntegralHeight = false };
            _languagesHost = WorkbenchUi.CreateInsetPanel(Point.Empty, Size.Empty);
            _languagesCard.Controls.AddRange(new Control[] { _lblLanguagesTitle, _txtLanguageSearch, _languagesHost });
            _languagesHost.Controls.Add(_clbLanguages);

            _precheckCard = WorkbenchUi.CreateCard(Point.Empty, Size.Empty);
            _lblPrecheckTitle = WorkbenchUi.CreateSectionTitle("安装前检查", new Point(20, 14), 200);
            _lblPrecheckStatus = new Label
            {
                Text = "等待检查",
                Font = new Font("Microsoft YaHei UI", 10F),
                ForeColor = _normalColor,
                AutoSize = false,
                Visible = false
            };
            WorkbenchUi.ApplyStatusStyle(_lblPrecheckStatus, WorkbenchTone.Default);
            _txtPrecheckResult = WorkbenchUi.CreateReadonlyTextBox(Point.Empty, Size.Empty);
            _precheckOutputHost = WorkbenchUi.CreateInsetPanel(Point.Empty, Size.Empty);
            _btnRefreshChecks = WorkbenchUi.CreateSecondaryButton("重新检查", Point.Empty, new Size(148, 40));
            _btnRefreshChecks.Click += async (_, _) => await RefreshPrecheckAsync();
            _precheckCard.Controls.AddRange(new Control[] { _lblPrecheckTitle, _lblPrecheckStatus, _precheckOutputHost, _btnRefreshChecks });
            _precheckOutputHost.Controls.Add(_txtPrecheckResult);

            _resultCard = WorkbenchUi.CreateCard(Point.Empty, Size.Empty);
            _lblResultTitle = WorkbenchUi.CreateSectionTitle("处理结果", new Point(20, 14), 200);
            _lblResult = new Label
            {
                Text = "准备就绪",
                Font = new Font("Microsoft YaHei UI", 10F),
                ForeColor = _normalColor,
                AutoSize = false,
                Visible = false
            };
            WorkbenchUi.ApplyStatusStyle(_lblResult, WorkbenchTone.Default);
            _txtResultDetails = WorkbenchUi.CreateReadonlyTextBox(Point.Empty, Size.Empty, "安装结果和执行日志会显示在这里");
            _resultOutputHost = WorkbenchUi.CreateInsetPanel(Point.Empty, Size.Empty);
            _progressBar = new ProgressBar { Style = ProgressBarStyle.Continuous, Visible = false };
            _btnOpenLogs = WorkbenchUi.CreateSecondaryButton("打开日志目录", Point.Empty, new Size(148, 40));
            _btnInstall = WorkbenchUi.CreatePrimaryButton("立即安装", Point.Empty, new Size(148, 40));
            _btnOpenLogs.Click += btnOpenInstallLog_Click;
            _btnInstall.Click += btnInstallOffice365_Click;
            _resultCard.Controls.AddRange(new Control[] { _lblResultTitle, _lblResult, _resultOutputHost, _progressBar, _btnOpenLogs, _btnInstall });
            _resultOutputHost.Controls.Add(_txtResultDetails);

            _contentPanel.Controls.AddRange(new Control[] { _lblTitle, _lblDescription, _basicCard, _optionsCard, _excludeAppsCard, _languagesCard, _precheckCard, _resultCard });
            _scrollHost.Controls.Add(_contentPanel);
            Controls.Add(_scrollHost);

            InitializeControls();
            SetupEventHandlers();
            WorkbenchUi.StyleCheckedListBox(_clbExcludeApps);
            WorkbenchUi.StyleCheckedListBox(_clbLanguages);
            WorkbenchUi.StyleInputTextBox(_txtLanguageSearch);
            _txtPrecheckResult.BorderStyle = BorderStyle.None;
            _txtResultDetails.BorderStyle = BorderStyle.None;
            ApplyLanguage(_language);
            Resize += (_, _) => LayoutControls();
            _scrollHost.Resize += (_, _) => LayoutControls();
            _scrollHost.Scroll += (_, _) => CloseAllDropDowns();
            _scrollHost.MouseWheel += (_, _) => CloseAllDropDowns();
            ParentChanged += (_, _) => this.BeginInvokeWhenReady(PrepareForDisplay);
            VisibleChanged += (_, _) =>
            {
                if (Visible)
                {
                    this.BeginInvokeWhenReady(PrepareForDisplay);
                }
            };
            LayoutControls();
            this.BeginInvokeWhenReady(() => _ = RefreshPrecheckAsync());
        }

        public void PrepareForDisplay()
        {
            if (_scrollHost.IsHandleCreated)
            {
                _scrollHost.AutoScrollPosition = new Point(0, 0);
            }

            LayoutControls();
        }

        private void InitializeControls()
        {
            RebindSelectionCombos();
            _chkPinToTaskbar.Checked = true;

            foreach (string appId in InstallationConfig.GetAvailableApps())
            {
                _clbExcludeApps.Items.Add(new ExcludeAppItem { Display = GetLocalizedAppDisplayName(appId), Value = appId }, false);
            }

            RefreshLanguageList();
            UpdateConfiguration();
        }

        private void SetupEventHandlers()
        {
            _cboEdition.SelectedIndexChanged += (_, _) => UpdateConfiguration();
            _cboInstallTarget.SelectedIndexChanged += (_, _) => UpdateConfiguration();
            _cboArchitecture.SelectedIndexChanged += (_, _) => UpdateConfiguration();
            _cboChannel.SelectedIndexChanged += (_, _) => UpdateConfiguration();
            _chkSilentInstall.CheckedChanged += (_, _) => UpdateConfiguration();
            _chkPinToTaskbar.CheckedChanged += (_, _) => UpdateConfiguration();
            _txtLanguageSearch.TextChanged += (_, _) => RefreshLanguageList();
            _clbExcludeApps.ItemCheck += (_, _) => this.BeginInvokeWhenReady(UpdateConfiguration);
            _clbLanguages.ItemCheck += (_, _) => this.BeginInvokeWhenReady(UpdateConfiguration);
        }

        private void LayoutControls()
        {
            if (_scrollHost.ClientSize.Width <= 0 || _scrollHost.ClientSize.Height <= 0)
            {
                return;
            }

            const int margin = 24;
            const int gap = 18;
            int viewportWidth = _scrollHost.ClientSize.Width;
            int viewportHeight = _scrollHost.ClientSize.Height;
            int availableWidth = Math.Max(700, Math.Min(1160, viewportWidth - margin * 2 - 4));
            int contentLeft = Math.Max(margin, (viewportWidth - availableWidth) / 2);
            int currentTop = 10;

            var fieldPairs = new List<(Control Label, Control Input)>
            {
                (_lblInstallTarget, _cboInstallTarget),
                (_lblArchitecture, _cboArchitecture),
                (_lblChannel, _cboChannel)
            };

            if (_lblEdition.Visible && _cboEdition.Visible)
            {
                fieldPairs.Insert(0, (_lblEdition, _cboEdition));
            }

            int maxColumns = availableWidth >= 980 ? 4 : availableWidth >= 760 ? 2 : 1;
            int fieldColumns = Math.Max(1, Math.Min(maxColumns, fieldPairs.Count));
            int fieldColumnWidth = (availableWidth - 48 - gap * (fieldColumns - 1)) / fieldColumns;
            int fieldRows = (int)Math.Ceiling(fieldPairs.Count / (double)fieldColumns);

            _basicCard.Location = new Point(contentLeft, currentTop);
            _basicCard.Size = new Size(availableWidth, 76 + fieldRows * 74 + 16);
            _lblBasicTitle.Location = new Point(20, 16);
            _lblBasicHint.Location = new Point(20, 44);
            _lblBasicHint.Size = new Size(availableWidth - 40, 20);

            for (int i = 0; i < fieldPairs.Count; i++)
            {
                int row = i / fieldColumns;
                int column = i % fieldColumns;
                int left = 24 + column * (fieldColumnWidth + gap);
                int fieldTop = 78 + row * 74;
                fieldPairs[i].Label.Location = new Point(left, fieldTop);
                fieldPairs[i].Label.Size = new Size(Math.Min(110, fieldColumnWidth), 24);
                fieldPairs[i].Input.Location = new Point(left, fieldTop + 30);
                fieldPairs[i].Input.Size = new Size(fieldColumnWidth, 32);
            }
            currentTop = _basicCard.Bottom + 18;

            _optionsCard.Location = new Point(contentLeft, currentTop);
            _optionsCard.Size = new Size(availableWidth, 86);
            _chkSilentInstall.Location = new Point(22, 46);
            _chkPinToTaskbar.Location = new Point(150, 46);
            currentTop = _optionsCard.Bottom + 18;

            bool isM365Target = _excludeAppsCard.Visible;
            bool stackGroups = !isM365Target || availableWidth < 920;
            if (isM365Target && !stackGroups)
            {
                int groupWidth = (availableWidth - gap) / 2;
                _excludeAppsCard.Location = new Point(contentLeft, currentTop);
                _excludeAppsCard.Size = new Size(groupWidth, 258);
                _lblExcludeAppsTitle.Location = new Point(18, 14);
                _excludeAppsHost.Location = new Point(18, 46);
                _excludeAppsHost.Size = new Size(groupWidth - 36, 190);
                _clbExcludeApps.Location = new Point(10, 10);
                _clbExcludeApps.Size = new Size(_excludeAppsHost.Width - 20, _excludeAppsHost.Height - 20);

                _languagesCard.Location = new Point(contentLeft + groupWidth + gap, currentTop);
                _languagesCard.Size = new Size(groupWidth, 258);
                _lblLanguagesTitle.Location = new Point(18, 14);
                _txtLanguageSearch.Location = new Point(18, 44);
                _txtLanguageSearch.Size = new Size(groupWidth - 36, 30);
                _languagesHost.Location = new Point(18, 82);
                _languagesHost.Size = new Size(groupWidth - 36, 154);
                _clbLanguages.Location = new Point(10, 10);
                _clbLanguages.Size = new Size(_languagesHost.Width - 20, _languagesHost.Height - 20);
                currentTop = _languagesCard.Bottom + 18;
            }
            else
            {
                if (isM365Target)
                {
                    _excludeAppsCard.Location = new Point(contentLeft, currentTop);
                    _excludeAppsCard.Size = new Size(availableWidth, 258);
                    _lblExcludeAppsTitle.Location = new Point(18, 14);
                    _excludeAppsHost.Location = new Point(18, 46);
                    _excludeAppsHost.Size = new Size(availableWidth - 36, 190);
                    _clbExcludeApps.Location = new Point(10, 10);
                    _clbExcludeApps.Size = new Size(_excludeAppsHost.Width - 20, _excludeAppsHost.Height - 20);
                    currentTop = _excludeAppsCard.Bottom + 18;
                }

                _languagesCard.Location = new Point(contentLeft, currentTop);
                _languagesCard.Size = new Size(availableWidth, 258);
                _lblLanguagesTitle.Location = new Point(18, 14);
                _txtLanguageSearch.Location = new Point(18, 44);
                _txtLanguageSearch.Size = new Size(availableWidth - 36, 30);
                _languagesHost.Location = new Point(18, 82);
                _languagesHost.Size = new Size(availableWidth - 36, 154);
                _clbLanguages.Location = new Point(10, 10);
                _clbLanguages.Size = new Size(_languagesHost.Width - 20, _languagesHost.Height - 20);
                currentTop = _languagesCard.Bottom + 18;
            }

            _precheckCard.Location = new Point(contentLeft, currentTop);
            _precheckCard.Size = new Size(availableWidth, 244);
            const int precheckHeaderHeight = 58;
            _btnRefreshChecks.Location = new Point(availableWidth - _btnRefreshChecks.Width - 20, (precheckHeaderHeight - _btnRefreshChecks.Height) / 2);
            _precheckOutputHost.Location = new Point(20, precheckHeaderHeight);
            _precheckOutputHost.Size = new Size(availableWidth - 40, 162);
            _txtPrecheckResult.Location = new Point(10, 10);
            _txtPrecheckResult.Size = new Size(_precheckOutputHost.Width - 20, _precheckOutputHost.Height - 20);
            currentTop = _precheckCard.Bottom + 18;

            int minResultHeight = _isInstalling ? 338 : 318;
            int resultHeight = Math.Max(minResultHeight, viewportHeight - currentTop - 28);
            _resultCard.Location = new Point(contentLeft, currentTop);
            _resultCard.Size = new Size(availableWidth, resultHeight);

            int footerHeight = 40;
            int footerGap = 16;
            int detailsTop = 58;
            int detailsBottom = _resultCard.Height - 20 - footerHeight - footerGap;
            int detailsHeight = Math.Max(162, detailsBottom - detailsTop);
            _resultOutputHost.Location = new Point(20, detailsTop);
            _resultOutputHost.Size = new Size(availableWidth - 40, detailsHeight);
            _txtResultDetails.Location = new Point(10, 10);
            _txtResultDetails.Size = new Size(_resultOutputHost.Width - 20, _resultOutputHost.Height - 20);

            int footerTop = _resultOutputHost.Bottom + footerGap;
            _btnInstall.Location = new Point(availableWidth - _btnInstall.Width - 20, footerTop);
            _btnOpenLogs.Location = new Point(_btnInstall.Left - _btnOpenLogs.Width - 12, footerTop);
            int progressRight = _btnOpenLogs.Left - 16;
            _progressBar.Location = new Point(20, footerTop + (_btnInstall.Height - 10) / 2);
            _progressBar.Size = new Size(Math.Max(180, progressRight - 20), 10);
            _progressBar.Visible = _isInstalling;

            _contentPanel.Width = Math.Max(10, viewportWidth - 1);
            _contentPanel.Height = _resultCard.Bottom + 28;
        }

        private void UpdateConfiguration()
        {
            _cboEdition.DroppedDown = false;
            _cboInstallTarget.DroppedDown = false;
            _basicCard.SuspendLayout();
            bool isM365Target = _cboInstallTarget.SelectedIndex <= 0;
            _lblEdition.Visible = isM365Target;
            _cboEdition.Visible = isM365Target;
            _excludeAppsCard.Visible = isM365Target;

            _currentConfig.Edition = _cboEdition.SelectedIndex == 0 ? OfficeEdition.Business : OfficeEdition.Enterprise;
            _currentConfig.Architecture = _cboArchitecture.SelectedIndex == 0 ? OfficeArchitecture.x64 : OfficeArchitecture.x86;
            _currentConfig.Channel = _cboChannel.SelectedIndex == 0 ? UpdateChannel.Monthly : UpdateChannel.Broad;
            _currentConfig.SilentInstall = _chkSilentInstall.Checked;
            _currentConfig.DisplayInstall = !_chkSilentInstall.Checked;
            _currentConfig.EnableUpdates = true;
            _currentConfig.PinItemsToTaskbar = _chkPinToTaskbar.Checked;
            _currentConfig.IncludeProject = _cboInstallTarget.SelectedIndex == 1;
            _currentConfig.IncludeVisio = _cboInstallTarget.SelectedIndex == 2;
            _currentConfig.DownloadPath = @"C:\Office365Install";
            _currentConfig.ExcludeApps = GetSelectedExcludeApps();
            _currentConfig.LanguageIDs = GetSelectedLanguages();
            LayoutControls();
            _basicCard.ResumeLayout();
            _basicCard.Invalidate(true);
        }

        private async System.Threading.Tasks.Task RefreshPrecheckAsync()
        {
            _btnRefreshChecks.Enabled = false;
            _lblPrecheckStatus.Text = T("正在检查系统环境...", "Checking system environment...");
            WorkbenchUi.ApplyStatusStyle(_lblPrecheckStatus, WorkbenchTone.Info);
            _txtPrecheckResult.Text = T("正在检查系统环境...", "Checking system environment...");
            try
            {
                ValidationReport report = await _validationService.GetValidationReportAsync(_language);
                _lblPrecheckStatus.Text = report.HasBlockingItems
                    ? T("存在阻止安装的检查项", "Blocking items were found")
                    : report.HasAttentionItems
                        ? T("发现需要关注的检查项", "Items requiring attention were found")
                        : T("检查完成", "Checks done");
                WorkbenchUi.ApplyStatusStyle(
                    _lblPrecheckStatus,
                    report.HasBlockingItems
                        ? WorkbenchTone.Danger
                        : report.HasAttentionItems
                            ? WorkbenchTone.Warning
                            : WorkbenchTone.Success);
                _txtPrecheckResult.Text = string.Join(
                    Environment.NewLine,
                    report.Items.Select(item => $"{GetValidationItemPrefix(item)} {item.Name}: {item.Detail}"));
            }
            catch (Exception ex)
            {
                _lblPrecheckStatus.Text = T("检查失败", "Check failed");
                WorkbenchUi.ApplyStatusStyle(_lblPrecheckStatus, WorkbenchTone.Danger);
                _txtPrecheckResult.Text = T($"环境检查失败：{ex.Message}", $"Environment check failed: {ex.Message}");
            }
            finally
            {
                _btnRefreshChecks.Enabled = !_isInstalling;
                LayoutControls();
            }
        }

        private string GetValidationItemPrefix(ValidationCheckResult item)
        {
            if (item.IsBlocking)
            {
                return T("[失败]", "[Fail]");
            }

            return item.RequiresAttention ? T("[风险]", "[Risk]") : T("[通过]", "[Pass]");
        }

        private async void btnInstallOffice365_Click(object? sender, EventArgs e)
        {
            if (_isInstalling) return;
            UpdateConfiguration();

            string targetName = GetSelectedInstallTargetDisplayName();
            string editionText = _currentConfig.IncludeProject || _currentConfig.IncludeVisio
                ? T("不适用", "N/A")
                : (_currentConfig.Edition == OfficeEdition.Business ? T("商业版", "Business") : T("企业版", "Enterprise"));
            string architectureText = _currentConfig.Architecture == OfficeArchitecture.x64 ? T("64位", "64-bit") : T("32位", "32-bit");
            string channelText = _currentConfig.Channel == UpdateChannel.Monthly ? T("当前频道", "Current Channel") : T("半年频道", "Semi-Annual Channel");
            string languageText = _currentConfig.LanguageIDs.Count > 0 ? string.Join(_language == UiLanguage.Chinese ? "、" : ", ", _currentConfig.LanguageIDs) : T("匹配系统语言", "Match system language");
            string excludeAppsText = _excludeAppsCard.Visible && _currentConfig.ExcludeApps.Count > 0
                ? string.Join(_language == UiLanguage.Chinese ? "、" : ", ", _currentConfig.ExcludeApps.Select(GetLocalizedAppDisplayName))
                : (_excludeAppsCard.Visible ? T("无", "None") : T("不适用", "N/A"));

            DialogResult confirm = MessageBox.Show(
                T(
                    $"请确认当前安装配置：\n\n安装产品：{targetName}\n版本类型：{editionText}\n位数：{architectureText}\n频道：{channelText}\n安装方式：{(_currentConfig.SilentInstall ? "静默安装" : "显示安装界面")}\n固定到任务栏：{(_currentConfig.PinItemsToTaskbar ? "是" : "否")}\n语言：{languageText}\n排除应用：{excludeAppsText}\n\n是否开始在线安装？",
                    $"Please confirm installation configuration:\n\nProduct: {targetName}\nEdition: {editionText}\nArchitecture: {architectureText}\nChannel: {channelText}\nInstall mode: {(_currentConfig.SilentInstall ? "Silent install" : "Show installer UI")}\nPin to taskbar: {(_currentConfig.PinItemsToTaskbar ? "Yes" : "No")}\nLanguage: {languageText}\nExclude apps: {excludeAppsText}\n\nStart online installation now?"),
                T("确认安装", "Confirm installation"),
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (confirm != DialogResult.Yes) return;

            try
            {
                _isInstalling = true;
                UpdateUiState();
                _progressBar.Visible = true;
                _progressBar.Value = 0;
                _lblResult.Text = T("正在准备安装...", "Preparing installation...");
                WorkbenchUi.ApplyStatusStyle(_lblResult, WorkbenchTone.Info);
                _txtResultDetails.Text = T("正在准备安装...\r\n", "Preparing installation...\r\n");
                LayoutControls();

                InstallationExecutionResult result = await _installationService.InstallOfficeAsync(_currentConfig, _currentConfig.ExcludeApps, _language);
                _lastLogFilePath = result.LogFilePath;

                if (!result.Success)
                {
                    _lblResult.Text = T("安装失败", "Install failed");
                    WorkbenchUi.ApplyStatusStyle(_lblResult, WorkbenchTone.Danger);
                    _txtResultDetails.Text = T(
                        $"安装失败：{result.Message}\r\n\r\n日志：\r\n{result.LogFilePath}",
                        $"Install failed: {result.Message}\r\n\r\nLog:\r\n{result.LogFilePath}");
                    MessageBox.Show(
                        T($"安装过程中发生异常：{result.Message}\n\n安装日志：\n{result.LogFilePath}", $"Exception occurred during installation: {result.Message}\n\nInstall log:\n{result.LogFilePath}"),
                        T("失败", "Failed"),
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                _progressBar.Value = _progressBar.Maximum;
                _lblResult.Text = T("安装任务已启动完成", "Install task started successfully");
                WorkbenchUi.ApplyStatusStyle(_lblResult, WorkbenchTone.Success);
                _txtResultDetails.Text = T(
                    $"安装任务已启动完成。\r\n\r\n日志：\r\n{result.LogFilePath}",
                    $"Install task started successfully.\r\n\r\nLog:\r\n{result.LogFilePath}");
                MessageBox.Show(
                    T($"安装任务已启动完成。\n\n安装日志：\n{result.LogFilePath}", $"Install task started successfully.\n\nInstall log:\n{result.LogFilePath}"),
                    T("完成", "Done"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                _lblResult.Text = T("安装失败", "Install failed");
                WorkbenchUi.ApplyStatusStyle(_lblResult, WorkbenchTone.Danger);
                _txtResultDetails.Text = T($"执行失败：{ex.Message}", $"Execution failed: {ex.Message}");
                MessageBox.Show(
                    T($"执行失败：{ex.Message}", $"Execution failed: {ex.Message}"),
                    T("错误", "Error"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                _progressBar.Visible = false;
                _isInstalling = false;
                UpdateUiState();
                LayoutControls();
            }
        }

        public void ApplyLanguage(UiLanguage language)
        {
            bool hadAppliedLanguage = _hasAppliedLanguage;
            bool shouldRebindLists = !hadAppliedLanguage || _language != language;
            _language = language;

            _lblTitle.Text = T("安装 M365 应用", "Install M365 Apps");
            _lblTitle.Visible = false;
            _lblDescription.Text = string.Empty;
            _lblDescription.Visible = false;

            _lblBasicTitle.Text = T("安装参数", "Install Parameters");
            _lblBasicHint.Text = T(
                "先确定安装对象、位数与频道，再执行预检查和安装。",
                "Set product, architecture and channel first, then run pre-checks and installation.");
            _lblBasicFlowText.Text = T(
                "先配置基础参数，再选择排除应用与语言，最后查看检查结果并执行安装。",
                "Configure the base parameters first, then choose excluded apps and languages before reviewing checks and starting the install.");
            _lblEdition.Text = T("版本类型", "Edition");
            _lblInstallTarget.Text = T("安装产品", "Product");
            _lblArchitecture.Text = T("位数", "Architecture");
            _lblChannel.Text = T("频道", "Channel");
            _lblOptionsTitle.Text = T("安装选项", "Install Options");
            _lblExcludeAppsTitle.Text = T("排除应用", "Exclude Apps");
            _lblLanguagesTitle.Text = T("语言", "Language");
            _lblPrecheckTitle.Text = T("安装前检查", "Pre-check");
            _lblResultTitle.Text = T("处理结果", "Results");

            _chkSilentInstall.Text = T("静默安装", "Silent install");
            _chkPinToTaskbar.Text = T("固定到任务栏", "Pin to taskbar");
            _btnRefreshChecks.Text = T("重新检查", "Recheck");
            _btnOpenLogs.Text = T("打开日志目录", "Open Log Folder");
            _btnInstall.Text = T("立即安装", "Install Now");
            _txtLanguageSearch.PlaceholderText = T("搜索语言", "Search language");

            if (!_isInstalling)
            {
                _lblResult.Text = T("准备就绪", "Ready");
                WorkbenchUi.ApplyStatusStyle(_lblResult, WorkbenchTone.Default);
                if (string.IsNullOrWhiteSpace(_txtResultDetails.Text) ||
                    _txtResultDetails.Text.Contains("安装结果", StringComparison.OrdinalIgnoreCase) ||
                    _txtResultDetails.Text.Contains("Installation result", StringComparison.OrdinalIgnoreCase))
                {
                    _txtResultDetails.Text = T("安装结果和执行日志会显示在这里", "Installation result and logs will be shown here");
                }
            }

            if (shouldRebindLists)
            {
                RebindSelectionCombos();
                RefreshExcludeAppDisplay();
                RefreshLanguageList();
                _hasAppliedLanguage = true;
                if (hadAppliedLanguage)
                {
                    this.BeginInvokeWhenReady(() => _ = RefreshPrecheckAsync());
                }
            }
        }

        private void RebindSelectionCombos()
        {
            int editionIndex = _cboEdition.SelectedIndex < 0 ? 0 : _cboEdition.SelectedIndex;
            int installTargetIndex = _cboInstallTarget.SelectedIndex < 0 ? 0 : _cboInstallTarget.SelectedIndex;
            int architectureIndex = _cboArchitecture.SelectedIndex < 0 ? 0 : _cboArchitecture.SelectedIndex;
            int channelIndex = _cboChannel.SelectedIndex < 0 ? 0 : _cboChannel.SelectedIndex;

            BindCombo(_cboEdition, new[]
            {
                new OptionItem { Key = "business", Display = T("商业版", "Business") },
                new OptionItem { Key = "enterprise", Display = T("企业版", "Enterprise") }
            }, editionIndex);

            BindCombo(_cboInstallTarget, new[]
            {
                new OptionItem { Key = "m365", Display = T("M365 应用", "M365 Apps") },
                new OptionItem { Key = "project", Display = "Project" },
                new OptionItem { Key = "visio", Display = "Visio" }
            }, installTargetIndex);

            BindCombo(_cboArchitecture, new[]
            {
                new OptionItem { Key = "x64", Display = T("64位", "64-bit") },
                new OptionItem { Key = "x86", Display = T("32位", "32-bit") }
            }, architectureIndex);

            BindCombo(_cboChannel, new[]
            {
                new OptionItem { Key = "monthly", Display = T("当前频道", "Current Channel") },
                new OptionItem { Key = "broad", Display = T("半年频道", "Semi-Annual Channel") }
            }, channelIndex);
        }

        private static void BindCombo(ComboBox comboBox, IReadOnlyList<OptionItem> items, int selectedIndex)
        {
            comboBox.BeginUpdate();
            comboBox.Items.Clear();
            foreach (OptionItem item in items)
            {
                comboBox.Items.Add(item);
            }

            comboBox.SelectedIndex = items.Count == 0
                ? -1
                : Math.Max(0, Math.Min(selectedIndex, items.Count - 1));
            comboBox.EndUpdate();
        }

        private void RefreshExcludeAppDisplay()
        {
            var checkedValues = new HashSet<string>(
                _clbExcludeApps.CheckedItems.Cast<object>()
                    .OfType<ExcludeAppItem>()
                    .Select(i => i.Value),
                StringComparer.OrdinalIgnoreCase);

            _clbExcludeApps.BeginUpdate();
            _clbExcludeApps.Items.Clear();
            foreach (string appId in InstallationConfig.GetAvailableApps())
            {
                int index = _clbExcludeApps.Items.Add(
                    new ExcludeAppItem { Display = GetLocalizedAppDisplayName(appId), Value = appId },
                    false);
                if (checkedValues.Contains(appId))
                {
                    _clbExcludeApps.SetItemChecked(index, true);
                }
            }
            _clbExcludeApps.EndUpdate();
        }

        private void btnOpenInstallLog_Click(object? sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(_lastLogFilePath) && File.Exists(_lastLogFilePath))
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = "explorer.exe", Arguments = $"/select,\"{_lastLogFilePath}\"", UseShellExecute = true });
                    return;
                }

                string logsDirectory = ToolPathService.GetModuleRootDirectory("Install");
                if (!Directory.Exists(logsDirectory))
                {
                    throw new DirectoryNotFoundException(T($"未找到日志目录：{logsDirectory}", $"Log folder not found: {logsDirectory}"));
                }

                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = "explorer.exe", Arguments = $"\"{logsDirectory}\"", UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    T($"打开日志目录失败：{ex.Message}", $"Failed to open log folder: {ex.Message}"),
                    T("错误", "Error"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void OnInstallationProgressChanged(object? sender, InstallationProgressEventArgs e)
        {
            if (IsDisposed || Disposing || !IsHandleCreated)
            {
                return;
            }

            if (InvokeRequired)
            {
                if (!this.TryBeginInvokeIfReady(() => OnInstallationProgressChanged(sender, e)))
                {
                    return;
                }
                return;
            }

            _lblResult.Text = e.Message;
            _lblResult.ForeColor = e.Progress >= 100 ? _successColor : WorkbenchUi.PrimaryColor;
            if (e.Progress >= 0) _progressBar.Value = Math.Min(e.Progress, _progressBar.Maximum);
            AppendResultLine(e.Message);
        }

        private void UpdateUiState()
        {
            bool enabled = !_isInstalling;
            _cboEdition.Enabled = enabled;
            _cboInstallTarget.Enabled = enabled;
            _cboArchitecture.Enabled = enabled;
            _cboChannel.Enabled = enabled;
            _chkSilentInstall.Enabled = enabled;
            _chkPinToTaskbar.Enabled = enabled;
            _clbExcludeApps.Enabled = enabled;
            _txtLanguageSearch.Enabled = enabled;
            _clbLanguages.Enabled = enabled;
            _btnRefreshChecks.Enabled = enabled;
            _btnInstall.Enabled = enabled;
            _btnOpenLogs.Enabled = true;
        }

        private List<string> GetSelectedExcludeApps()
        {
            var result = new List<string>();
            if (!_excludeAppsCard.Visible) return result;
            foreach (object checkedItem in _clbExcludeApps.CheckedItems)
            {
                if (checkedItem is ExcludeAppItem item) result.Add(item.Value);
            }
            return result;
        }

        private List<string> GetSelectedLanguages()
        {
            var result = new List<string>();
            foreach (object checkedItem in _clbLanguages.CheckedItems)
            {
                string value = GetLanguageValue(checkedItem);
                if (!string.IsNullOrWhiteSpace(value)) result.Add(value);
            }
            return result;
        }

        private void RefreshLanguageList()
        {
            string searchText = _txtLanguageSearch.Text.Trim();
            var selectedLanguages = new HashSet<string>(_currentConfig.LanguageIDs, StringComparer.OrdinalIgnoreCase);
            foreach (object checkedItem in _clbLanguages.CheckedItems) selectedLanguages.Add(GetLanguageValue(checkedItem));

            _clbLanguages.Items.Clear();
            string systemLanguage = GetPreferredSystemLanguage();
            List<string> languages = SupportedLanguages.Select(item => item.Code).ToList();
            if (!string.IsNullOrWhiteSpace(systemLanguage))
            {
                languages.RemoveAll(language => string.Equals(language, systemLanguage, StringComparison.OrdinalIgnoreCase));
                languages.Insert(0, systemLanguage);
            }

            foreach (string language in languages)
            {
                string display = BuildLanguageDisplayName(language);
                bool matchesSearch = string.IsNullOrWhiteSpace(searchText) || language.Contains(searchText, StringComparison.OrdinalIgnoreCase) || display.Contains(searchText, StringComparison.OrdinalIgnoreCase);
                if (!matchesSearch) continue;
                _clbLanguages.Items.Add(new LanguageItem { Value = language, Display = display });
            }

            for (int i = 0; i < _clbLanguages.Items.Count; i++)
            {
                string language = GetLanguageValue(_clbLanguages.Items[i]);
                bool isChecked = selectedLanguages.Count == 0 ? IsSystemPreferredLanguage(language) : selectedLanguages.Contains(language);
                _clbLanguages.SetItemChecked(i, isChecked);
            }

            if (_clbLanguages.Items.Count > 0) _clbLanguages.TopIndex = 0;
        }

        private string GetSelectedInstallTargetDisplayName() => _cboInstallTarget.SelectedIndex switch
        {
            1 => "Project",
            2 => "Visio",
            _ => T("M365 应用", "M365 Apps")
        };
        private static string GetLanguageValue(object? item) => item is LanguageItem languageItem ? languageItem.Value : item?.ToString() ?? string.Empty;

        private string GetLocalizedAppDisplayName(string appId)
        {
            if (string.Equals(appId, "OutlookForWindows", StringComparison.OrdinalIgnoreCase))
            {
                return T("新版 Outlook", "New Outlook");
            }

            return InstallationConfig.GetAppDisplayName(appId);
        }

        private static string GetPreferredSystemLanguage()
        {
            string[] candidates = { CultureInfo.InstalledUICulture.Name, CultureInfo.CurrentUICulture.Name, CultureInfo.CurrentCulture.Name };
            foreach (string candidate in candidates)
            {
                if (!string.IsNullOrWhiteSpace(candidate)) return candidate;
            }
            return "zh-CN";
        }

        private static bool IsSystemPreferredLanguage(string language)
        {
            string systemLanguage = GetPreferredSystemLanguage();
            if (string.Equals(language, systemLanguage, StringComparison.OrdinalIgnoreCase)) return true;
            try
            {
                string systemTwoLetter = CultureInfo.GetCultureInfo(systemLanguage).TwoLetterISOLanguageName;
                string languageTwoLetter = CultureInfo.GetCultureInfo(language).TwoLetterISOLanguageName;
                return string.Equals(systemTwoLetter, languageTwoLetter, StringComparison.OrdinalIgnoreCase);
            }
            catch (CultureNotFoundException)
            {
                return false;
            }
        }

        private string BuildLanguageDisplayName(string language)
        {
            string name = GetLanguageName(language);
            string suffix = IsSystemPreferredLanguage(language)
                ? T("（系统默认）", " (System default)")
                : string.Empty;
            return $"{language} - {name}{suffix}";
        }

        private string GetLanguageName(string language)
        {
            foreach ((string code, string chineseName, string englishName) in SupportedLanguages)
            {
                if (string.Equals(code, language, StringComparison.OrdinalIgnoreCase))
                {
                    return _language == UiLanguage.English ? englishName : chineseName;
                }
            }

            try
            {
                CultureInfo culture = CultureInfo.GetCultureInfo(language);
                return _language == UiLanguage.English ? culture.EnglishName : culture.NativeName;
            }
            catch (CultureNotFoundException)
            {
                return T("未命名语言", "Unnamed language");
            }
        }

        private void AppendResultLine(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return;
            }

            string normalized = message.Trim();
            string current = _txtResultDetails.Text?.Trim() ?? string.Empty;
            if (current.EndsWith(normalized, StringComparison.Ordinal))
            {
                return;
            }

            _txtResultDetails.AppendText((string.IsNullOrWhiteSpace(_txtResultDetails.Text) ? string.Empty : Environment.NewLine) + normalized);
        }

        private void CloseAllDropDowns()
        {
            _cboEdition.DroppedDown = false;
            _cboInstallTarget.DroppedDown = false;
            _cboArchitecture.DroppedDown = false;
            _cboChannel.DroppedDown = false;
        }

        private string T(string zh, string en) => LocalizationService.Localize(_language, zh, en);
    }
}
