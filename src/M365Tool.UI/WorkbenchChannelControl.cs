using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Office365CleanupTool.Models;
using Office365CleanupTool.Services;

namespace Office365CleanupTool
{
    public class WorkbenchChannelControl : UserControl
    {
        private sealed class ChannelOptionView
        {
            public OfficeChannelOption Raw { get; set; } = new();
            public string Display { get; set; } = string.Empty;
            public override string ToString() => Display;
        }

        private sealed class VersionOptionView
        {
            public OfficeVersionOption Raw { get; set; } = new();
            public string Display { get; set; } = string.Empty;
            public override string ToString() => Display;
        }

        private readonly OfficeChannelService _officeChannelService;
        private readonly Label _lblTitle;
        private readonly Label _lblDescription;
        private readonly Panel _settingsCard;
        private readonly Panel _resultCard;
        private readonly Panel _resultHost;
        private readonly Label _lblSettingsTitle;
        private readonly Label _lblChannelTitle;
        private readonly Label _lblVersionTitle;
        private readonly Label _lblResultTitle;
        private readonly Label _lblResultStatus;
        private readonly ComboBox _cboChannel;
        private readonly ComboBox _cboVersion;
        private readonly CheckBox _chkDisableUpdates;
        private readonly Label _lblLoading;
        private readonly TextBox _txtResult;
        private readonly Button _btnApply;

        private bool _isRunning;
        private bool _suppressChannelSelectionChanged;
        private bool _hasLoadedVersions;
        private bool _isLoadingVersions;
        private int _versionLoadId;
        private UiLanguage _language;
        private List<OfficeVersionOption> _loadedVersions = new();
        private string _versionSummaryText = string.Empty;

        public WorkbenchChannelControl(OfficeChannelService officeChannelService)
        {
            AutoScaleMode = AutoScaleMode.None;
            _officeChannelService = officeChannelService;
            _language = LocalizationService.ResolveLanguage("Auto");

            Dock = DockStyle.Fill;
            BackColor = WorkbenchUi.CanvasColor;
            AutoScroll = true;

            _lblTitle = new Label { Font = WorkbenchUi.CreateUiFont(17F, FontStyle.Bold), ForeColor = WorkbenchUi.PrimaryTextColor, BackColor = Color.Transparent };
            _lblDescription = new Label { Font = WorkbenchUi.CreateUiFont(10.25F), ForeColor = WorkbenchUi.SecondaryTextColor, BackColor = Color.Transparent };

            _settingsCard = WorkbenchUi.CreateCard(Point.Empty, Size.Empty);
            _lblSettingsTitle = WorkbenchUi.CreateSectionTitle(string.Empty, new Point(20, 14));
            _lblChannelTitle = WorkbenchUi.CreateFieldLabel(string.Empty, new Point(20, 48));
            _cboChannel = WorkbenchUi.CreateComboBox(Point.Empty, new Size(320, 32));
            _lblVersionTitle = WorkbenchUi.CreateFieldLabel(string.Empty, new Point(372, 48));
            _cboVersion = WorkbenchUi.CreateComboBox(Point.Empty, new Size(420, 32));
            _lblLoading = new Label { Font = WorkbenchUi.CreateUiFont(9F), ForeColor = Color.FromArgb(100, 100, 100), BackColor = Color.Transparent, Visible = false };
            _chkDisableUpdates = WorkbenchUi.CreateCheckBox(string.Empty, Point.Empty, 220);
            _settingsCard.Controls.AddRange(new Control[] { _lblSettingsTitle, _lblChannelTitle, _cboChannel, _lblVersionTitle, _cboVersion, _lblLoading, _chkDisableUpdates });

            _resultCard = WorkbenchUi.CreateCard(Point.Empty, Size.Empty);
            _lblResultTitle = WorkbenchUi.CreateSectionTitle(string.Empty, new Point(20, 14));
            _lblResultStatus = new Label { AutoSize = false, TextAlign = ContentAlignment.MiddleLeft };
            WorkbenchUi.ApplyStatusStyle(_lblResultStatus, WorkbenchTone.Info);
            _lblResultStatus.Visible = false;
            _resultHost = WorkbenchUi.CreateInsetPanel(Point.Empty, Size.Empty);
            _txtResult = WorkbenchUi.CreateReadonlyTextBox(Point.Empty, Size.Empty, string.Empty);
            _txtResult.BorderStyle = BorderStyle.None;
            _txtResult.BackColor = Color.White;
            _btnApply = WorkbenchUi.CreatePrimaryButton(string.Empty, Point.Empty, new Size(138, 42));
            _btnApply.Click += btnApply_Click;
            _resultHost.Controls.Add(_txtResult);
            _resultCard.Controls.AddRange(new Control[] { _lblResultTitle, _lblResultStatus, _resultHost, _btnApply });

            Controls.AddRange(new Control[] { _lblTitle, _lblDescription, _settingsCard, _resultCard });

            _cboChannel.SelectedIndexChanged += async (_, _) =>
            {
                if (_suppressChannelSelectionChanged)
                {
                    return;
                }

                await LoadVersionsAsync();
            };

            Resize += (_, _) => LayoutControls();
            Scroll += (_, _) => CloseAllDropDowns();
            MouseWheel += (_, _) => CloseAllDropDowns();

            ApplyLanguage(_language);
            LayoutControls();
        }

        public void ApplyLanguage(UiLanguage language)
        {
            _language = language;

            _lblTitle.Text = T("更新频道", "Update Channel");
            _lblTitle.Visible = false;
            _lblDescription.Text = string.Empty;
            _lblDescription.Visible = false;
            _lblSettingsTitle.Text = T("频道设置", "Channel Settings");
            _lblChannelTitle.Text = T("目标频道", "Target Channel");
            _lblVersionTitle.Text = T("目标版本", "Target Version");
            _lblResultTitle.Text = T("处理结果", "Results");
            _lblLoading.Text = T("正在加载 Microsoft 官方版本列表...", "Loading official Microsoft version list...");
            _chkDisableUpdates.Text = T("切换后禁用自动更新", "Disable automatic updates after switching");
            _btnApply.Text = T("开始切换", "Start Switch");

            RebuildChannelItems();
            RebuildVersionItems();
            RefreshResultTextForLanguage();
            UpdateUiState();
            LayoutControls();
        }

        public void PrepareForDisplay()
        {
            if (IsHandleCreated)
            {
                AutoScrollPosition = new Point(0, 0);
            }

            LayoutControls();

            if (!_hasLoadedVersions)
            {
                _hasLoadedVersions = true;
                this.BeginInvokeWhenReady(async () => await LoadVersionsAsync());
            }
        }

        private void RebuildChannelItems()
        {
            UpdateChannel? selected = (_cboChannel.SelectedItem as ChannelOptionView)?.Raw.Channel;
            _suppressChannelSelectionChanged = true;
            _cboChannel.BeginUpdate();
            _cboChannel.Items.Clear();

            foreach (OfficeChannelOption channel in _officeChannelService.GetSupportedChannels(_language))
            {
                _cboChannel.Items.Add(new ChannelOptionView
                {
                    Raw = channel,
                    Display = channel.DisplayName
                });
            }

            if (_cboChannel.Items.Count > 0)
            {
                int index = 0;
                if (selected.HasValue)
                {
                    for (int i = 0; i < _cboChannel.Items.Count; i++)
                    {
                        if (_cboChannel.Items[i] is ChannelOptionView view && view.Raw.Channel == selected.Value)
                        {
                            index = i;
                            break;
                        }
                    }
                }

                _cboChannel.SelectedIndex = Math.Max(0, Math.Min(index, _cboChannel.Items.Count - 1));
            }

            _cboChannel.EndUpdate();
            _suppressChannelSelectionChanged = false;
        }

        private void RebuildVersionItems()
        {
            string? selectedBuild = (_cboVersion.SelectedItem as VersionOptionView)?.Raw.Build;
            _cboVersion.BeginUpdate();
            _cboVersion.Items.Clear();

            foreach (OfficeVersionOption version in _loadedVersions)
            {
                _cboVersion.Items.Add(new VersionOptionView
                {
                    Raw = version,
                    Display = GetVersionDisplayName(version)
                });
            }

            if (_cboVersion.Items.Count > 0)
            {
                int selectedIndex = 0;
                if (!string.IsNullOrWhiteSpace(selectedBuild))
                {
                    for (int i = 0; i < _cboVersion.Items.Count; i++)
                    {
                        if (_cboVersion.Items[i] is VersionOptionView view &&
                            string.Equals(view.Raw.Build, selectedBuild, StringComparison.OrdinalIgnoreCase))
                        {
                            selectedIndex = i;
                            break;
                        }
                    }
                }

                _cboVersion.SelectedIndex = selectedIndex;
            }

            _cboVersion.EndUpdate();
            _cboVersion.Enabled = !_isRunning && !_isLoadingVersions && _cboVersion.Items.Count > 0;
        }

        private void LayoutControls()
        {
            int viewportWidth = ClientSize.Width;
            int viewportHeight = ClientSize.Height;
            float scale = WorkbenchUi.GetAdaptiveUiScale(viewportWidth, viewportHeight);
            int S(int value) => WorkbenchUi.Scale(value, scale);

            WorkbenchUi.ApplyAdaptiveFonts(this, scale);

            int margin = WorkbenchUi.GetAdaptivePageMargin(viewportWidth);
            int availableWidth = WorkbenchUi.GetAdaptiveContentWidth(
                viewportWidth,
                560,
                WorkbenchUi.DefaultContentMaxWidth,
                margin,
                scrollbarAllowance: SystemInformation.VerticalScrollBarWidth);
            Point scrollOffset = AutoScrollPosition;
            int contentLeft = WorkbenchUi.GetAdaptiveContentLeft(viewportWidth, availableWidth, margin) + scrollOffset.X;

            bool stacked = availableWidth < 860;
            _settingsCard.Location = new Point(contentLeft, S(18) + scrollOffset.Y);
            _settingsCard.Size = new Size(availableWidth, stacked ? S(240) : S(194));
            _lblSettingsTitle.Location = new Point(S(20), S(14));
            _lblSettingsTitle.Size = new Size(availableWidth - S(40), S(28));
            _lblChannelTitle.Location = new Point(S(20), S(48));
            _lblChannelTitle.Size = new Size(S(160), S(24));
            _cboChannel.Location = new Point(S(20), S(78));
            _cboChannel.Size = new Size(stacked ? availableWidth - S(40) : Math.Min(S(320), availableWidth - S(40)), S(32));
            _cboVersion.Location = stacked ? new Point(S(20), S(146)) : new Point(Math.Min(S(372), availableWidth / 2), S(78));
            _cboVersion.Size = new Size(
                stacked
                    ? availableWidth - S(40)
                    : Math.Max(S(260), availableWidth - _cboVersion.Left - S(20)),
                S(32));
            _lblVersionTitle.Location = stacked ? new Point(S(20), S(116)) : new Point(_cboVersion.Left, S(48));
            _lblVersionTitle.Size = new Size(S(180), S(24));
            _lblLoading.Location = new Point(S(20), stacked ? S(188) : S(120));
            _lblLoading.Size = new Size(Math.Min(availableWidth - S(40), S(520)), S(24));
            _chkDisableUpdates.Location = new Point(S(20), stacked ? S(210) : S(150));
            _chkDisableUpdates.Size = new Size(Math.Min(availableWidth - S(40), S(360)), S(28));

            _resultCard.Location = new Point(contentLeft, _settingsCard.Bottom + S(18));
            bool stackButton = availableWidth < 820;
            _btnApply.Size = WorkbenchUi.ScaleSize(new Size(138, 42), scale);
            int resultHeight = WorkbenchUi.GetAdaptiveResultCardHeight(viewportHeight, _resultCard.Top - scrollOffset.Y, stackButton ? 292 : 276, 430, scale);
            _resultCard.Size = new Size(availableWidth, resultHeight);
            _lblResultTitle.Location = new Point(S(20), S(14));
            _lblResultTitle.Size = new Size(availableWidth - S(40), S(28));
            _lblResultStatus.Location = new Point(availableWidth - S(186), S(16));
            _lblResultStatus.Size = new Size(S(166), S(28));
            _resultHost.Location = new Point(S(20), S(50));
            int buttonTop = _resultCard.Height - _btnApply.Height - S(20);
            _resultHost.Size = new Size(availableWidth - S(40), Math.Max(S(144), buttonTop - _resultHost.Top - S(16)));
            _txtResult.Location = new Point(S(12), S(12));
            _txtResult.Size = new Size(_resultHost.Width - S(24), _resultHost.Height - S(24));
            _btnApply.Location = stackButton
                ? new Point(S(20), buttonTop)
                : new Point(availableWidth - _btnApply.Width - S(20), buttonTop);

            AutoScrollMinSize = new Size(0, _resultCard.Bottom - scrollOffset.Y + S(24));
        }

        private async Task LoadVersionsAsync()
        {
            if (_cboChannel.SelectedItem is not ChannelOptionView channelView)
            {
                return;
            }

            int requestId = ++_versionLoadId;
            _isLoadingVersions = true;
            _loadedVersions.Clear();
            try
            {
                _lblLoading.Visible = true;
                _btnApply.Enabled = false;
                _cboVersion.Items.Clear();
                _cboVersion.Enabled = false;
                WorkbenchUi.ApplyStatusStyle(_lblResultStatus, WorkbenchTone.Info);
                _lblResultStatus.Text = T("正在加载版本", "Loading Versions");
                _txtResult.Text = T("正在加载版本列表，请稍候...", "Loading versions, please wait...");
                _versionSummaryText = _txtResult.Text;

                var versions = await _officeChannelService.GetAvailableVersionsAsync(channelView.Raw.Channel);
                if (IsDisposed || requestId != _versionLoadId) return;
                _loadedVersions = versions.ToList();
                if (_loadedVersions.Count == 0)
                {
                    throw new InvalidOperationException(T("未获取到该频道的版本列表。", "No version list was retrieved for this channel."));
                }

                RebuildVersionItems();

                _isLoadingVersions = false;
                RefreshResultTextForLanguage();
            }
            catch (Exception ex)
            {
                if (IsDisposed || requestId != _versionLoadId) return;
                WorkbenchUi.ApplyStatusStyle(_lblResultStatus, WorkbenchTone.Danger);
                _lblResultStatus.Text = T("加载失败", "Load Failed");
                _txtResult.Text = T($"加载版本列表失败：{ex.Message}", $"Failed to load versions: {ex.Message}");
                MessageBox.Show(
                    T($"加载版本列表失败：{ex.Message}", $"Failed to load versions: {ex.Message}"),
                    T("版本加载失败", "Version Load Failed"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            finally
            {
                if (!IsDisposed && requestId == _versionLoadId)
                {
                    _isLoadingVersions = false;
                    _lblLoading.Visible = false;
                    UpdateUiState();
                }
            }
        }

        private async void btnApply_Click(object? sender, EventArgs e)
        {
            if (_isRunning || _isLoadingVersions)
            {
                return;
            }

            if (_cboChannel.SelectedItem is not ChannelOptionView channelView)
            {
                MessageBox.Show(
                    T("请先选择目标频道。", "Please select a target channel first."),
                    T("提示", "Info"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            OfficeVersionOption? version = (_cboVersion.SelectedItem as VersionOptionView)?.Raw;
            string versionText = version == null || version.IsLatest
                ? T("最新版本", "Latest Version")
                : $"Build {version.Build}";

            DialogResult confirm = MessageBox.Show(
                T(
                    $"将切换到：{channelView.Display}\n目标版本：{versionText}\n禁用自动更新：{(_chkDisableUpdates.Checked ? "是" : "否")}\n\n请先保存文档；如更新提示关闭应用，请按提示处理。\n\n是否继续？",
                    $"Will switch to: {channelView.Display}\nTarget version: {versionText}\nDisable automatic updates: {(_chkDisableUpdates.Checked ? "Yes" : "No")}\n\nSave your documents first and follow any prompts to close apps.\n\nContinue?"),
                T("确认切换更新频道", "Confirm Channel Switch"),
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            if (confirm != DialogResult.Yes)
            {
                return;
            }

            try
            {
                _isRunning = true;
                UpdateUiState();
                _txtResult.Text = T(
                    $"正在切换到 {channelView.Display}，目标版本：{versionText}...",
                    $"Switching to {channelView.Display}, target version: {versionText}...");
                WorkbenchUi.ApplyStatusStyle(_lblResultStatus, WorkbenchTone.Info);
                _lblResultStatus.Text = T("正在切换", "Switching");

                if (!_officeChannelService.IsOfficeClickToRunAvailable())
                {
                    throw new InvalidOperationException(T("未检测到 Office ClickToRun 客户端。", "Office ClickToRun client was not detected."));
                }

                bool disableUpdates = _chkDisableUpdates.Checked;
                UiLanguage executionLanguage = _language;
                await Task.Run(() => _officeChannelService.SwitchChannel(channelView.Raw, disableUpdates, version, executionLanguage));
                if (IsDisposed || Disposing) return;
                WorkbenchUi.ApplyStatusStyle(_lblResultStatus, WorkbenchTone.Success);
                _lblResultStatus.Text = T("更新已发起", "Update Requested");
                _txtResult.Text = T(
                    $"频道配置已提交，Office 更新已发起。\r\n\r\n目标频道：{channelView.Display}\r\n目标版本：{versionText}\r\n频道变更需安装目标频道的更新后生效；组织策略可能覆盖本地设置。",
                    $"Channel configuration submitted and Office update requested.\r\n\r\nTarget channel: {channelView.Display}\r\nTarget version: {versionText}\r\nThe channel takes effect after the target update is installed. Organization policies may override local settings.");
                MessageBox.Show(
                    T("Office 更新已发起，安装完成后频道变更才会生效。", "Office update has been requested. The channel changes after the update is installed."),
                    T("完成", "Done"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                if (IsDisposed || Disposing) return;
                WorkbenchUi.ApplyStatusStyle(_lblResultStatus, WorkbenchTone.Danger);
                _lblResultStatus.Text = T("切换失败", "Switch Failed");
                _txtResult.Text = T($"切换失败：{ex.Message}", $"Switch failed: {ex.Message}");
                MessageBox.Show(
                    T($"切换失败：{ex.Message}", $"Switch failed: {ex.Message}"),
                    T("错误", "Error"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                _isRunning = false;
                if (!IsDisposed && !Disposing) UpdateUiState();
            }
        }

        private void UpdateUiState()
        {
            bool enabled = !_isRunning;
            _cboChannel.Enabled = enabled;
            _cboVersion.Enabled = enabled && !_isLoadingVersions && _cboVersion.Items.Count > 0;
            _chkDisableUpdates.Enabled = enabled;
            _btnApply.Enabled = enabled && !_isLoadingVersions && _cboVersion.Items.Count > 0;
        }

        private void CloseAllDropDowns()
        {
            _cboChannel.DroppedDown = false;
            _cboVersion.DroppedDown = false;
        }

        private string GetVersionDisplayName(OfficeVersionOption version)
        {
            if (version.IsLatest)
            {
                return T(
                    $"最新版本 ({version.Version} / Build {version.Build})",
                    $"Latest Version ({version.Version} / Build {version.Build})");
            }

            return $"Version {version.Version} (Build {version.Build}) - {version.ReleaseDate}";
        }

        private void RefreshResultTextForLanguage()
        {
            if (_isLoadingVersions)
            {
                WorkbenchUi.ApplyStatusStyle(_lblResultStatus, WorkbenchTone.Info);
                _lblResultStatus.Text = T("正在加载版本", "Loading Versions");
                _txtResult.Text = T("正在加载版本列表，请稍候...", "Loading versions, please wait...");
                _versionSummaryText = _txtResult.Text;
                return;
            }

            if (_isRunning)
            {
                return;
            }

            if (!string.IsNullOrWhiteSpace(_txtResult.Text) && _txtResult.Text != _versionSummaryText)
            {
                return;
            }

            if (_loadedVersions.Count == 0 || _cboChannel.SelectedItem is not ChannelOptionView channelView)
            {
                WorkbenchUi.ApplyStatusStyle(_lblResultStatus, WorkbenchTone.Info);
                _lblResultStatus.Text = T("等待选择频道", "Awaiting Selection");
                _txtResult.Text = T(
                    "可以在这里查看频道切换的状态、目标版本和执行结果。",
                    "View channel switch status, target version, and execution results here.");
                _versionSummaryText = _txtResult.Text;
                return;
            }

            WorkbenchUi.ApplyStatusStyle(_lblResultStatus, WorkbenchTone.Success);
            _lblResultStatus.Text = T("版本列表就绪", "Versions Ready");
            _txtResult.Text = T(
                $"已加载 {channelView.Display} 的版本列表，可选择最新版本或固定版本。\r\n\r\n{GetChannelGuidance(channelView.Raw.Channel)}",
                $"Loaded versions for {channelView.Display}. You can choose latest or fixed version.\r\n\r\n{GetChannelGuidance(channelView.Raw.Channel)}");
            _versionSummaryText = _txtResult.Text;
        }

        private string GetChannelGuidance(UpdateChannel channel) => channel == UpdateChannel.Broad
            ? T("自 2606 起，原半年企业频道接收月度企业更新；旧 CDN 的 Build 可能不同。本列表仅使用旧频道对应版本，新部署建议选择月度企业频道。",
                "From version 2606, legacy Semi-Annual devices receive monthly enterprise updates. Legacy CDN builds may differ; this list uses only that channel's builds. Use Monthly Enterprise for new deployments.")
            : T("频道变更需安装目标频道的更新后生效；组织策略优先于本地配置。",
                "The channel changes after the target update is installed. Organization policies take precedence over local configuration.");

        private string T(string zh, string en) => LocalizationService.Localize(_language, zh, en);
    }
}
