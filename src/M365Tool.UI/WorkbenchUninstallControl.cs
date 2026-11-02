using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Threading;
using Office365CleanupTool.Models;
using Office365CleanupTool.Services;

namespace Office365CleanupTool
{
    public class WorkbenchUninstallControl : UserControl
    {
        private sealed class TargetOptionView
        {
            public OfficeUninstallTargetOption Raw { get; set; } = new();
            public string Display { get; set; } = string.Empty;
            public string Description { get; set; } = string.Empty;
            public override string ToString() => Display;
        }

        private readonly OfficeUninstallService _officeUninstallService;
        private readonly Label _lblTitle;
        private readonly Label _lblDescriptionTitle;
        private readonly Panel _targetCard;
        private readonly Panel _detectedCard;
        private readonly Panel _resultCard;
        private readonly Panel _detectedHost;
        private readonly Panel _resultHost;
        private readonly Label _lblTargetTitle;
        private readonly Label _lblDetectedTitle;
        private readonly Label _lblResultTitle;
        private readonly Label _lblDetectedStatus;
        private readonly ComboBox _cboTarget;
        private readonly ComboBox _cboEngine;
        private readonly CheckBox _chkForceClose;
        private CancellationTokenSource? _cancellation;
        private readonly Label _lblDescription;
        private readonly TextBox _txtDetectedOffice;
        private readonly TextBox _txtResult;
        private readonly Button _btnOpenLogs;
        private readonly Button _btnStartUninstall;

        private bool _isRunning;
        private bool _rebindingControls;
        private int _detectedPreferredHeight = 170;
        private string _lastLogDirectory = string.Empty;
        private UiLanguage _language;

        public WorkbenchUninstallControl(OfficeUninstallService officeUninstallService)
        {
            AutoScaleMode = AutoScaleMode.None;
            _officeUninstallService = officeUninstallService;
            _officeUninstallService.ProgressChanged += OnProgressChanged;
            _language = LocalizationService.ResolveLanguage("Auto");

            Dock = DockStyle.Fill;
            BackColor = WorkbenchUi.CanvasColor;
            AutoScroll = true;

            _lblTitle = new Label { Font = WorkbenchUi.CreateUiFont(17F, FontStyle.Bold), ForeColor = WorkbenchUi.PrimaryTextColor };
            _lblDescriptionTitle = new Label { Font = WorkbenchUi.CreateUiFont(10.25F), ForeColor = WorkbenchUi.SecondaryTextColor };

            _targetCard = WorkbenchUi.CreateCard(Point.Empty, Size.Empty);
            _lblTargetTitle = WorkbenchUi.CreateSectionTitle(string.Empty, new Point(20, 14));
            _cboTarget = WorkbenchUi.CreateComboBox(Point.Empty, new Size(420, 32));
            _cboEngine = WorkbenchUi.CreateComboBox(Point.Empty, new Size(420, 32));
            _chkForceClose = new CheckBox { Font = WorkbenchUi.CreateUiFont(9F), AutoSize = false, BackColor = Color.Transparent, ForeColor = WorkbenchUi.PrimaryTextColor };
            _cboEngine.SelectedIndexChanged += (_, _) =>
            {
                if (!_rebindingControls) RefreshDetectedOffice();
            };
            _lblDescription = new Label
            {
                Font = WorkbenchUi.CreateUiFont(9.5F),
                ForeColor = WorkbenchUi.SecondaryTextColor,
                BackColor = Color.Transparent,
                UseMnemonic = false
            };
            _targetCard.Controls.AddRange(new Control[] { _lblTargetTitle, _cboEngine, _cboTarget, _chkForceClose, _lblDescription });

            _detectedCard = WorkbenchUi.CreateCard(Point.Empty, Size.Empty);
            _lblDetectedTitle = WorkbenchUi.CreateSectionTitle(string.Empty, new Point(20, 14));
            _lblDetectedStatus = new Label { AutoSize = false, TextAlign = ContentAlignment.MiddleLeft };
            WorkbenchUi.ApplyStatusStyle(_lblDetectedStatus, WorkbenchTone.Info);
            _lblDetectedStatus.Visible = false;
            _detectedHost = WorkbenchUi.CreateInsetPanel(Point.Empty, Size.Empty);
            _txtDetectedOffice = WorkbenchUi.CreateReadonlyTextBox(Point.Empty, Size.Empty);
            _txtDetectedOffice.BorderStyle = BorderStyle.None;
            _txtDetectedOffice.BackColor = Color.White;
            _txtDetectedOffice.WordWrap = false;
            _txtDetectedOffice.ScrollBars = ScrollBars.Both;
            _detectedHost.Controls.Add(_txtDetectedOffice);
            _detectedCard.Controls.AddRange(new Control[] { _lblDetectedTitle, _lblDetectedStatus, _detectedHost });

            _resultCard = WorkbenchUi.CreateCard(Point.Empty, Size.Empty);
            _lblResultTitle = WorkbenchUi.CreateSectionTitle(string.Empty, new Point(20, 14));
            _resultHost = WorkbenchUi.CreateInsetPanel(Point.Empty, Size.Empty);
            _txtResult = WorkbenchUi.CreateReadonlyTextBox(Point.Empty, Size.Empty, string.Empty);
            _txtResult.BorderStyle = BorderStyle.None;
            _txtResult.BackColor = Color.White;
            _btnOpenLogs = WorkbenchUi.CreateSecondaryButton(string.Empty, Point.Empty, new Size(132, 42));
            _btnStartUninstall = WorkbenchUi.CreatePrimaryButton(string.Empty, Point.Empty, new Size(132, 42));
            _btnOpenLogs.Click += btnOpenLogs_Click;
            _btnStartUninstall.Click += btnStartUninstall_Click;
            _resultHost.Controls.Add(_txtResult);
            _resultCard.Controls.AddRange(new Control[] { _lblResultTitle, _resultHost, _btnOpenLogs, _btnStartUninstall });

            Controls.AddRange(new Control[] { _lblTitle, _lblDescriptionTitle, _targetCard, _detectedCard, _resultCard });

            _cboTarget.SelectedIndexChanged += (_, _) =>
            {
                UpdateTargetDescription();
                UpdateUiState();
            };

            Resize += (_, _) => LayoutControls();
            Scroll += (_, _) => CloseAllDropDowns();
            MouseWheel += (_, _) => CloseAllDropDowns();
            ParentChanged += (_, _) => this.BeginInvokeWhenReady(PrepareForDisplay);
            VisibleChanged += (_, _) =>
            {
                if (Visible)
                {
                    this.BeginInvokeWhenReady(PrepareForDisplay);
                }
            };

            ApplyLanguage(_language);
            LayoutControls();
            UpdateUiState();
        }

        public void ApplyLanguage(UiLanguage language)
        {
            _language = language;
            _lblTitle.Text = T("卸载", "Uninstall");
            _lblTitle.Visible = false;
            _lblDescriptionTitle.Text = string.Empty;
            _lblDescriptionTitle.Visible = false;
            _lblTargetTitle.Text = T("卸载方式与目标", "Uninstall Method and Target");
            int engineIndex = Math.Max(0, _cboEngine.SelectedIndex);
            _rebindingControls = true;
            try
            {
                _cboEngine.Items.Clear();
                _cboEngine.Items.Add(T("常规卸载（ODT，仅 C2R）", "Standard uninstall (ODT, C2R only)"));
                _cboEngine.Items.Add(T("深度清理（Get Help，含 MSI）", "Deep cleanup (Get Help, includes MSI)"));
                _cboEngine.SelectedIndex = engineIndex;
            }
            finally { _rebindingControls = false; }
            _chkForceClose.Text = T("允许强制关闭 Office（未保存内容可能丢失）", "Allow force closing Office (unsaved work may be lost)");
            _lblDetectedTitle.Text = T("当前检测到的 Office", "Detected Office");
            _lblResultTitle.Text = T("处理结果", "Results");
            _btnOpenLogs.Text = T("打开日志目录", "Open Log");
            _btnStartUninstall.Text = T("开始卸载", "Start Uninstall");

            if (string.IsNullOrWhiteSpace(_txtResult.Text) ||
                _txtResult.Text == "可以在这里查看卸载结果、返回码、进程处理摘要和下一步建议。" ||
                _txtResult.Text == "View uninstall results, return code, process summary and next-step suggestions here.")
            {
                _txtResult.Text = T(
                    "可以在这里查看卸载结果、返回码、进程处理摘要和下一步建议。",
                    "View uninstall results, return code, process summary and next-step suggestions here.");
            }

            RefreshDetectedOffice();
            UpdateUiState();
        }

        public void PrepareForDisplay()
        {
            if (IsHandleCreated)
            {
                AutoScrollPosition = new Point(0, 0);
            }

            if (!_isRunning) RefreshDetectedOffice();
            else LayoutControls();
        }

        private void RebuildTargets(IReadOnlyList<InstalledOfficeInfo> installedOffice)
        {
            string? selectedKey = (_cboTarget.SelectedItem as TargetOptionView)?.Raw.SaraOfficeVersion ?? "__AUTO__";
            _cboTarget.BeginUpdate();
            _cboTarget.Items.Clear();

            foreach (OfficeUninstallTargetOption option in _officeUninstallService.GetAvailableTargets(
                _cboEngine.SelectedIndex == 0 ? OfficeUninstallMode.Standard : OfficeUninstallMode.DeepCleanup, installedOffice, _language))
            {
                _cboTarget.Items.Add(new TargetOptionView
                {
                    Raw = option,
                    Display = GetTargetDisplayName(option),
                    Description = GetTargetDescription(option)
                });
            }

            if (_cboTarget.Items.Count > 0)
            {
                int selectedIndex = 0;
                for (int i = 0; i < _cboTarget.Items.Count; i++)
                {
                    if (_cboTarget.Items[i] is TargetOptionView view)
                    {
                        string key = view.Raw.SaraOfficeVersion ?? "__AUTO__";
                        if (string.Equals(key, selectedKey, StringComparison.OrdinalIgnoreCase))
                        {
                            selectedIndex = i;
                            break;
                        }
                    }
                }
                _cboTarget.SelectedIndex = selectedIndex;
            }

            _cboTarget.EndUpdate();
            UpdateTargetDescription();
            UpdateUiState();
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

            _targetCard.Location = new Point(contentLeft, S(18) + scrollOffset.Y);
            _targetCard.Size = new Size(availableWidth, S(224));
            _lblTargetTitle.Location = new Point(S(20), S(14));
            _lblTargetTitle.Size = new Size(availableWidth - S(40), S(28));
            _cboEngine.Location = new Point(S(20), S(46));
            _cboEngine.Size = new Size(availableWidth - S(40), S(32));
            _cboTarget.Location = new Point(S(20), S(88));
            _cboTarget.Size = new Size(Math.Min(S(420), availableWidth - S(40)), S(32));
            _chkForceClose.Location = new Point(S(20), S(126));
            _chkForceClose.Size = new Size(availableWidth - S(40), S(28));
            _lblDescription.Location = new Point(S(20), S(158));
            _lblDescription.Size = new Size(availableWidth - S(40), S(54));

            _detectedCard.Location = new Point(contentLeft, _targetCard.Bottom + S(18));
            _detectedCard.Size = new Size(availableWidth, S(_detectedPreferredHeight));
            _lblDetectedTitle.Location = new Point(S(20), S(14));
            _lblDetectedTitle.Size = new Size(availableWidth - S(40), S(28));
            _lblDetectedStatus.Location = new Point(availableWidth - S(188), S(16));
            _lblDetectedStatus.Size = new Size(S(168), S(28));
            _detectedHost.Location = new Point(S(20), S(52));
            _detectedHost.Size = new Size(availableWidth - S(40), _detectedCard.Height - S(72));
            _txtDetectedOffice.Location = new Point(S(12), S(12));
            _txtDetectedOffice.Size = new Size(_detectedHost.Width - S(24), _detectedHost.Height - S(24));

            _resultCard.Location = new Point(contentLeft, _detectedCard.Bottom + S(18));
            int resultHeight = WorkbenchUi.GetAdaptiveResultCardHeight(viewportHeight, _resultCard.Top - scrollOffset.Y, 304, 450, scale);
            _resultCard.Size = new Size(availableWidth, resultHeight);
            _lblResultTitle.Location = new Point(S(20), S(14));
            _lblResultTitle.Size = new Size(availableWidth - S(40), S(28));
            _btnOpenLogs.Size = WorkbenchUi.ScaleSize(new Size(132, 42), scale);
            _btnStartUninstall.Size = WorkbenchUi.ScaleSize(new Size(132, 42), scale);
            _resultHost.Location = new Point(S(20), S(46));
            int buttonTop = _resultCard.Height - _btnStartUninstall.Height - S(20);
            _resultHost.Size = new Size(availableWidth - S(40), Math.Max(S(166), buttonTop - _resultHost.Top - S(20)));
            _txtResult.Location = new Point(S(12), S(12));
            _txtResult.Size = new Size(_resultHost.Width - S(24), _resultHost.Height - S(24));
            _btnStartUninstall.Location = new Point(availableWidth - _btnStartUninstall.Width - S(20), buttonTop);
            _btnOpenLogs.Location = new Point(_btnStartUninstall.Left - _btnOpenLogs.Width - S(10), buttonTop);

            if (availableWidth < 640)
            {
                _btnStartUninstall.Location = new Point(S(20), buttonTop);
                _btnOpenLogs.Location = new Point(S(20) + _btnStartUninstall.Width + S(10), buttonTop);
            }

            AutoScrollMinSize = new Size(0, _resultCard.Bottom - scrollOffset.Y + S(24));
        }

        private void RefreshDetectedOffice()
        {
            var installedOffice = _officeUninstallService.DetectInstalledOffice();
            _txtDetectedOffice.Text = installedOffice.Count == 0
                ? T("未检测到已安装的 Office。", "No installed Office detected.")
                : string.Join(Environment.NewLine, installedOffice.Select(item => $"• {item}"));

            if (installedOffice.Count == 0)
            {
                WorkbenchUi.ApplyStatusStyle(_lblDetectedStatus, WorkbenchTone.Warning);
                _lblDetectedStatus.Text = T("未检测到 Office", "No Office Found");
            }
            else
            {
                WorkbenchUi.ApplyStatusStyle(_lblDetectedStatus, WorkbenchTone.Info);
                _lblDetectedStatus.Text = T($"已检测到 {installedOffice.Count} 项", $"{installedOffice.Count} Items Detected");
            }

            RebuildTargets(installedOffice);
            int lineCount = Math.Max(1, _txtDetectedOffice.Lines.Length);
            _detectedPreferredHeight = Math.Min(360, Math.Max(190, 94 + lineCount * 22));
            LayoutControls();
        }

        private void UpdateTargetDescription()
        {
            _lblDescription.Text = _cboEngine.SelectedIndex == 0
                ? _cboTarget.SelectedItem is TargetOptionView
                    ? T("仅移除检测到的 C2R 产品。MSI 或损坏的安装可选择深度清理。取消后，已启动的卸载可能继续运行。", "Removes detected C2R products. Use deep cleanup for MSI or damaged installs. Canceling monitoring may leave uninstall running.")
                    : T("未检测到可常规卸载的 C2R 产品。MSI 或损坏的安装请选择深度清理。", "No C2R products are available for standard uninstall. Use deep cleanup for MSI or damaged installs.")
                : (_cboTarget.SelectedItem as TargetOptionView)?.Description ?? string.Empty;
        }

        private async void btnStartUninstall_Click(object? sender, EventArgs e)
        {
            if (_isRunning)
            {
                _cancellation?.Cancel();
                _btnStartUninstall.Enabled = false;
                return;
            }

            if (_cboTarget.SelectedItem is not TargetOptionView targetView)
            {
                MessageBox.Show(
                    T("请先选择卸载目标。", "Please select an uninstall target first."),
                    T("提示", "Info"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            DialogResult confirm = MessageBox.Show(
                T(
                    $"将执行：{_cboEngine.Text} — {targetView.Display}\n\n请先保存文档。工具准备完成后会尝试关闭 Office；强制关闭：{_chkForceClose.Checked}。取消等待不会强制中止已启动的卸载。\n\n是否继续？",
                    $"Run: {_cboEngine.Text} — {targetView.Display}\n\nSave your work first. Office will be closed after tools are ready; force close: {_chkForceClose.Checked}. Canceling monitoring will not forcibly stop an uninstall already started.\n\nContinue?"),
                T("确认卸载", "Confirm Uninstall"),
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);
            if (confirm != DialogResult.Yes)
            {
                return;
            }

            try
            {
                _isRunning = true;
                _cancellation = new CancellationTokenSource();
                UpdateUiState();
                WorkbenchUi.ApplyStatusStyle(_lblDetectedStatus, WorkbenchTone.Info);
                _lblDetectedStatus.Text = T("正在分析目标", "Preparing");
                _txtResult.Text = T("正在准备新版 Office 卸载引擎...", "Preparing latest Office uninstall engine...");

                OfficeUninstallResult result = await _officeUninstallService.RunAsync(new OfficeUninstallRequest
                {
                    Target = targetView.Raw,
                    Mode = _cboEngine.SelectedIndex == 0 ? OfficeUninstallMode.Standard : OfficeUninstallMode.DeepCleanup,
                    ForceCloseApplications = _chkForceClose.Checked
                }, _language, _cancellation.Token);

                _lastLogDirectory = result.LogDirectory;
                WorkbenchUi.ApplyStatusStyle(_lblDetectedStatus, result.IsSuccess ? WorkbenchTone.Success : WorkbenchTone.Warning);
                _lblDetectedStatus.Text = result.IsSuccess ? T("卸载完成", "Completed") : T("需要关注", "Needs Review");
                _txtResult.Text = result.Details + Environment.NewLine + Environment.NewLine + T("下一步建议：", "Next step suggestion:") + result.NextStepAdvice;
                MessageBox.Show(
                    result.IsSuccess
                        ? T("Office 卸载已完成，可查看日志目录。", "Office uninstall completed. You can open the log folder.")
                        : T("卸载未完全完成，请查看结果详情。", "Uninstall did not fully complete. Please review the result details."),
                    result.IsSuccess ? T("完成", "Done") : T("提示", "Info"),
                    MessageBoxButtons.OK,
                    result.IsSuccess ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
                RefreshDetectedOffice();
            }
            catch (OperationCanceledException)
            {
                WorkbenchUi.ApplyStatusStyle(_lblDetectedStatus, WorkbenchTone.Warning);
                _lblDetectedStatus.Text = T("已取消准备或等待", "Canceled");
                _txtResult.Text = T("已取消准备或等待；已启动的卸载可能仍在后台运行，请确认结束后再重试。", "Preparation or monitoring was canceled. An uninstall already started may still be running; check before retrying.");
            }
            catch (Exception ex)
            {
                WorkbenchUi.ApplyStatusStyle(_lblDetectedStatus, WorkbenchTone.Danger);
                _lblDetectedStatus.Text = T("卸载失败", "Failed");
                _txtResult.Text = T($"卸载失败：{ex.Message}", $"Uninstall failed: {ex.Message}");
                MessageBox.Show(
                    T($"卸载失败：{ex.Message}", $"Uninstall failed: {ex.Message}"),
                    T("错误", "Error"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                _lastLogDirectory = _officeUninstallService.LastLogDirectory;
                _isRunning = false;
                _cancellation?.Dispose();
                _cancellation = null;
                UpdateUiState();
            }
        }

        private void btnOpenLogs_Click(object? sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(_lastLogDirectory))
                {
                    throw new InvalidOperationException(T("当前还没有可打开的卸载日志目录。", "No uninstall log folder is available yet."));
                }

                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = $"\"{_lastLogDirectory}\"",
                    UseShellExecute = true
                });
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

        private void OnProgressChanged(object? sender, OfficeUninstallProgressEventArgs e)
        {
            if (IsDisposed || Disposing || !IsHandleCreated)
            {
                return;
            }

            if (InvokeRequired)
            {
                if (!this.TryBeginInvokeIfReady(() => OnProgressChanged(sender, e)))
                {
                    return;
                }
                return;
            }

            _txtResult.Text = e.Message;
        }

        private void UpdateUiState()
        {
            bool enabled = !_isRunning;
            bool hasTarget = _cboTarget.SelectedItem is TargetOptionView;
            _cboTarget.Enabled = enabled && hasTarget;
            _cboEngine.Enabled = enabled;
            _chkForceClose.Enabled = enabled && hasTarget;
            _btnStartUninstall.Enabled = _isRunning ? _cancellation?.IsCancellationRequested == false : hasTarget;
            _btnStartUninstall.Text = _isRunning ? T("取消/停止等待", "Cancel / Stop Waiting") : T("开始卸载", "Start Uninstall");
            _btnOpenLogs.Enabled = !string.IsNullOrWhiteSpace(_lastLogDirectory);
        }

        private void CloseAllDropDowns()
        {
            _cboTarget.DroppedDown = false;
        }

        private string GetTargetDisplayName(OfficeUninstallTargetOption option)
        {
            string key = option.SaraOfficeVersion ?? "__AUTO__";
            return key switch
            {
                "__AUTO__" => T("自动检测已安装的 Office", "Auto-detect installed Office"),
                "All" => _cboEngine.SelectedIndex == 0
                    ? T("全部已检测到的 C2R 产品", "All detected C2R products")
                    : T("所有 Office 版本", "All Office Versions"),
                _ => option.SaraOfficeVersion?.StartsWith("20", StringComparison.OrdinalIgnoreCase) == true
                    ? $"Office {option.SaraOfficeVersion}"
                    : option.SaraOfficeVersion == "M365" ? "Microsoft 365 Apps" : option.DisplayName
            };
        }

        private string GetTargetDescription(OfficeUninstallTargetOption option)
        {
            string key = option.SaraOfficeVersion ?? "__AUTO__";
            return key switch
            {
                "__AUTO__" => T(
                    "让微软官方 SaRA 自动检测当前设备上的 Office 版本并尝试卸载。若设备上同时存在多个版本，SaRA 会返回多版本结果码。",
                    "Let Microsoft SaRA auto-detect installed Office versions and attempt uninstall. If multiple versions exist, SaRA returns multi-version result codes."),
                "M365" => T(
                    "卸载订阅版 Office，例如 Microsoft 365 Apps for enterprise / business。",
                    "Uninstall subscription Office, such as Microsoft 365 Apps for enterprise / business."),
                "All" => T(
                    "调用 SaRA 的 All 模式，尝试移除设备上的所有 Office 版本。",
                    "Use SaRA All mode to remove all installed Office versions."),
                _ => T($"仅卸载 Office {key}。", $"Uninstall Office {key} only.")
            };
        }

        private string T(string zh, string en) => LocalizationService.Localize(_language, zh, en);
    }
}
