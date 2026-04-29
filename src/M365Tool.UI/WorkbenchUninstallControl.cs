using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
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
        private readonly Label _lblDescription;
        private readonly TextBox _txtDetectedOffice;
        private readonly TextBox _txtResult;
        private readonly Button _btnOpenLogs;
        private readonly Button _btnStartUninstall;

        private bool _isRunning;
        private int _detectedPreferredHeight = 170;
        private string _lastLogDirectory = string.Empty;
        private UiLanguage _language;

        public WorkbenchUninstallControl(OfficeUninstallService officeUninstallService)
        {
            AutoScaleMode = AutoScaleMode.Dpi;
            _officeUninstallService = officeUninstallService;
            _officeUninstallService.ProgressChanged += OnProgressChanged;
            _language = LocalizationService.ResolveLanguage("Auto");

            Dock = DockStyle.Fill;
            BackColor = WorkbenchUi.CanvasColor;
            AutoScroll = true;

            _lblTitle = new Label { Font = new Font("Microsoft YaHei UI", 17F, FontStyle.Bold), ForeColor = WorkbenchUi.PrimaryTextColor };
            _lblDescriptionTitle = new Label { Font = new Font("Microsoft YaHei UI", 10.25F), ForeColor = WorkbenchUi.SecondaryTextColor };

            _targetCard = WorkbenchUi.CreateCard(Point.Empty, Size.Empty);
            _lblTargetTitle = WorkbenchUi.CreateSectionTitle(string.Empty, new Point(20, 14));
            _cboTarget = WorkbenchUi.CreateComboBox(Point.Empty, new Size(420, 32));
            _lblDescription = new Label
            {
                Font = new Font("Microsoft YaHei UI", 9.5F),
                ForeColor = WorkbenchUi.SecondaryTextColor,
                BackColor = Color.Transparent,
                UseMnemonic = false
            };
            _targetCard.Controls.AddRange(new Control[] { _lblTargetTitle, _cboTarget, _lblDescription });

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

            _cboTarget.SelectedIndexChanged += (_, _) => UpdateTargetDescription();

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
            RefreshDetectedOffice();
            UpdateUiState();
        }

        public void ApplyLanguage(UiLanguage language)
        {
            _language = language;
            _lblTitle.Text = T("卸载", "Uninstall");
            _lblTitle.Visible = false;
            _lblDescriptionTitle.Text = string.Empty;
            _lblDescriptionTitle.Visible = false;
            _lblTargetTitle.Text = T("卸载目标", "Uninstall Target");
            _lblDetectedTitle.Text = T("当前检测到的 Office", "Detected Office");
            _lblResultTitle.Text = T("处理结果", "Results");
            _btnOpenLogs.Text = T("打开日志目录", "Open Log");
            _btnStartUninstall.Text = T("开始卸载", "Start Uninstall");

            if (string.IsNullOrWhiteSpace(_txtResult.Text) ||
                _txtResult.Text.Contains("查看卸载结果", StringComparison.OrdinalIgnoreCase) ||
                _txtResult.Text.Contains("View uninstall results", StringComparison.OrdinalIgnoreCase))
            {
                _txtResult.Text = T(
                    "可以在这里查看卸载结果、返回码、进程处理摘要和下一步建议。",
                    "View uninstall results, return code, process summary and next-step suggestions here.");
            }

            RebuildTargets();
            UpdateTargetDescription();
            RefreshDetectedOffice();
        }

        public void PrepareForDisplay()
        {
            if (IsHandleCreated)
            {
                AutoScrollPosition = new Point(0, 0);
            }

            LayoutControls();
        }

        private void RebuildTargets()
        {
            string? selectedKey = (_cboTarget.SelectedItem as TargetOptionView)?.Raw.SaraOfficeVersion ?? "__AUTO__";
            _cboTarget.BeginUpdate();
            _cboTarget.Items.Clear();

            foreach (OfficeUninstallTargetOption option in _officeUninstallService.GetSupportedTargets(_language))
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
        }

        private void LayoutControls()
        {
            const int margin = 24;
            int viewportWidth = ClientSize.Width;
            int availableWidth = Math.Max(560, Math.Min(1160, viewportWidth - margin * 2));
            int contentLeft = Math.Max(margin, (viewportWidth - availableWidth) / 2);

            _targetCard.Location = new Point(contentLeft, 18);
            _targetCard.Size = new Size(availableWidth, 162);
            _cboTarget.Location = new Point(20, 46);
            _cboTarget.Width = Math.Min(420, availableWidth - 40);
            _lblDescription.Location = new Point(20, 88);
            _lblDescription.Size = new Size(availableWidth - 40, 52);

            _detectedCard.Location = new Point(contentLeft, _targetCard.Bottom + 18);
            _detectedCard.Size = new Size(availableWidth, _detectedPreferredHeight);
            _detectedHost.Location = new Point(20, 52);
            _detectedHost.Size = new Size(availableWidth - 40, _detectedCard.Height - 72);
            _txtDetectedOffice.Location = new Point(12, 12);
            _txtDetectedOffice.Size = new Size(_detectedHost.Width - 24, _detectedHost.Height - 24);

            _resultCard.Location = new Point(contentLeft, _detectedCard.Bottom + 18);
            _resultCard.Size = new Size(availableWidth, 304);
            _resultHost.Location = new Point(20, 46);
            _resultHost.Size = new Size(availableWidth - 40, 166);
            _txtResult.Location = new Point(12, 12);
            _txtResult.Size = new Size(_resultHost.Width - 24, _resultHost.Height - 24);
            int buttonTop = _resultHost.Bottom + 20;
            _btnStartUninstall.Location = new Point(availableWidth - _btnStartUninstall.Width - 20, buttonTop);
            _btnOpenLogs.Location = new Point(_btnStartUninstall.Left - _btnOpenLogs.Width - 10, buttonTop);

            if (availableWidth < 640)
            {
                _btnStartUninstall.Location = new Point(20, buttonTop);
                _btnOpenLogs.Location = new Point(20 + _btnStartUninstall.Width + 10, buttonTop);
            }

            AutoScrollMinSize = new Size(0, _resultCard.Bottom + 24);
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

            int lineCount = Math.Max(1, _txtDetectedOffice.Lines.Length);
            _detectedPreferredHeight = Math.Min(360, Math.Max(190, 94 + lineCount * 22));
            LayoutControls();
        }

        private void UpdateTargetDescription()
        {
            _lblDescription.Text = (_cboTarget.SelectedItem as TargetOptionView)?.Description ?? string.Empty;
        }

        private async void btnStartUninstall_Click(object? sender, EventArgs e)
        {
            if (_isRunning)
            {
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
                    $"将使用微软官方卸载场景执行：{targetView.Display}\n\n此操作可能关闭 Office 相关进程，并建议在完成后重启电脑。\n\n是否继续？",
                    $"The official Microsoft uninstall scenario will run for: {targetView.Display}\n\nThis may close Office processes and a restart is recommended after completion.\n\nContinue?"),
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
                UpdateUiState();
                WorkbenchUi.ApplyStatusStyle(_lblDetectedStatus, WorkbenchTone.Info);
                _lblDetectedStatus.Text = T("正在分析目标", "Preparing");
                _txtResult.Text = T("正在准备新版 Office 卸载引擎...", "Preparing latest Office uninstall engine...");

                OfficeUninstallResult result = await _officeUninstallService.RunAsync(new OfficeUninstallRequest
                {
                    Target = targetView.Raw
                }, _language);

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
                _isRunning = false;
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
            _cboTarget.Enabled = enabled;
            _btnStartUninstall.Enabled = enabled;
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
                "All" => T("所有 Office 版本", "All Office Versions"),
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
