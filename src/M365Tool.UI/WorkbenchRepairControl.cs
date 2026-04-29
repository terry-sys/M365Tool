using System;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Office365CleanupTool.Services;

namespace Office365CleanupTool
{
    public class WorkbenchRepairControl : UserControl
    {
        private readonly CleanupService _cleanupService;
        private readonly NetworkRepairService _networkRepairService;
        private readonly Panel _scrollHost;
        private readonly Panel _contentPanel;
        private readonly Label _lblTitle;
        private readonly Label _lblDescription;
        private readonly Panel _cardsHost;
        private readonly Panel _cleanupCard;
        private readonly Panel _repairCard;
        private readonly Label _lblCleanupCardTitle;
        private readonly Label _lblCleanupCardDescription;
        private readonly Label _lblRepairCardTitle;
        private readonly Label _lblRepairCardDescription;
        private readonly Panel _resultCard;
        private readonly Panel _resultHost;
        private readonly Label _lblResultTitle;
        private readonly Label _lblResultStatus;
        private readonly TextBox _txtResult;
        private readonly Button _btnCleanup;
        private readonly Button _btnRepairNetwork;
        private bool _isRunning;
        private UiLanguage _language;

        public WorkbenchRepairControl(CleanupService cleanupService, NetworkRepairService networkRepairService)
        {
            AutoScaleMode = AutoScaleMode.Dpi;
            _cleanupService = cleanupService;
            _networkRepairService = networkRepairService;
            _cleanupService.ProgressChanged += OnCleanupProgressChanged;
            _networkRepairService.ProgressChanged += OnNetworkRepairProgressChanged;
            _language = LocalizationService.ResolveLanguage("Auto");

            Dock = DockStyle.Fill;
            BackColor = WorkbenchUi.CanvasColor;
            AutoScroll = false;
            DoubleBuffered = true;

            _scrollHost = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = WorkbenchUi.PageSurfaceColor,
                AutoScroll = true
            };

            _contentPanel = new Panel
            {
                Location = new Point(0, 0),
                BackColor = WorkbenchUi.PageSurfaceColor,
                Dock = DockStyle.Top
            };

            _lblTitle = new Label { Font = new Font("Microsoft YaHei UI", 17F, FontStyle.Bold), ForeColor = WorkbenchUi.PrimaryTextColor, BackColor = Color.Transparent };
            _lblDescription = new Label { Font = new Font("Microsoft YaHei UI", 10.25F), ForeColor = WorkbenchUi.SecondaryTextColor, BackColor = Color.Transparent };

            _cardsHost = new Panel { BackColor = Color.Transparent };
            _cleanupCard = WorkbenchUi.CreateActionCard(string.Empty, string.Empty, Point.Empty, out _btnCleanup, true, string.Empty);
            _repairCard = WorkbenchUi.CreateActionCard(string.Empty, string.Empty, Point.Empty, out _btnRepairNetwork, false, string.Empty);
            (_lblCleanupCardTitle, _lblCleanupCardDescription) = GetActionCardLabels(_cleanupCard);
            (_lblRepairCardTitle, _lblRepairCardDescription) = GetActionCardLabels(_repairCard);
            _cardsHost.Controls.AddRange(new Control[] { _cleanupCard, _repairCard });

            _resultCard = WorkbenchUi.CreateCard(Point.Empty, Size.Empty);
            _lblResultTitle = WorkbenchUi.CreateSectionTitle(string.Empty, new Point(20, 14));
            _lblResultStatus = new Label { AutoSize = false, TextAlign = ContentAlignment.MiddleLeft };
            WorkbenchUi.ApplyStatusStyle(_lblResultStatus, WorkbenchTone.Info);
            _lblResultStatus.Visible = false;
            _resultHost = WorkbenchUi.CreateInsetPanel(Point.Empty, Size.Empty);
            _txtResult = WorkbenchUi.CreateReadonlyTextBox(Point.Empty, Size.Empty, string.Empty);
            _txtResult.BorderStyle = BorderStyle.None;
            _txtResult.BackColor = Color.White;
            _resultHost.Controls.Add(_txtResult);
            _resultCard.Controls.AddRange(new Control[] { _lblResultTitle, _lblResultStatus, _resultHost });

            _contentPanel.Controls.AddRange(new Control[] { _lblTitle, _lblDescription, _cardsHost, _resultCard });
            _scrollHost.Controls.Add(_contentPanel);
            Controls.Add(_scrollHost);
            _btnCleanup.Click += btnCleanup_Click;
            _btnRepairNetwork.Click += btnRepairNetwork_Click;

            Resize += (_, _) => LayoutControls();
            _scrollHost.Resize += (_, _) => LayoutControls();
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
        }

        public void ApplyLanguage(UiLanguage language)
        {
            _language = language;

            _lblTitle.Text = T("清理与修复", "Cleanup & Repair");
            _lblTitle.Visible = false;
            _lblDescription.Text = string.Empty;
            _lblDescription.Visible = false;

            _lblCleanupCardTitle.Text = T("清理激活信息", "Clean Activation Info");
            _lblCleanupCardDescription.Text = T(
                "下载并执行微软官方清理工具，处理 Office 激活信息、WAM 账号和 WPJ 关联残留。",
                "Run official Microsoft cleanup tool to remove Office activation remnants, WAM account and WPJ traces.");
            _btnCleanup.Text = T("开始处理", "Start");

            _lblRepairCardTitle.Text = T("修复网络与激活环境", "Repair Network & Activation");
            _lblRepairCardDescription.Text = T(
                "修改 IE / WinHTTP 相关设置与代理配置，帮助恢复 Office 激活所需网络环境。",
                "Repair IE/WinHTTP settings and proxy to restore network prerequisites for Office activation.");
            _btnRepairNetwork.Text = T("开始处理", "Start");

            _lblResultTitle.Text = T("处理结果", "Results");
            if (string.IsNullOrWhiteSpace(_txtResult.Text) ||
                _txtResult.Text.Contains("查看清理", StringComparison.OrdinalIgnoreCase) ||
                _txtResult.Text.Contains("View cleanup", StringComparison.OrdinalIgnoreCase))
            {
                WorkbenchUi.ApplyStatusStyle(_lblResultStatus, WorkbenchTone.Info);
                _lblResultStatus.Text = T("等待执行", "Ready");
                _txtResult.Text = T(
                    "可以在这里查看清理或修复的执行进度与结果。",
                    "View cleanup/repair progress and results here.");
            }
        }

        public void PrepareForDisplay()
        {
            if (_scrollHost.IsHandleCreated)
            {
                _scrollHost.AutoScrollPosition = new Point(0, 0);
            }

            LayoutControls();
        }

        private static (Label Title, Label Description) GetActionCardLabels(Panel card)
        {
            Label[] labels = card.Controls.OfType<Label>().OrderBy(item => item.Top).ToArray();
            return labels.Length >= 2
                ? (labels[0], labels[1])
                : (new Label(), new Label());
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
            int availableWidth = Math.Max(560, Math.Min(1160, viewportWidth - margin * 2 - 4));
            int contentLeft = Math.Max(margin, (viewportWidth - availableWidth) / 2);
            int columns = availableWidth >= 900 ? 2 : 1;
            int cardWidth = columns == 1 ? availableWidth : (availableWidth - gap) / 2;
            int cardHeight = 160;

            _cardsHost.Location = new Point(contentLeft, 18);
            _cardsHost.Size = new Size(availableWidth, columns == 1 ? cardHeight * 2 + gap : cardHeight);
            _cardsHost.Controls[0].Location = new Point(0, 0);
            _cardsHost.Controls[0].Size = new Size(cardWidth, cardHeight);
            _cardsHost.Controls[1].Location = columns == 1 ? new Point(0, cardHeight + gap) : new Point(cardWidth + gap, 0);
            _cardsHost.Controls[1].Size = new Size(cardWidth, cardHeight);

            _resultCard.Location = new Point(contentLeft, _cardsHost.Bottom + 16);
            _resultCard.Size = new Size(availableWidth, 248);
            _lblResultStatus.Location = new Point(availableWidth - 168, 16);
            _lblResultStatus.Size = new Size(148, 28);
            _resultHost.Location = new Point(20, 50);
            _resultHost.Size = new Size(availableWidth - 40, 176);
            _txtResult.Location = new Point(12, 12);
            _txtResult.Size = new Size(_resultHost.Width - 24, _resultHost.Height - 24);

            _contentPanel.Width = Math.Max(10, viewportWidth - 1);
            _contentPanel.Height = _resultCard.Bottom + 24;
        }

        private async void btnCleanup_Click(object? sender, EventArgs e)
        {
            if (_isRunning) return;
            DialogResult confirm = MessageBox.Show(
                T(
                    "确定要清理激活信息吗？此操作将执行微软官方清理工具，建议关闭 Office 后继续。",
                    "Are you sure to clean activation info? This runs official Microsoft cleanup tool. Close Office first."),
                T("确认清理激活信息", "Confirm Clean Activation Info"),
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            if (confirm != DialogResult.Yes) return;

            await RunBusyAsync(async () =>
            {
                WorkbenchUi.ApplyStatusStyle(_lblResultStatus, WorkbenchTone.Info);
                _lblResultStatus.Text = T("正在处理", "Processing");
                _txtResult.Text = T("开始清理过程...", "Starting cleanup...");
                await _cleanupService.RunAsync(@"C:\TempPath\o365reset\", _language);
                WorkbenchUi.ApplyStatusStyle(_lblResultStatus, WorkbenchTone.Success);
                _lblResultStatus.Text = T("处理完成", "Completed");
                _txtResult.Text += Environment.NewLine + T("清理完成。", "Cleanup completed.");
                MessageBox.Show(
                    T("Office 365 清理过程已完成。", "Office 365 cleanup completed."),
                    T("完成", "Done"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            });
        }

        private async void btnRepairNetwork_Click(object? sender, EventArgs e)
        {
            if (_isRunning) return;
            DialogResult confirm = MessageBox.Show(
                T(
                    "确定要修复网络与激活环境吗？这会修改代理和 Internet Settings 相关配置。",
                    "Are you sure to repair network/activation environment? This modifies proxy and Internet Settings."),
                T("确认修复网络与激活环境", "Confirm Repair Network/Activation"),
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            if (confirm != DialogResult.Yes) return;

            await RunBusyAsync(async () =>
            {
                WorkbenchUi.ApplyStatusStyle(_lblResultStatus, WorkbenchTone.Info);
                _lblResultStatus.Text = T("正在处理", "Processing");
                _txtResult.Text = T("正在修复网络与激活环境...", "Repairing network and activation environment...");
                await _networkRepairService.RunAsync(_language);
                WorkbenchUi.ApplyStatusStyle(_lblResultStatus, WorkbenchTone.Success);
                _lblResultStatus.Text = T("处理完成", "Completed");
                _txtResult.Text += Environment.NewLine + T("网络问题修复完成。", "Network repair completed.");
                MessageBox.Show(
                    T("网络问题修复完成。", "Network repair completed."),
                    T("完成", "Done"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            });
        }

        private async Task RunBusyAsync(Func<Task> action)
        {
            try
            {
                _isRunning = true;
                _btnCleanup.Enabled = false;
                _btnRepairNetwork.Enabled = false;
                UseWaitCursor = true;
                await action();
            }
            catch (Exception ex)
            {
                WorkbenchUi.ApplyStatusStyle(_lblResultStatus, WorkbenchTone.Danger);
                _lblResultStatus.Text = T("处理失败", "Failed");
                _txtResult.Text = T($"执行失败：{ex.Message}", $"Execution failed: {ex.Message}");
                MessageBox.Show(
                    T($"执行失败：{ex.Message}", $"Execution failed: {ex.Message}"),
                    T("错误", "Error"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                _isRunning = false;
                _btnCleanup.Enabled = true;
                _btnRepairNetwork.Enabled = true;
                UseWaitCursor = false;
            }
        }

        private void OnCleanupProgressChanged(object? sender, CleanupProgressEventArgs e)
        {
            if (IsDisposed || Disposing || !IsHandleCreated)
            {
                return;
            }

            if (InvokeRequired)
            {
                if (!this.TryBeginInvokeIfReady(() => OnCleanupProgressChanged(sender, e)))
                {
                    return;
                }
                return;
            }
            _txtResult.Text = e.Message;
        }

        private void OnNetworkRepairProgressChanged(object? sender, NetworkRepairProgressEventArgs e)
        {
            if (IsDisposed || Disposing || !IsHandleCreated)
            {
                return;
            }

            if (InvokeRequired)
            {
                if (!this.TryBeginInvokeIfReady(() => OnNetworkRepairProgressChanged(sender, e)))
                {
                    return;
                }
                return;
            }
            _txtResult.Text = e.Message;
        }

        private string T(string zh, string en) => LocalizationService.Localize(_language, zh, en);
    }
}
