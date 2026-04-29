using System;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Office365CleanupTool.Services;

namespace Office365CleanupTool
{
    public class WorkbenchTeamsControl : UserControl
    {
        private readonly TeamsMaintenanceService _teamsMaintenanceService;
        private readonly Panel _scrollHost;
        private readonly Panel _contentPanel;
        private readonly Label _lblTitle;
        private readonly Label _lblDescription;
        private readonly Panel _cardsHost;
        private readonly Panel _maintenanceCard;
        private readonly Panel _signInCard;
        private readonly Panel _repairCard;
        private readonly Label _lblMaintenanceCardTitle;
        private readonly Label _lblMaintenanceCardDescription;
        private readonly Label _lblSignInCardTitle;
        private readonly Label _lblSignInCardDescription;
        private readonly Label _lblRepairCardTitle;
        private readonly Label _lblRepairCardDescription;
        private readonly Panel _helperCard;
        private readonly Label _lblHelperTitle;
        private readonly Label _lblHelperDescription;
        private readonly Panel _resultCard;
        private readonly Panel _resultHost;
        private readonly Label _lblResultTitle;
        private readonly Label _lblResultStatus;
        private readonly TextBox _txtResult;
        private readonly Button _btnRunMaintenance;
        private readonly Button _btnDeleteSignInRecords;
        private readonly Button _btnRepairOutlookAddin;
        private readonly Button _btnOpenCacheFolder;
        private UiLanguage _language;

        public WorkbenchTeamsControl(TeamsMaintenanceService teamsMaintenanceService)
        {
            AutoScaleMode = AutoScaleMode.Dpi;
            _teamsMaintenanceService = teamsMaintenanceService;
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

            _cardsHost = new Panel { BackColor = Color.Transparent };

            _maintenanceCard = WorkbenchUi.CreateActionCard(string.Empty, string.Empty, Point.Empty, out _btnRunMaintenance, true, string.Empty);
            _maintenanceCard.Name = "maintenanceCard";
            (_lblMaintenanceCardTitle, _lblMaintenanceCardDescription) = GetActionCardLabels(_maintenanceCard);

            _signInCard = WorkbenchUi.CreateActionCard(string.Empty, string.Empty, Point.Empty, out _btnDeleteSignInRecords, false, string.Empty);
            _signInCard.Name = "signInCard";
            (_lblSignInCardTitle, _lblSignInCardDescription) = GetActionCardLabels(_signInCard);

            _repairCard = WorkbenchUi.CreateActionCard(string.Empty, string.Empty, Point.Empty, out _btnRepairOutlookAddin, false, string.Empty);
            _repairCard.Name = "repairCard";
            (_lblRepairCardTitle, _lblRepairCardDescription) = GetActionCardLabels(_repairCard);

            _cardsHost.Controls.AddRange(new Control[] { _maintenanceCard, _signInCard, _repairCard });

            _helperCard = WorkbenchUi.CreateCard(Point.Empty, Size.Empty);
            _lblHelperTitle = WorkbenchUi.CreateSectionTitle(string.Empty, new Point(20, 14));
            _lblHelperDescription = new Label
            {
                Font = new Font("Microsoft YaHei UI", 9.5F),
                ForeColor = WorkbenchUi.SecondaryTextColor,
                BackColor = Color.Transparent,
                Location = new Point(20, 42),
                Size = new Size(560, 22)
            };
            _btnOpenCacheFolder = WorkbenchUi.CreateSecondaryButton(string.Empty, new Point(806, 22), new Size(146, 42));
            _helperCard.Controls.AddRange(new Control[] { _lblHelperTitle, _lblHelperDescription, _btnOpenCacheFolder });

            _resultCard = WorkbenchUi.CreateCard(Point.Empty, Size.Empty);
            _lblResultTitle = WorkbenchUi.CreateSectionTitle(string.Empty, new Point(20, 14));
            _lblResultStatus = new Label { AutoSize = false, TextAlign = ContentAlignment.MiddleLeft };
            WorkbenchUi.ApplyStatusStyle(_lblResultStatus, WorkbenchTone.Info);
            _lblResultStatus.Visible = false;
            _resultHost = WorkbenchUi.CreateInsetPanel(Point.Empty, Size.Empty);
            _txtResult = WorkbenchUi.CreateReadonlyTextBox(
                Point.Empty,
                Size.Empty,
                string.Empty);
            _txtResult.BorderStyle = BorderStyle.None;
            _txtResult.BackColor = Color.White;
            _resultHost.Controls.Add(_txtResult);
            _resultCard.Controls.AddRange(new Control[] { _lblResultTitle, _lblResultStatus, _resultHost });

            _contentPanel.Controls.AddRange(new Control[] { _lblTitle, _lblDescription, _cardsHost, _helperCard, _resultCard });
            _scrollHost.Controls.Add(_contentPanel);
            Controls.Add(_scrollHost);

            _btnRunMaintenance.Click += btnRunMaintenance_Click;
            _btnDeleteSignInRecords.Click += btnDeleteSignInRecords_Click;
            _btnRepairOutlookAddin.Click += btnRepairOutlookAddin_Click;
            _btnOpenCacheFolder.Click += btnOpenCacheFolder_Click;

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

            _lblTitle.Text = T("Teams 工具", "Teams Tools");
            _lblTitle.Visible = false;
            _lblDescription.Text = string.Empty;
            _lblDescription.Visible = false;

            _lblMaintenanceCardTitle.Text = T("重置并清理缓存", "Reset and Clear Cache");
            _lblMaintenanceCardDescription.Text = T(
                "重置新版 Teams，并同步清理本地缓存，不影响云端聊天记录和文件。",
                "Reset new Teams and clear local cache without affecting cloud chats/files.");
            _btnRunMaintenance.Text = T("开始处理", "Start");

            _lblSignInCardTitle.Text = T("删除登录记录", "Delete Sign-in Records");
            _lblSignInCardDescription.Text = T(
                "清理本地登录记录，处理完成后必须先重启电脑，再重新打开 New Teams。",
                "Clear local sign-in records. Restart PC before opening New Teams again.");
            _btnDeleteSignInRecords.Text = T("开始处理", "Start");

            _lblRepairCardTitle.Text = T("会议插件修复", "Meeting Add-in Repair");
            _lblRepairCardDescription.Text = T(
                "调用 Microsoft Get Help 官方场景，修复 Outlook 中的 Teams Meeting Add-in。",
                "Run official Microsoft Get Help scenario to repair Teams Meeting Add-in in Outlook.");
            _btnRepairOutlookAddin.Text = T("开始处理", "Start");

            _lblHelperTitle.Text = T("辅助操作", "Helper Actions");
            _lblHelperDescription.Text = T(
                "用于打开新版 Teams 缓存目录，便于人工检查处理结果。",
                "Open new Teams cache folder for manual inspection.");
            _btnOpenCacheFolder.Text = T("打开缓存目录", "Open Cache");

            _lblResultTitle.Text = T("处理结果", "Results");
            if (string.IsNullOrWhiteSpace(_txtResult.Text) ||
                _txtResult.Text.Contains("查看 Teams", StringComparison.OrdinalIgnoreCase) ||
                _txtResult.Text.Contains("View Teams", StringComparison.OrdinalIgnoreCase))
            {
                WorkbenchUi.ApplyStatusStyle(_lblResultStatus, WorkbenchTone.Info);
                _lblResultStatus.Text = T("等待执行", "Ready");
                _txtResult.Text = T(
                    "可以在这里查看 Teams 处理结果、已关闭进程、已处理目录和下一步建议。",
                    "View Teams results, closed processes, processed folders and next-step suggestions here.");
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

            int columns = availableWidth >= 980 ? 3 : availableWidth >= 680 ? 2 : 1;
            int cardWidth = (availableWidth - gap * (columns - 1)) / columns;
            int cardHeight = 160;
            _cardsHost.Location = new Point(contentLeft, 18);
            int rows = (int)Math.Ceiling(_cardsHost.Controls.Count / (double)columns);
            _cardsHost.Size = new Size(availableWidth, rows * cardHeight + (rows - 1) * gap);

            for (int i = 0; i < _cardsHost.Controls.Count; i++)
            {
                Control card = _cardsHost.Controls[i];
                int row = i / columns;
                int column = i % columns;
                card.Location = new Point(column * (cardWidth + gap), row * (cardHeight + gap));
                card.Size = new Size(cardWidth, cardHeight);
            }

            _helperCard.Location = new Point(contentLeft, _cardsHost.Bottom + 16);
            bool helperStacked = availableWidth < 760;
            _helperCard.Size = new Size(availableWidth, helperStacked ? 114 : 78);
            _btnOpenCacheFolder.Location = helperStacked
                ? new Point(20, 62)
                : new Point(_helperCard.Width - _btnOpenCacheFolder.Width - 20, 18);

            _resultCard.Location = new Point(contentLeft, _helperCard.Bottom + 16);
            _resultCard.Size = new Size(availableWidth, 252);
            _lblResultStatus.Location = new Point(availableWidth - 168, 16);
            _lblResultStatus.Size = new Size(148, 28);
            _resultHost.Location = new Point(20, 50);
            _resultHost.Size = new Size(_resultCard.Width - 40, 182);
            _txtResult.Location = new Point(12, 12);
            _txtResult.Size = new Size(_resultHost.Width - 24, _resultHost.Height - 24);

            _contentPanel.Width = Math.Max(10, viewportWidth - 1);
            _contentPanel.Height = _resultCard.Bottom + 24;
        }

        private async void btnRunMaintenance_Click(object? sender, EventArgs e)
        {
            DialogResult confirm = MessageBox.Show(
                T(
                    "这会按顺序执行以下操作：\n\n1. 关闭正在运行的 Teams 进程\n2. 重置新版 Teams 应用\n3. 清理新版 Teams 本地缓存\n\n此操作不会影响用户云端数据，聊天记录、团队和文件等内容仍保存在 Microsoft 365 云端。\n\n是否继续？",
                    "This will:\n\n1. Close running Teams processes\n2. Reset new Teams app\n3. Clear local Teams cache\n\nCloud data remains intact. Continue?"),
                T("确认处理 Teams", "Confirm Teams Processing"),
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (confirm != DialogResult.Yes)
            {
                return;
            }

            await RunBusyAsync(async () =>
            {
                WorkbenchUi.ApplyStatusStyle(_lblResultStatus, WorkbenchTone.Info);
                _lblResultStatus.Text = T("正在处理", "Processing");
                _txtResult.Text = T("正在处理 Teams，请稍候...", "Processing Teams, please wait...");
                TeamsMaintenanceResult result = await _teamsMaintenanceService.ResetAndClearCacheAsync(_language);
                WorkbenchUi.ApplyStatusStyle(_lblResultStatus, result.Success ? WorkbenchTone.Success : WorkbenchTone.Warning);
                _lblResultStatus.Text = result.Success ? T("处理完成", "Completed") : T("需要关注", "Needs Review");
                _txtResult.Text = result.Message;

                MessageBox.Show(
                    result.Success ? T("Teams 已完成重置并清理缓存。", "Teams reset and cache cleanup completed.") : T("Teams 处理已执行，请查看结果详情。", "Teams processing completed. Please review details."),
                    result.Success ? T("完成", "Done") : T("提示", "Info"),
                    MessageBoxButtons.OK,
                    result.Success ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
            });
        }

        private async void btnDeleteSignInRecords_Click(object? sender, EventArgs e)
        {
            DialogResult confirm = MessageBox.Show(
                T(
                    "这会删除新版 Teams 的本地登录记录，并清理以下目录中的文件和文件夹：\n\n- %LocalAppData%\\Microsoft\\IdentityCache\n- %LocalAppData%\\Microsoft\\TokenBroker\n- %LocalAppData%\\Packages\\MSTeams_8wekyb3d8bbwe\n- %LocalAppData%\\Packages\\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\n- %LocalAppData%\\Microsoft\\OneAuth\n\n此操作不会影响用户云端数据，但完成后必须重启电脑。若不重启就直接打开 New Teams，可能出现 Error 1001。\n\n是否继续？",
                    "This removes local sign-in records and clears required folders.\nCloud data remains safe, but you must restart the computer before opening New Teams again.\n\nContinue?"),
                T("确认删除 Teams 登录记录", "Confirm Delete Teams Sign-in Records"),
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (confirm != DialogResult.Yes)
            {
                return;
            }

            await RunBusyAsync(async () =>
            {
                WorkbenchUi.ApplyStatusStyle(_lblResultStatus, WorkbenchTone.Info);
                _lblResultStatus.Text = T("正在处理", "Processing");
                _txtResult.Text = T("正在删除 Teams 登录记录，请稍候...", "Deleting Teams sign-in records, please wait...");
                TeamsSignInRecordCleanupResult result = await _teamsMaintenanceService.DeleteSignInRecordsAsync(_language);
                WorkbenchUi.ApplyStatusStyle(_lblResultStatus, result.Success ? WorkbenchTone.Success : WorkbenchTone.Warning);
                _lblResultStatus.Text = result.Success ? T("处理完成", "Completed") : T("需要关注", "Needs Review");
                _txtResult.Text = result.Message;

                MessageBox.Show(
                    T("Teams 登录记录处理已完成。\n\n请务必先重启电脑，再重新打开 New Teams 登录。", "Teams sign-in cleanup completed.\n\nPlease restart your PC before opening New Teams again."),
                    T("完成", "Done"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            });
        }

        private async void btnRepairOutlookAddin_Click(object? sender, EventArgs e)
        {
            DialogResult confirm = MessageBox.Show(
                T(
                    "这会调用 Microsoft Get Help 官方场景，自动检查并尝试恢复 Outlook 的 Teams Meeting Add-in。\n\n执行过程中可能会关闭 Outlook。\n此功能用于 Outlook 中的 Teams 会议插件，不会影响 Teams 云端数据。\n\n是否继续？",
                    "This runs official Microsoft Get Help scenario to repair Teams Meeting Add-in in Outlook.\nOutlook may be closed during execution.\n\nContinue?"),
                T("确认修复 Outlook 会议插件", "Confirm Repair Outlook Meeting Add-in"),
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (confirm != DialogResult.Yes)
            {
                return;
            }

            await RunBusyAsync(async () =>
            {
                WorkbenchUi.ApplyStatusStyle(_lblResultStatus, WorkbenchTone.Info);
                _lblResultStatus.Text = T("正在处理", "Processing");
                _txtResult.Text = T("正在修复 Outlook 的 Teams 会议插件，请稍候...", "Repairing Outlook Teams meeting add-in, please wait...");
                TeamsAddinRepairResult result = await _teamsMaintenanceService.RepairOutlookAddinAsync(_language);
                WorkbenchUi.ApplyStatusStyle(_lblResultStatus, result.Success ? WorkbenchTone.Success : WorkbenchTone.Warning);
                _lblResultStatus.Text = result.Success ? T("处理完成", "Completed") : T("需要关注", "Needs Review");
                _txtResult.Text = result.Message;

                MessageBox.Show(
                    result.Success
                        ? T("Outlook 的 Teams 会议插件修复已完成，请重新启动 Outlook 检查。", "Outlook Teams meeting add-in repair completed. Restart Outlook to verify.")
                        : T("修复操作已执行，请查看结果详情。", "Repair action executed. Please review details."),
                    result.Success ? T("完成", "Done") : T("提示", "Info"),
                    MessageBoxButtons.OK,
                    result.Success ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
            });
        }

        private void btnOpenCacheFolder_Click(object? sender, EventArgs e)
        {
            TeamsPathTarget target = _teamsMaintenanceService.GetPrimaryCacheTarget(_language);

            try
            {
                _teamsMaintenanceService.OpenFolder(target.Path, _language);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    T($"打开缓存目录失败：\n{ex.Message}\n\n目录位置：\n{target.Path}", $"Failed to open cache folder:\n{ex.Message}\n\nFolder:\n{target.Path}"),
                    T("错误", "Error"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private async Task RunBusyAsync(Func<Task> action)
        {
            try
            {
                UseWaitCursor = true;
                _btnRunMaintenance.Enabled = false;
                _btnDeleteSignInRecords.Enabled = false;
                _btnRepairOutlookAddin.Enabled = false;
                _btnOpenCacheFolder.Enabled = false;
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
                UseWaitCursor = false;
                _btnRunMaintenance.Enabled = true;
                _btnDeleteSignInRecords.Enabled = true;
                _btnRepairOutlookAddin.Enabled = true;
                _btnOpenCacheFolder.Enabled = true;
            }
        }

        private string T(string zh, string en) => LocalizationService.Localize(_language, zh, en);
    }
}
