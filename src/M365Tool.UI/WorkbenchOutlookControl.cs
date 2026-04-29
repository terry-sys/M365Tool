using System;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Office365CleanupTool.Services;

namespace Office365CleanupTool
{
    public class WorkbenchOutlookControl : UserControl
    {
        private readonly OutlookDiagnosticsService _outlookDiagnosticsService;
        private readonly Panel _scrollHost;
        private readonly Panel _contentPanel;
        private readonly Label _lblTitle;
        private readonly Label _lblDescription;
        private readonly Panel _cardsHost;
        private readonly Panel _fullScanCard;
        private readonly Panel _offlineScanCard;
        private readonly Panel _calendarScanCard;
        private readonly Label _lblFullScanTitle;
        private readonly Label _lblFullScanDescription;
        private readonly Label _lblOfflineScanTitle;
        private readonly Label _lblOfflineScanDescription;
        private readonly Label _lblCalendarScanTitle;
        private readonly Label _lblCalendarScanDescription;
        private readonly Panel _logCard;
        private readonly Label _lblLogTitle;
        private readonly Label _lblLogStatus;
        private readonly Panel _resultCard;
        private readonly Panel _resultHost;
        private readonly Label _lblResultTitle;
        private readonly Label _lblResultStatus;
        private readonly TextBox _txtResult;
        private readonly Label _lblLastLogPath;
        private readonly Button _btnFullScan;
        private readonly Button _btnOfflineScan;
        private readonly Button _btnCalendarScan;
        private readonly Button _btnOpenLogFolder;
        private string _lastLogDirectory = string.Empty;
        private UiLanguage _language;

        public WorkbenchOutlookControl(OutlookDiagnosticsService outlookDiagnosticsService)
        {
            AutoScaleMode = AutoScaleMode.Dpi;
            _outlookDiagnosticsService = outlookDiagnosticsService;
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

            _fullScanCard = WorkbenchUi.CreateActionCard(string.Empty, string.Empty, Point.Empty, out _btnFullScan, true, string.Empty);
            (_lblFullScanTitle, _lblFullScanDescription) = GetActionCardLabels(_fullScanCard);

            _offlineScanCard = WorkbenchUi.CreateActionCard(string.Empty, string.Empty, Point.Empty, out _btnOfflineScan, false, string.Empty);
            (_lblOfflineScanTitle, _lblOfflineScanDescription) = GetActionCardLabels(_offlineScanCard);

            _calendarScanCard = WorkbenchUi.CreateActionCard(string.Empty, string.Empty, Point.Empty, out _btnCalendarScan, false, string.Empty);
            (_lblCalendarScanTitle, _lblCalendarScanDescription) = GetActionCardLabels(_calendarScanCard);

            _cardsHost.Controls.AddRange(new Control[] { _fullScanCard, _offlineScanCard, _calendarScanCard });

            _logCard = WorkbenchUi.CreateCard(Point.Empty, Size.Empty);
            _lblLogTitle = WorkbenchUi.CreateSectionTitle(string.Empty, new Point(20, 14));
            _lblLogStatus = new Label { AutoSize = false, TextAlign = ContentAlignment.MiddleLeft };
            WorkbenchUi.ApplyStatusStyle(_lblLogStatus, WorkbenchTone.Warning);
            _lblLogStatus.Visible = false;
            _lblLastLogPath = new Label
            {
                Font = new Font("Microsoft YaHei UI", 9.5F),
                ForeColor = WorkbenchUi.SecondaryTextColor,
                BackColor = Color.Transparent,
                Location = new Point(20, 42),
                Size = new Size(760, 36)
            };
            _btnOpenLogFolder = WorkbenchUi.CreateSecondaryButton(string.Empty, new Point(806, 28), new Size(146, 42));
            _logCard.Controls.AddRange(new Control[] { _lblLogTitle, _lblLogStatus, _lblLastLogPath, _btnOpenLogFolder });

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

            _contentPanel.Controls.AddRange(new Control[] { _lblTitle, _lblDescription, _cardsHost, _logCard, _resultCard });
            _scrollHost.Controls.Add(_contentPanel);
            Controls.Add(_scrollHost);

            _btnFullScan.Click += btnFullScan_Click;
            _btnOfflineScan.Click += btnOfflineScan_Click;
            _btnCalendarScan.Click += btnCalendarScan_Click;
            _btnOpenLogFolder.Click += btnOpenLogFolder_Click;

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

            _lblTitle.Text = T("Outlook 工具", "Outlook Tools");
            _lblTitle.Visible = false;
            _lblDescription.Text = string.Empty;
            _lblDescription.Visible = false;

            _lblFullScanTitle.Text = T("全量扫描", "Full Scan");
            _lblFullScanDescription.Text = T(
                "对 Outlook、Office 和 Windows 执行完整扫描，并生成综合诊断报告。",
                "Run comprehensive scan for Outlook, Office and Windows, and generate diagnostics report.");
            _btnFullScan.Text = T("开始扫描", "Start Scan");

            _lblOfflineScanTitle.Text = T("离线扫描", "Offline Scan");
            _lblOfflineScanDescription.Text = T(
                "\u517c\u5bb9\u6a21\u5f0f\u6267\u884c Outlook \u626b\u63cf\uff0c\u5982 Outlook \u672a\u8fd0\u884c\u4f1a\u81ea\u52a8\u8f6c\u4e3a\u79bb\u7ebf\u626b\u63cf\u3002",
                "Run Outlook scan in compatibility mode. If Outlook is not running, it should fall back to offline scan.");
            _btnOfflineScan.Text = T("开始扫描", "Start Scan");

            _lblCalendarScanTitle.Text = T("日历扫描", "Calendar Scan");
            _lblCalendarScanDescription.Text = T(
                "调用 CalCheck 检查 Outlook 日历中的已知问题并生成结果。",
                "Use CalCheck to scan known issues in Outlook calendar and generate results.");
            _btnCalendarScan.Text = T("开始扫描", "Start Scan");

            _lblLogTitle.Text = T("最近报告目录", "Recent Report Folder");
            _lblLastLogPath.Text = string.IsNullOrWhiteSpace(_lastLogDirectory)
                ? T("尚未生成报告", "No report generated yet")
                : _lastLogDirectory;
            WorkbenchUi.ApplyStatusStyle(_lblLogStatus, string.IsNullOrWhiteSpace(_lastLogDirectory) ? WorkbenchTone.Warning : WorkbenchTone.Success);
            _lblLogStatus.Text = string.IsNullOrWhiteSpace(_lastLogDirectory)
                ? T("暂无报告", "No Report")
                : T("报告可用", "Report Ready");
            _btnOpenLogFolder.Text = T("打开日志目录", "Open Log");

            _lblResultTitle.Text = T("处理结果", "Results");
            if (string.IsNullOrWhiteSpace(_txtResult.Text) ||
                _txtResult.Text.Contains("查看 Outlook", StringComparison.OrdinalIgnoreCase) ||
                _txtResult.Text.Contains("View Outlook", StringComparison.OrdinalIgnoreCase))
            {
                WorkbenchUi.ApplyStatusStyle(_lblResultStatus, WorkbenchTone.Info);
                _lblResultStatus.Text = T("等待执行", "Ready");
                _txtResult.Text = T(
                    "可以在这里查看 Outlook 扫描结果、返回码以及日志目录。",
                    "View Outlook scan results, return codes and log folder here.");
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

            _logCard.Location = new Point(contentLeft, _cardsHost.Bottom + 20);
            bool logStacked = availableWidth < 760;
            _logCard.Size = new Size(availableWidth, logStacked ? 134 : 108);
            _lblLogStatus.Location = new Point(availableWidth - 152, 16);
            _lblLogStatus.Size = new Size(132, 28);
            _lblLastLogPath.Size = new Size(logStacked ? availableWidth - 40 : Math.Max(320, _logCard.Width - 200), 36);
            _btnOpenLogFolder.Location = logStacked
                ? new Point(20, 82)
                : new Point(_logCard.Width - _btnOpenLogFolder.Width - 20, 34);

            _resultCard.Location = new Point(contentLeft, _logCard.Bottom + 18);
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

        private async void btnFullScan_Click(object? sender, EventArgs e)
        {
            await RunScenarioAsync(
                T(
                    "这会调用 Microsoft Get Help 官方场景，对 Outlook、Office、Windows 和邮箱配置执行完整扫描，并生成诊断报告。\n\n是否继续？",
                    "This runs official Microsoft Get Help scenario for full Outlook scan and diagnostics report.\n\nContinue?"),
                T("确认执行 Outlook 全量扫描", "Confirm Outlook Full Scan"),
                T("正在执行 Outlook 全量扫描，请稍候...", "Running Outlook full scan, please wait..."),
                () => _outlookDiagnosticsService.RunOutlookScanAsync(false, _language));
        }

        private async void btnOfflineScan_Click(object? sender, EventArgs e)
        {
            await RunScenarioAsync(
                T(
                    "\u8fd9\u4f1a\u8c03\u7528 Microsoft Get Help \u5b98\u65b9\u573a\u666f\uff0c\u4ee5\u517c\u5bb9\u6a21\u5f0f\u6267\u884c Outlook \u626b\u63cf\u5e76\u751f\u6210\u8bca\u65ad\u62a5\u544a\u3002\n\n\u662f\u5426\u7ee7\u7eed\uff1f",
                    "This runs official Microsoft Get Help scenario in compatibility mode.\n\nContinue?"),
                T("确认执行 Outlook 离线扫描", "Confirm Outlook Offline Scan"),
                T("正在执行 Outlook 离线扫描，请稍候...", "Running Outlook offline scan, please wait..."),
                () => _outlookDiagnosticsService.RunOutlookScanAsync(true, _language));
        }

        private async void btnCalendarScan_Click(object? sender, EventArgs e)
        {
            await RunScenarioAsync(
                T(
                    "这会调用 Microsoft Get Help 官方场景，使用 CalCheck 扫描 Outlook 日历中的已知问题，并生成诊断结果。\n\n是否继续？",
                    "This runs official Microsoft Get Help scenario to scan Outlook calendar issues via CalCheck.\n\nContinue?"),
                T("确认执行 Outlook 日历扫描", "Confirm Outlook Calendar Scan"),
                T("正在执行 Outlook 日历扫描，请稍候...", "Running Outlook calendar scan, please wait..."),
                () => _outlookDiagnosticsService.RunCalendarScanAsync(_language));
        }

        private void btnOpenLogFolder_Click(object? sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(_lastLogDirectory))
                {
                    throw new InvalidOperationException(T("当前还没有可打开的 Outlook 诊断日志目录。", "No Outlook diagnostics log folder is available yet."));
                }

                _outlookDiagnosticsService.OpenFolder(_lastLogDirectory, _language);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    T($"打开日志目录失败：\n{ex.Message}", $"Failed to open log folder:\n{ex.Message}"),
                    T("错误", "Error"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private async Task RunScenarioAsync(
            string confirmMessage,
            string confirmTitle,
            string runningMessage,
            Func<Task<OutlookDiagnosticResult>> action)
        {
            DialogResult confirm = MessageBox.Show(
                confirmMessage,
                confirmTitle,
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (confirm != DialogResult.Yes)
            {
                return;
            }

            await RunBusyAsync(async () =>
            {
                WorkbenchUi.ApplyStatusStyle(_lblResultStatus, WorkbenchTone.Info);
                _lblResultStatus.Text = T("正在扫描", "Scanning");
                _txtResult.Text = runningMessage;
                OutlookDiagnosticResult result = await action();
                _lastLogDirectory = result.LogDirectory;
                _lblLastLogPath.Text = string.IsNullOrWhiteSpace(_lastLogDirectory)
                    ? T("尚未生成报告", "No report generated yet")
                    : _lastLogDirectory;
                WorkbenchUi.ApplyStatusStyle(_lblLogStatus, string.IsNullOrWhiteSpace(_lastLogDirectory) ? WorkbenchTone.Warning : WorkbenchTone.Success);
                _lblLogStatus.Text = string.IsNullOrWhiteSpace(_lastLogDirectory)
                    ? T("暂无报告", "No Report")
                    : T("报告可用", "Report Ready");
                WorkbenchUi.ApplyStatusStyle(_lblResultStatus, result.Success ? WorkbenchTone.Success : WorkbenchTone.Warning);
                _lblResultStatus.Text = result.Success ? T("扫描完成", "Completed") : T("需要关注", "Needs Review");
                _txtResult.Text = result.Message;

                MessageBox.Show(
                    result.Success
                        ? T("Outlook 诊断已完成，可查看报告目录。", "Outlook diagnostics completed. You can open the report folder.")
                        : T("诊断操作已执行，请查看结果详情。", "Diagnostics action completed. Please review details."),
                    result.Success ? T("完成", "Done") : T("提示", "Info"),
                    MessageBoxButtons.OK,
                    result.Success ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
            });
        }

        private async Task RunBusyAsync(Func<Task> action)
        {
            try
            {
                UseWaitCursor = true;
                _btnFullScan.Enabled = false;
                _btnOfflineScan.Enabled = false;
                _btnCalendarScan.Enabled = false;
                _btnOpenLogFolder.Enabled = false;
                await action();
            }
            catch (Exception ex)
            {
                WorkbenchUi.ApplyStatusStyle(_lblResultStatus, WorkbenchTone.Danger);
                _lblResultStatus.Text = T("扫描失败", "Failed");
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
                _btnFullScan.Enabled = true;
                _btnOfflineScan.Enabled = true;
                _btnCalendarScan.Enabled = true;
                _btnOpenLogFolder.Enabled = true;
            }
        }

        private string T(string zh, string en) => LocalizationService.Localize(_language, zh, en);
    }
}
