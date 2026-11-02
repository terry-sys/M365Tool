using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using Office365CleanupTool.Models;
using Office365CleanupTool.Services;

namespace Office365CleanupTool
{
    public class WorkbenchHomeControl : UserControl
    {
        private sealed class BufferedPanel : Panel
        {
            public BufferedPanel()
            {
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                    ControlStyles.ResizeRedraw | ControlStyles.UserPaint | ControlStyles.Selectable, true);
                DoubleBuffered = true;
            }
            protected override bool IsInputKey(Keys keyData) =>
                keyData is Keys.Enter or Keys.Space || base.IsInputKey(keyData);
            protected override void OnGotFocus(EventArgs e) { base.OnGotFocus(e); Invalidate(); }
            protected override void OnLostFocus(EventArgs e) { base.OnLostFocus(e); Invalidate(); }
            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                if (Focused && ShowFocusCues)
                    ControlPaint.DrawFocusRectangle(e.Graphics, Rectangle.Inflate(ClientRectangle, -5, -5));
            }
        }

        private sealed class IconPictureBox : PictureBox
        {
            protected override void OnPaint(PaintEventArgs e)
            {
                if (Image == null) return;
                e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                e.Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                e.Graphics.DrawImage(Image, ClientRectangle);
            }
        }

        private enum ModuleIconKind
        {
            Install,
            Uninstall,
            Channel,
            Repair,
            Teams,
            Outlook
        }

        private static readonly IReadOnlyDictionary<ModuleIconKind, string> ModuleIconResourceNames = new Dictionary<ModuleIconKind, string>
        {
            [ModuleIconKind.Install] = "M365Tool.Assets.Icons.FluentOutline.install.png",
            [ModuleIconKind.Uninstall] = "M365Tool.Assets.Icons.FluentOutline.uninstall.png",
            [ModuleIconKind.Channel] = "M365Tool.Assets.Icons.FluentOutline.channel.png",
            [ModuleIconKind.Repair] = "M365Tool.Assets.Icons.FluentOutline.repair.png",
            [ModuleIconKind.Teams] = "M365Tool.Assets.Icons.FluentOutline.teams.png",
            [ModuleIconKind.Outlook] = "M365Tool.Assets.Icons.FluentOutline.outlook.png"
        };

        private static readonly Lazy<IReadOnlyDictionary<ModuleIconKind, Image>> ModuleIconImages = new(LoadModuleIconImages);

        private sealed class ModuleDescriptionControls
        {
            public string PageKey { get; set; } = string.Empty;
            public Panel Container { get; init; } = null!;
            public PictureBox Icon { get; init; } = null!;
            public Label Title { get; init; } = null!;
            public Label Summary { get; init; } = null!;
            public Label Detail { get; init; } = null!;
            public Label Arrow { get; init; } = null!;
        }

        private UiLanguage _language;
        private readonly Panel _scrollHost;
        private readonly Panel _contentPanel;
        private readonly Panel _workspaceCard;
        private readonly Panel _environmentCard;
        private readonly Label[] _environmentLabels = new Label[3];
        private readonly Label[] _environmentValues = new Label[3];
        private readonly Label _lblTitle;
        private readonly Label _lblHint;
        private readonly Label _applicationTitle;
        private readonly Label _applicationHint;
        private readonly Label _diagnosticsTitle;
        private readonly Label _diagnosticsHint;
        private readonly ModuleDescriptionControls[] _modules;
        private IReadOnlyList<InstalledOfficeProduct> _products = Array.Empty<InstalledOfficeProduct>();
        private bool _inventoryUnavailable;
        private bool _layingOut;
        public event EventHandler<string>? ModuleSelected;

        public WorkbenchHomeControl(AppSettingsService settingsService, AppSettings settings, UiLanguage language)
        {
            AutoScaleMode = AutoScaleMode.None;
            _language = language;
            DoubleBuffered = true;
            BackColor = WorkbenchUi.CanvasColor;
            Dock = DockStyle.Fill;
            _scrollHost = new BufferedPanel { Dock = DockStyle.Fill, BackColor = BackColor, AutoScroll = true };
            _contentPanel = new BufferedPanel { Dock = DockStyle.Top, BackColor = BackColor };
            _workspaceCard = new BufferedPanel { BackColor = BackColor };
            _environmentCard = WorkbenchUi.CreateCard(Point.Empty, Size.Empty);
            _environmentCard.Paint += (_, e) =>
            {
                using var pen = new Pen(WorkbenchUi.CardBorderColor);
                for (int i = 1; i < 3; i++)
                    e.Graphics.DrawLine(pen, _environmentCard.Width * i / 3, 23, _environmentCard.Width * i / 3, 67);
            };
            for (int i = 0; i < 3; i++)
            {
                _environmentLabels[i] = CreateLabel(8.8F, FontStyle.Regular, WorkbenchUi.SecondaryTextColor);
                _environmentValues[i] = CreateLabel(10.5F, FontStyle.Bold, WorkbenchUi.PrimaryTextColor);
                _environmentValues[i].AutoEllipsis = true;
                _environmentCard.Controls.AddRange(new Control[] { _environmentLabels[i], _environmentValues[i] });
            }
            _lblTitle = CreateLabel(18.75F, FontStyle.Bold, WorkbenchUi.PrimaryTextColor);
            _lblTitle.TextAlign = ContentAlignment.MiddleLeft;
            _lblHint = CreateLabel(10.1F, FontStyle.Regular, WorkbenchUi.SecondaryTextColor);
            _applicationTitle = CreateLabel(11.25F, FontStyle.Bold, WorkbenchUi.PrimaryTextColor);
            _applicationHint = CreateLabel(8.8F, FontStyle.Regular, WorkbenchUi.SecondaryTextColor);
            _diagnosticsTitle = CreateLabel(11.25F, FontStyle.Bold, WorkbenchUi.PrimaryTextColor);
            _diagnosticsHint = CreateLabel(8.8F, FontStyle.Regular, WorkbenchUi.SecondaryTextColor);
            _applicationHint.TextAlign = _diagnosticsHint.TextAlign = ContentAlignment.MiddleRight;
            _workspaceCard.Controls.AddRange(new Control[] { _lblTitle, _lblHint, _applicationTitle,
                _applicationHint, _diagnosticsTitle, _diagnosticsHint });
            _workspaceCard.Controls.Add(_environmentCard);
            _modules = Enumerable.Range(0, 6).Select(_ => CreateModuleDescription()).ToArray();
            foreach (var module in _modules)
            {
                WireModuleNavigation(module);
                _workspaceCard.Controls.Add(module.Container);
            }
            _contentPanel.Controls.Add(_workspaceCard);
            _scrollHost.Controls.Add(_contentPanel);
            Controls.Add(_scrollHost);
            _scrollHost.Resize += (_, _) => LayoutControls();
            ApplyLanguage(language);
            LayoutControls();
        }

        public void PrepareForDisplay()
        {
            // Registry inventory is read-only and refreshes whenever Home is revisited.
            try { _products = new OfficeInventoryService().DetectInstalledProducts(); _inventoryUnavailable = false; }
            catch (Exception ex) when (ex is System.Security.SecurityException or UnauthorizedAccessException or IOException)
            { _products = Array.Empty<InstalledOfficeProduct>(); _inventoryUnavailable = true; }
            UpdateEnvironmentSummary();
            LayoutControls();
        }

        public void ApplyLanguage(UiLanguage language)
        {
            _language = language;
            _lblTitle.Text = T("Office 运维工作台", "Office Operations");
            _lblHint.Text = T("集中管理 Office 部署、版本与客户端维护。", "Manage Office deployment, versions and client maintenance.");
            _applicationTitle.Text = T("应用管理", "App Management");
            _applicationHint.Text = T("部署、移除与版本控制", "Deployment and version control");
            _diagnosticsTitle.Text = T("诊断与维护", "Diagnostics & Maintenance");
            _diagnosticsHint.Text = T("常见问题处理与诊断报告", "Troubleshooting and reports");
            string[] keys = { "install", "uninstall", "channel", "repair", "teams", "outlook" };
            string[] titles = { T("安装 Office", "Install Office"), T("卸载 Office", "Uninstall Office"), T("更新频道", "Update Channel"),
                T("清理与修复", "Cleanup & Repair"), T("Teams 工具", "Teams Tools"), T("Outlook 工具", "Outlook Tools") };
            string[] summaries = {
                T("部署 M365 应用、Project 与 Visio。\n配置位数、语言与排除应用。", "Deploy M365 Apps, Project and Visio.\nConfigure architecture, languages and apps."),
                T("使用微软官方工具移除 Office。\n支持常规卸载与深度清理。", "Remove Office with Microsoft tools.\nChoose standard removal or deep cleanup."),
                T("切换更新频道或指定目标版本。\n支持 Office 版本回退。", "Switch update channels or target a build.\nRoll back Office versions when needed."),
                T("处理激活信息与账户残留。\n修复代理及网络配置。", "Clear activation and account traces.\nRepair proxy and network settings."),
                T("重置缓存、清理登录记录。\n修复 Outlook 会议插件。", "Reset caches and sign-in records.\nRepair the Outlook meeting add-in."),
                T("运行完整、离线及日历扫描。\n查看诊断报告与日志。", "Run full, offline and calendar scans.\nReview diagnostic reports and logs.") };
            string[] actions = { T("配置安装", "Configure install"), T("打开卸载", "Open uninstall"), T("管理更新", "Manage updates"),
                T("打开修复", "Open repair"), T("打开 Teams 工具", "Open Teams tools"), T("打开 Outlook 工具", "Open Outlook tools") };
            for (int i = 0; i < _modules.Length; i++)
            {
                var module = _modules[i];
                module.PageKey = keys[i];
                module.Icon.Image = GetModuleIcon((ModuleIconKind)i);
                module.Title.Text = titles[i];
                module.Summary.Text = summaries[i];
                module.Detail.Text = actions[i];
                module.Container.AccessibleName = titles[i];
                module.Container.AccessibleDescription = summaries[i];
            }
            UpdateEnvironmentSummary();
            LayoutControls();
        }

        private void UpdateEnvironmentSummary()
        {
            var product = _products.OrderByDescending(p => string.Equals(p.ProductFamily, "Office", StringComparison.OrdinalIgnoreCase))
                .ThenByDescending(p => string.Equals(p.Source, "ClickToRun", StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            _environmentLabels[0].Text = T("当前产品", "Current product");
            _environmentLabels[1].Text = T("已安装版本", "Installed version");
            _environmentLabels[2].Text = T("更新频道", "Update channel");
            string? productName = product != null && product.Source == "ClickToRun" && !string.IsNullOrWhiteSpace(product.ProductId)
                ? OfficeInventoryService.GetClickToRunDisplayName(product.ProductId, _language) : product?.DisplayName;
            _environmentValues[0].Text = productName?.Replace(" (ClickToRun)", "", StringComparison.Ordinal) ?? (_inventoryUnavailable ? T("无法读取", "Unavailable") : T("未检测到 Office", "Office not detected"));
            _environmentValues[1].Text = string.IsNullOrWhiteSpace(product?.Version) ? "—" : product.Version;
            _environmentValues[2].Text = string.IsNullOrWhiteSpace(product?.Channel) ? "—" : product.Channel;
            if (!string.IsNullOrWhiteSpace(product?.Architecture))
            {
                string architecture = product.Architecture.Replace("x64 位", "64 位").Replace("x86 位", "32 位");
                _environmentLabels[1].Text += " · " + (_language == UiLanguage.English ? architecture.Replace(" 位", "-bit") : architecture);
            }
            if (_language == UiLanguage.English)
                _environmentValues[2].Text = _environmentValues[2].Text switch
                {
                    "当前频道" => "Current Channel",
                    "半年频道" or "半年企业频道" => "Semi-Annual Enterprise Channel",
                    "月度企业频道" => "Monthly Enterprise Channel",
                    "月度企业频道（原半年配置）" => "Monthly Enterprise (legacy SAEC)",
                    "当前频道（预览）" => "Current Channel (Preview)",
                    "半年企业频道（预览）" => "Semi-Annual Enterprise Channel (Preview)",
                    "Beta 频道" => "Beta Channel",
                    "其他频道" => "Other channel",
                    var channel => channel
                };
        }

        private ModuleDescriptionControls CreateModuleDescription()
        {
            var panel = new BufferedPanel { BackColor = Color.Transparent, TabStop = true, AccessibleRole = AccessibleRole.PushButton };
            bool hover = false;
            panel.Paint += (_, e) =>
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                using var path = WorkbenchUi.BuildRoundedPath(new Rectangle(0, 0, panel.Width - 1, panel.Height - 1), 10);
                using var fill = new SolidBrush(Color.White);
                e.Graphics.FillPath(fill, path);
                using var border = new Pen(hover ? WorkbenchUi.PrimaryColor : WorkbenchUi.CardBorderColor);
                e.Graphics.DrawPath(border, path);
            };
            panel.SizeChanged += (_, _) =>
            {
                if (panel.Width <= 1 || panel.Height <= 1) return;
                using var path = WorkbenchUi.BuildRoundedPath(panel.ClientRectangle, 10);
                var old = panel.Region; panel.Region = new Region(path); old?.Dispose();
            };
            var icon = new IconPictureBox { BackColor = Color.Transparent, SizeMode = PictureBoxSizeMode.CenterImage, TabStop = false };
            var module = new ModuleDescriptionControls
            {
                Container = panel, Icon = icon,
                Title = CreateLabel(12F, FontStyle.Bold, WorkbenchUi.PrimaryTextColor),
                Summary = CreateLabel(10.5F, FontStyle.Regular, WorkbenchUi.SecondaryTextColor),
                Detail = CreateLabel(9.4F, FontStyle.Regular, WorkbenchUi.PrimaryColor),
                Arrow = CreateLabel(9.75F, FontStyle.Regular, WorkbenchUi.PrimaryColor)
            };
            module.Title.TextAlign = ContentAlignment.MiddleLeft;
            module.Detail.TextAlign = ContentAlignment.MiddleLeft;
            module.Arrow.Text = "\uE76C";
            module.Arrow.Font = WorkbenchUi.CreateIconFont(9.75F);
            module.Arrow.TextAlign = ContentAlignment.MiddleCenter;
            foreach (Control control in GetModuleClickTargets(module))
            {
                control.MouseEnter += (_, _) => { hover = true; panel.Invalidate(); };
                control.MouseLeave += (_, _) =>
                {
                    hover = panel.ClientRectangle.Contains(panel.PointToClient(Cursor.Position)); panel.Invalidate();
                };
            }
            panel.Controls.AddRange(new Control[] { module.Icon, module.Title, module.Summary, module.Detail, module.Arrow });
            return module;
        }

        private void WireModuleNavigation(ModuleDescriptionControls module)
        {
            module.Container.KeyDown += (_, e) =>
            {
                if (e.KeyCode is Keys.Enter or Keys.Space)
                { e.SuppressKeyPress = true; ModuleSelected?.Invoke(this, module.PageKey); }
            };
            foreach (Control control in GetModuleClickTargets(module))
            {
                control.Cursor = Cursors.Hand;
                control.Click += (_, _) => ModuleSelected?.Invoke(this, module.PageKey);
                control.MouseDown += (_, e) => { if (e.Button == MouseButtons.Left) module.Container.Focus(); };
            }
        }

        private static IEnumerable<Control> GetModuleClickTargets(ModuleDescriptionControls module)
        {
            yield return module.Container; yield return module.Icon; yield return module.Title;
            yield return module.Summary; yield return module.Detail; yield return module.Arrow;
        }

        private static Label CreateLabel(float size, FontStyle style, Color color) => new()
        {
            Font = WorkbenchUi.CreateUiFont(size, style), ForeColor = color, BackColor = Color.Transparent,
            AutoSize = false, UseMnemonic = false
        };

        private static Image GetModuleIcon(ModuleIconKind icon) => ModuleIconImages.Value[icon];

        private static IReadOnlyDictionary<ModuleIconKind, Image> LoadModuleIconImages()
        {
            var images = new Dictionary<ModuleIconKind, Image>();
            var assembly = typeof(WorkbenchHomeControl).Assembly;

            foreach (KeyValuePair<ModuleIconKind, string> resource in ModuleIconResourceNames)
            {
                using Stream stream = assembly.GetManifestResourceStream(resource.Value)
                    ?? throw new InvalidOperationException($"Missing home module icon resource: {resource.Value}");
                using Image source = Image.FromStream(stream);
                images[resource.Key] = new Bitmap(source);
            }

            return images;
        }

        private void LayoutControls()
        {
            if (_layingOut || _scrollHost.ClientSize.Width <= 0) return;
            _layingOut = true;
            try
            {
                int viewportWidth = _scrollHost.Width;
                int margin = viewportWidth < 800 ? 24 : 34;
                int width = Math.Max(300, viewportWidth - margin * 2);
                int columns = 1, rowsPerGroup = 3, cardWidth = width, cardHeight = 174;
                for (int pass = 0; pass < 2; pass++)
                {
                    columns = Math.Clamp((width + 18) / 298, 1, 3);
                    rowsPerGroup = (3 + columns - 1) / columns;
                    cardWidth = (width - 18 * (columns - 1)) / columns;
                    int descriptionHeight = _modules.Max(module => TextRenderer.MeasureText(module.Summary.Text,
                        module.Summary.Font, new Size(Math.Max(1, cardWidth - 40), int.MaxValue),
                        TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix).Height);
                    cardHeight = Math.Max(174, 134 + descriptionHeight);
                    int expectedHeight = 302 + rowsPerGroup * (cardHeight + 18) * 2;
                    if (pass == 0 && expectedHeight > _scrollHost.Height)
                        width -= SystemInformation.VerticalScrollBarWidth;
                    else break;
                }
                _workspaceCard.SetBounds(margin, 28, width, 640);
                _lblTitle.SetBounds(0, 0, width, 38);
                _lblHint.SetBounds(0, 42, width, 26);
                _environmentCard.SetBounds(0, 72, width, 90);
                for (int i = 0; i < 3; i++)
                {
                    int x = 23 + width * i / 3;
                    int columnWidth = width / 3 - 46;
                    _environmentLabels[i].SetBounds(x, 16, columnWidth, 22);
                    _environmentValues[i].SetBounds(x, 42, columnWidth, 28);
                }
                int firstTop = 228;
                int secondTop = firstTop + rowsPerGroup * (cardHeight + 18) + 36;
                _applicationTitle.SetBounds(0, firstTop - 40, width / 2, 25);
                _applicationHint.SetBounds(width / 2, firstTop - 40, width - width / 2, 25);
                _diagnosticsTitle.SetBounds(0, secondTop - 40, width / 2, 25);
                _diagnosticsHint.SetBounds(width / 2, secondTop - 40, width - width / 2, 25);
                for (int i = 0; i < 6; i++)
                {
                    var module = _modules[i]; int index = i % 3;
                    module.Container.SetBounds(index % columns * (cardWidth + 18),
                        (i < 3 ? firstTop : secondTop) + index / columns * (cardHeight + 18), cardWidth, cardHeight);
                    module.Icon.SetBounds(20, 21, 36, 36);
                    module.Title.SetBounds(72, 20, cardWidth - 92, 38);
                    module.Summary.SetBounds(20, 74, cardWidth - 40, cardHeight - 134);
                    module.Detail.SetBounds(20, cardHeight - 46, cardWidth - 90, 30);
                    module.Arrow.SetBounds(cardWidth - 42, cardHeight - 46, 22, 30);
                }
                _workspaceCard.Height = _modules.Max(m => m.Container.Bottom);
                _contentPanel.Height = _workspaceCard.Bottom + 28;
                _environmentCard.BringToFront();
            }
            finally { _layingOut = false; }
        }

        private string T(string zh, string en) => LocalizationService.Localize(_language, zh, en);
    }
}
