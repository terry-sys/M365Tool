using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using Office365CleanupTool.Services;

namespace Office365CleanupTool
{
    public class WorkbenchHomeControl : UserControl
    {
        private sealed class BufferedPanel : Panel
        {
            public BufferedPanel()
            {
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.UserPaint, true);
                DoubleBuffered = true;
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
            [ModuleIconKind.Install] = "M365Tool.Assets.Icons.Home.install_exact_preview_tile_128.png",
            [ModuleIconKind.Uninstall] = "M365Tool.Assets.Icons.Home.uninstall_exact_preview_tile_128.png",
            [ModuleIconKind.Channel] = "M365Tool.Assets.Icons.Home.update_channel_exact_preview_tile_128.png",
            [ModuleIconKind.Repair] = "M365Tool.Assets.Icons.Home.cleanup_repair_exact_preview_tile_128.png",
            [ModuleIconKind.Teams] = "M365Tool.Assets.Icons.Home.teams_tools_exact_preview_tile_128.png",
            [ModuleIconKind.Outlook] = "M365Tool.Assets.Icons.Home.outlook_tools_exact_preview_tile_128.png"
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

        private readonly AppSettings _settings;
        private UiLanguage _language;

        private readonly Panel _scrollHost;
        private readonly Panel _contentPanel;
        private readonly Panel _workspaceCard;
        private readonly Label _lblTitle;
        private readonly Label _lblHint;
        private readonly ModuleDescriptionControls[] _modules;

        public event EventHandler<string>? ModuleSelected;

        public WorkbenchHomeControl(
            AppSettingsService settingsService,
            AppSettings settings,
            UiLanguage language)
        {
            AutoScaleMode = AutoScaleMode.None;
            _settings = settings;
            _language = language;

            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint, true);
            BackColor = Color.FromArgb(244, 250, 255);
            Dock = DockStyle.Fill;
            AutoScroll = false;
            DoubleBuffered = true;

            _scrollHost = new BufferedPanel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(244, 250, 255),
                AutoScroll = true
            };
            _scrollHost.Paint += (_, e) => PaintHomeBackground(e.Graphics, _scrollHost.ClientRectangle);

            _contentPanel = new BufferedPanel
            {
                Location = new Point(0, 0),
                BackColor = Color.Transparent,
                Dock = DockStyle.Top
            };

            _workspaceCard = CreateWorkspacePanel();
            _lblTitle = CreateLabel(18.8F, FontStyle.Bold, WorkbenchUi.PrimaryTextColor);
            _lblHint = CreateLabel(11.8F, FontStyle.Regular, WorkbenchUi.SecondaryTextColor);

            _modules = new[]
            {
                CreateModuleDescription(),
                CreateModuleDescription(),
                CreateModuleDescription(),
                CreateModuleDescription(),
                CreateModuleDescription(),
                CreateModuleDescription()
            };

            _workspaceCard.Controls.AddRange(new Control[] { _lblTitle, _lblHint });

            foreach (ModuleDescriptionControls module in _modules)
            {
                WireModuleNavigation(module);
                _workspaceCard.Controls.Add(module.Container);
            }

            _contentPanel.Controls.Add(_workspaceCard);
            _scrollHost.Controls.Add(_contentPanel);
            Controls.Add(_scrollHost);

            _scrollHost.Resize += (_, _) => LayoutControls();

            ApplyLanguage(_language);
            LayoutControls();
        }

        public void PrepareForDisplay()
        {
            LayoutControls();
        }

        public void ApplyLanguage(UiLanguage language)
        {
            _language = language;

            _lblTitle.Text = T("功能说明", "Feature Guide");
            _lblHint.Text = T("了解各模块覆盖的 Microsoft 365 运维场景。", "Review the Microsoft 365 operations covered by each module.");

            ApplyModuleCopy(
                _modules[0],
                "install",
                ModuleIconKind.Install,
                T("安装", "Install"),
                T("标准化安装 M365 应用、Project 与 Visio，支持常用部署参数。", "Standardized deployment for M365 Apps, Project, and Visio."),
                T("支持版本、位数、频道、语言与排除应用配置。", "Supports edition, architecture, channel, language, and app exclusions."));
            ApplyModuleCopy(
                _modules[1],
                "uninstall",
                ModuleIconKind.Uninstall,
                T("卸载", "Uninstall"),
                T("调用微软官方能力卸载 Office，并保留必要处理记录。", "Uses official Microsoft capability to uninstall Office."),
                T("保留处理痕迹，适用于重装前清理和异常恢复。", "Keeps execution records for cleanup and recovery scenarios."));
            ApplyModuleCopy(
                _modules[2],
                "channel",
                ModuleIconKind.Channel,
                T("更新频道", "Update Channel"),
                T("调整 Office 更新频道，支持目标版本切换或回退。", "Adjusts Office channels and supports build rollback."),
                T("适用于版本治理、问题回退与渠道统一。", "Fits version governance, rollback, and channel alignment."));
            ApplyModuleCopy(
                _modules[3],
                "repair",
                ModuleIconKind.Repair,
                T("清理与修复", "Cleanup & Repair"),
                T("清理激活残留、账户痕迹，并修复代理与网络依赖。", "Cleans activation traces, account records, proxy, and network dependencies."),
                T("提升后续安装、激活与登录成功率。", "Improves follow-up install, activation, and sign-in success."));
            ApplyModuleCopy(
                _modules[4],
                "teams",
                ModuleIconKind.Teams,
                T("Teams 工具", "Teams Tools"),
                T("处理 Teams 缓存、登录记录和常见客户端配置异常。", "Handles Teams cache, sign-in records, and common client configuration issues."),
                T("覆盖新版 Teams 启动异常与一线排障。", "Covers new Teams startup issues and frontline troubleshooting."));
            ApplyModuleCopy(
                _modules[5],
                "outlook",
                ModuleIconKind.Outlook,
                T("Outlook 工具", "Outlook Tools"),
                T("执行 Outlook 扫描、诊断、日历检查和日志导出。", "Runs Outlook scans, diagnostics, calendar checks, and log export."),
                T("适用于启动、联机和日历同步问题。", "Fits startup, connectivity, and calendar sync issues."));
        }

        private ModuleDescriptionControls CreateModuleDescription()
        {
            Panel panel = CreateHomeModuleCard();
            PictureBox icon = CreateIconLabel();
            var title = CreateLabel(16.2F, FontStyle.Bold, Color.FromArgb(29, 48, 72));
            var summary = CreateLabel(12.7F, FontStyle.Regular, Color.FromArgb(91, 112, 140));
            var detail = CreateLabel(8.8F, FontStyle.Regular, Color.FromArgb(148, 162, 180));
            detail.Visible = false;
            var arrow = CreateLabel(24F, FontStyle.Regular, WorkbenchUi.PrimaryColor);
            arrow.Text = "\uE76C";
            arrow.Font = WorkbenchUi.CreateIconFont(14F, FontStyle.Regular);
            arrow.TextAlign = ContentAlignment.MiddleCenter;
            arrow.BackColor = Color.Transparent;
            arrow.Visible = true;

            panel.Controls.AddRange(new Control[] { icon, title, summary, detail, arrow });
            return new ModuleDescriptionControls
            {
                Container = panel,
                Icon = icon,
                Title = title,
                Summary = summary,
                Detail = detail,
                Arrow = arrow
            };
        }

        private void WireModuleNavigation(ModuleDescriptionControls module)
        {
            foreach (Control control in GetModuleClickTargets(module))
            {
                control.Cursor = Cursors.Hand;
                control.Click += (_, _) => NavigateToModule(module);
                control.MouseDown += (_, e) =>
                {
                    if (e.Button == MouseButtons.Left)
                    {
                        module.Container.Padding = new Padding(1, 1, 0, 0);
                    }
                };
                control.MouseUp += (_, _) => module.Container.Padding = Padding.Empty;
                control.MouseLeave += (_, _) => module.Container.Padding = Padding.Empty;
            }
        }

        private static IEnumerable<Control> GetModuleClickTargets(ModuleDescriptionControls module)
        {
            yield return module.Container;
            yield return module.Icon;
            yield return module.Title;
            yield return module.Summary;
            yield return module.Detail;
            yield return module.Arrow;
        }

        private void NavigateToModule(ModuleDescriptionControls module)
        {
            if (string.IsNullOrWhiteSpace(module.PageKey))
            {
                return;
            }

            ModuleSelected?.Invoke(this, module.PageKey);
        }

        private static Panel CreateHomeModuleCard()
        {
            var panel = new BufferedPanel
            {
                BackColor = Color.Transparent,
                BorderStyle = BorderStyle.None,
                Margin = Padding.Empty
            };

            void UpdateRegion()
            {
                if (panel.Width <= 1 || panel.Height <= 1)
                {
                    return;
                }

                using var path = BuildRoundedPath(new Rectangle(0, 0, panel.Width, panel.Height), 30);
                panel.Region = new Region(path);
            }

            panel.SizeChanged += (_, _) =>
            {
                UpdateRegion();
                panel.Invalidate();
            };

            panel.Paint += (_, e) =>
            {
                if (panel.Width <= 12 || panel.Height <= 18)
                {
                    return;
                }

                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                using var shadowPath = BuildRoundedPath(new Rectangle(12, 16, panel.Width - 24, panel.Height - 22), 30);
                using var shadowBrush = new SolidBrush(Color.FromArgb(18, 58, 105, 164));
                e.Graphics.FillPath(shadowBrush, shadowPath);
                using var softShadowPath = BuildRoundedPath(new Rectangle(5, 8, panel.Width - 10, panel.Height - 14), 30);
                using var softShadow = new SolidBrush(Color.FromArgb(12, 133, 184, 232));
                e.Graphics.FillPath(softShadow, softShadowPath);

                using var path = BuildRoundedPath(new Rectangle(0, 0, panel.Width - 2, panel.Height - 14), 30);
                using var fill = new LinearGradientBrush(
                    new Rectangle(0, 0, panel.Width - 2, panel.Height - 14),
                    Color.FromArgb(255, 255, 255),
                    Color.FromArgb(251, 254, 255),
                    LinearGradientMode.Vertical);
                e.Graphics.FillPath(fill, path);
                using var border = new Pen(Color.FromArgb(226, 238, 252));
                e.Graphics.DrawPath(border, path);
            };

            UpdateRegion();
            return panel;
        }

        private static PictureBox CreateIconLabel()
        {
            return new PictureBox
            {
                BackColor = Color.Transparent,
                Margin = Padding.Empty,
                SizeMode = PictureBoxSizeMode.Zoom,
                TabStop = false
            };
        }

        private static Label CreateLabel(float size, FontStyle style, Color foreColor)
        {
            return new Label
            {
                AutoSize = false,
                BackColor = Color.Transparent,
                Font = WorkbenchUi.CreateUiFont(size, style),
                ForeColor = foreColor,
                UseMnemonic = false
            };
        }

        private static void ApplyModuleCopy(ModuleDescriptionControls module, string pageKey, ModuleIconKind icon, string title, string summary, string detail)
        {
            module.PageKey = pageKey;
            module.Icon.Image = GetModuleIcon(icon);
            module.Title.Text = title;
            module.Summary.Text = summary;
            module.Detail.Text = detail;
            module.Container.AccessibleName = title;
            module.Container.AccessibleDescription = summary;
            module.Container.AccessibleRole = AccessibleRole.PushButton;
        }

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
            if (_scrollHost.ClientSize.Width <= 0 || _scrollHost.ClientSize.Height <= 0)
            {
                return;
            }

            int viewportWidth = _scrollHost.ClientSize.Width;
            int viewportHeight = _scrollHost.ClientSize.Height;
            float scale = WorkbenchUi.GetAdaptiveUiScale(viewportWidth, viewportHeight);
            int S(int value) => WorkbenchUi.Scale(value, scale);

            bool compact = viewportHeight < 690 || viewportWidth < 980;
            int margin = WorkbenchUi.GetAdaptivePageMargin(viewportWidth);
            int availableWidth = WorkbenchUi.GetAdaptiveContentWidth(
                viewportWidth,
                520,
                WorkbenchUi.WideContentMaxWidth,
                margin,
                scrollbarAllowance: SystemInformation.VerticalScrollBarWidth);
            int contentLeft = WorkbenchUi.GetAdaptiveContentLeft(viewportWidth, availableWidth, margin);
            int workspaceTop = S(compact ? 30 : 34);
            int titleTop = S(compact ? 34 : 42);
            int moduleTop = S(compact ? 108 : 132);
            int cardGap = S(compact ? 22 : 30);
            int workspaceBottomPadding = S(compact ? 30 : 50);
            int columnCount = availableWidth >= 980 ? 3 : availableWidth >= 680 ? 2 : 1;
            int rowCount = (int)Math.Ceiling(_modules.Length / (double)columnCount);
            int visibleRows = Math.Min(rowCount, 2);
            int moduleHeight = compact
                ? Math.Max(S(204), Math.Min(S(232), (viewportHeight - workspaceTop - S(14) - moduleTop - cardGap - workspaceBottomPadding) / Math.Max(1, visibleRows)))
                : Math.Max(S(238), Math.Min(S(314), (viewportHeight - S(298)) / 2));
            bool roomy = moduleHeight > S(260) && columnCount == 3;
            int innerLeft = S(availableWidth < 680 ? 30 : compact ? 42 : roomy ? 56 : 48);
            int innerWidth = availableWidth - innerLeft * 2;
            int columnWidth = (innerWidth - cardGap * (columnCount - 1)) / columnCount;

            _workspaceCard.Location = new Point(contentLeft, workspaceTop);
            _workspaceCard.Size = new Size(availableWidth, S(620));

            ApplyLabelFont(_lblTitle, 18.8F, FontStyle.Bold, scale);
            ApplyLabelFont(_lblHint, 11.8F, FontStyle.Regular, scale);
            _lblTitle.Location = new Point(innerLeft, titleTop);
            _lblTitle.Size = new Size(innerWidth, S(36));
            _lblHint.Location = new Point(innerLeft, titleTop + S(42));
            _lblHint.Size = new Size(innerWidth, S(30));

            int top = compact ? moduleTop : roomy ? S(166) : S(132);
            for (int i = 0; i < _modules.Length; i++)
            {
                int column = i % columnCount;
                int row = i / columnCount;
                int left = innerLeft + column * (columnWidth + cardGap);
                int cardTop = top + row * (moduleHeight + cardGap);

                LayoutModuleDescription(_modules[i], left, cardTop, columnWidth, moduleHeight, scale);
            }

            int contentBottom = top + rowCount * moduleHeight + (rowCount - 1) * cardGap;
            _workspaceCard.Height = contentBottom + (compact ? workspaceBottomPadding : S(roomy ? 58 : 50));
            _contentPanel.Width = Math.Max(viewportWidth - 1, contentLeft + availableWidth + margin);
            _contentPanel.Height = _workspaceCard.Bottom + workspaceTop;
        }

        private static void LayoutModuleDescription(ModuleDescriptionControls module, int left, int top, int width, int height, float scale)
        {
            int S(int value) => WorkbenchUi.Scale(value, scale);

            module.Container.Location = new Point(left, top);
            module.Container.Size = new Size(width, height);
            bool roomy = height > S(260);
            bool compact = height < S(230);
            int leftPad = S(compact ? 30 : roomy ? 38 : 34);
            int rightPad = leftPad;
            int iconSize = S(compact ? 46 : roomy ? 58 : 52);
            int iconTop = S(compact ? 24 : roomy ? 34 : 30);
            int titleHeight = iconSize;
            int titleGap = S(compact ? 14 : 18);
            int titleLeft = leftPad + iconSize + titleGap;
            int titleTop = iconTop;
            int summaryTop = iconTop + iconSize + S(compact ? 18 : roomy ? 24 : 20);
            int bottomPad = S(compact ? 18 : 24);

            ApplyLabelFont(module.Title, compact ? 17.2F : roomy ? 20.2F : 18.8F, FontStyle.Bold, scale);
            ApplyLabelFont(module.Summary, compact ? 11.2F : 12.7F, FontStyle.Regular, scale);
            WorkbenchUi.ApplyIconFont(module.Arrow, 13F, FontStyle.Regular, scale);

            module.Icon.Location = new Point(leftPad, iconTop);
            module.Icon.Size = new Size(iconSize, iconSize);
            module.Title.Location = new Point(titleLeft, titleTop);
            module.Title.Size = new Size(Math.Max(80, width - titleLeft - rightPad), titleHeight);
            module.Title.TextAlign = ContentAlignment.MiddleLeft;
            module.Summary.Location = new Point(leftPad, summaryTop);
            module.Summary.Size = new Size(width - leftPad - rightPad, Math.Max(S(54), height - summaryTop - bottomPad));
            module.Detail.Location = new Point(leftPad, S(202));
            module.Detail.Size = new Size(width - S(112), S(24));
            module.Arrow.Location = new Point(width - S(68), roomy ? height - S(66) : height - S(54));
            module.Arrow.Size = new Size(S(34), S(34));
        }

        private static void ApplyLabelFont(Label label, float size, FontStyle style, float scale = 1F)
        {
            WorkbenchUi.ApplyUiFont(label, size, style, scale);
        }

        private static Panel CreateWorkspacePanel()
        {
            var panel = new BufferedPanel
            {
                BackColor = Color.Transparent,
                BorderStyle = BorderStyle.None
            };

            void UpdateRegion()
            {
                if (panel.Width <= 1 || panel.Height <= 1)
                {
                    return;
                }

                using var path = BuildRoundedPath(new Rectangle(0, 0, panel.Width, panel.Height), 38);
                panel.Region = new Region(path);
            }

            panel.SizeChanged += (_, _) =>
            {
                UpdateRegion();
                panel.Invalidate();
            };

            panel.Paint += (_, e) =>
            {
                if (panel.Width <= 12 || panel.Height <= 18)
                {
                    return;
                }

                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                using var ambientShadowPath = BuildRoundedPath(new Rectangle(8, 12, panel.Width - 16, panel.Height - 18), 40);
                using var ambientShadow = new SolidBrush(Color.FromArgb(16, 86, 136, 198));
                e.Graphics.FillPath(ambientShadow, ambientShadowPath);

                using var shadowPath = BuildRoundedPath(new Rectangle(18, 24, panel.Width - 36, panel.Height - 34), 40);
                using var shadowBrush = new SolidBrush(Color.FromArgb(12, 64, 116, 176));
                e.Graphics.FillPath(shadowBrush, shadowPath);

                using var path = BuildRoundedPath(new Rectangle(0, 0, panel.Width - 2, panel.Height - 4), 40);
                using var fill = new LinearGradientBrush(
                    new Rectangle(0, 0, panel.Width, panel.Height),
                    Color.FromArgb(248, 252, 255),
                    Color.FromArgb(240, 248, 255),
                    LinearGradientMode.Vertical);
                using var pen = new Pen(Color.FromArgb(220, 236, 252));
                e.Graphics.FillPath(fill, path);
                e.Graphics.DrawPath(pen, path);
            };

            UpdateRegion();
            return panel;
        }

        private static void PaintHomeBackground(Graphics graphics, Rectangle bounds)
        {
            if (bounds.Width <= 0 || bounds.Height <= 0)
            {
                return;
            }

            graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using var brush = new LinearGradientBrush(
                bounds,
                Color.FromArgb(255, 255, 255),
                Color.FromArgb(226, 243, 255),
                LinearGradientMode.ForwardDiagonal);
            graphics.FillRectangle(brush, bounds);

            using var glowBrush = new SolidBrush(Color.FromArgb(38, 219, 240, 255));
            graphics.FillEllipse(glowBrush, new Rectangle(bounds.Width - 520, bounds.Height - 360, 680, 460));
            using var blueGlow = new SolidBrush(Color.FromArgb(18, 139, 198, 255));
            graphics.FillEllipse(blueGlow, new Rectangle(40, 10, 560, 340));
            using var topMist = new SolidBrush(Color.FromArgb(48, 246, 251, 255));
            graphics.FillEllipse(topMist, new Rectangle(bounds.Width / 4, -150, bounds.Width, 330));
        }

        private string T(string zh, string en) => LocalizationService.Localize(_language, zh, en);

        private static System.Drawing.Drawing2D.GraphicsPath BuildRoundedPath(Rectangle rect, int radius)
        {
            int r = Math.Max(1, radius);
            int d = r * 2;
            var path = new System.Drawing.Drawing2D.GraphicsPath();

            Rectangle arc = new Rectangle(rect.X, rect.Y, d, d);
            path.AddArc(arc, 180, 90);
            arc.X = rect.Right - d;
            path.AddArc(arc, 270, 90);
            arc.Y = rect.Bottom - d;
            path.AddArc(arc, 0, 90);
            arc.X = rect.X;
            path.AddArc(arc, 90, 90);
            path.CloseFigure();
            return path;
        }

        private static void ApplyRoundedRegion(Control control, int radius)
        {
            void UpdateRegion()
            {
                if (control.Width <= 1 || control.Height <= 1)
                {
                    return;
                }

                using var path = BuildRoundedPath(new Rectangle(0, 0, control.Width, control.Height), radius);
                control.Region = new Region(path);
            }

            control.SizeChanged += (_, _) => UpdateRegion();
            UpdateRegion();
        }
    }
}
