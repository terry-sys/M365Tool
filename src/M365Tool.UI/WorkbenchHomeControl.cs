using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
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

        private sealed class ModuleDescriptionControls
        {
            public Panel Container { get; init; } = null!;
            public Label Icon { get; init; } = null!;
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

        public WorkbenchHomeControl(
            AppSettingsService settingsService,
            AppSettings settings,
            UiLanguage language)
        {
            AutoScaleMode = AutoScaleMode.Dpi;
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
                AutoScroll = false
            };
            _scrollHost.Paint += (_, e) => PaintHomeBackground(e.Graphics, _scrollHost.ClientRectangle);

            _contentPanel = new BufferedPanel
            {
                Location = new Point(0, 0),
                BackColor = Color.Transparent,
                Dock = DockStyle.Fill
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
                ModuleIconKind.Install,
                T("安装", "Install"),
                T("标准化安装 M365 应用、Project 与 Visio，支持常用部署参数。", "Standardized deployment for M365 Apps, Project, and Visio."),
                T("支持版本、位数、频道、语言与排除应用配置。", "Supports edition, architecture, channel, language, and app exclusions."));
            ApplyModuleCopy(
                _modules[1],
                ModuleIconKind.Uninstall,
                T("卸载", "Uninstall"),
                T("调用微软官方能力卸载 Office，并保留必要处理记录。", "Uses official Microsoft capability to uninstall Office."),
                T("保留处理痕迹，适用于重装前清理和异常恢复。", "Keeps execution records for cleanup and recovery scenarios."));
            ApplyModuleCopy(
                _modules[2],
                ModuleIconKind.Channel,
                T("更新频道", "Update Channel"),
                T("调整 Office 更新频道，支持目标版本切换或回退。", "Adjusts Office channels and supports build rollback."),
                T("适用于版本治理、问题回退与渠道统一。", "Fits version governance, rollback, and channel alignment."));
            ApplyModuleCopy(
                _modules[3],
                ModuleIconKind.Repair,
                T("清理与修复", "Cleanup & Repair"),
                T("清理激活残留、账户痕迹，并修复代理与网络依赖。", "Cleans activation traces, account records, proxy, and network dependencies."),
                T("提升后续安装、激活与登录成功率。", "Improves follow-up install, activation, and sign-in success."));
            ApplyModuleCopy(
                _modules[4],
                ModuleIconKind.Teams,
                T("Teams 工具", "Teams Tools"),
                T("处理 Teams 缓存、登录记录和常见客户端配置异常。", "Handles Teams cache, sign-in records, and common client configuration issues."),
                T("覆盖新版 Teams 启动异常与一线排障。", "Covers new Teams startup issues and frontline troubleshooting."));
            ApplyModuleCopy(
                _modules[5],
                ModuleIconKind.Outlook,
                T("Outlook 工具", "Outlook Tools"),
                T("执行 Outlook 扫描、诊断、日历检查和日志导出。", "Runs Outlook scans, diagnostics, calendar checks, and log export."),
                T("适用于启动、联机和日历同步问题。", "Fits startup, connectivity, and calendar sync issues."));
        }

        private ModuleDescriptionControls CreateModuleDescription()
        {
            Panel panel = CreateHomeModuleCard();
            var icon = CreateIconLabel();
            var title = CreateLabel(16.2F, FontStyle.Bold, Color.FromArgb(29, 48, 72));
            var summary = CreateLabel(12.7F, FontStyle.Regular, Color.FromArgb(91, 112, 140));
            var detail = CreateLabel(8.8F, FontStyle.Regular, Color.FromArgb(148, 162, 180));
            detail.Visible = false;
            var arrow = CreateLabel(24F, FontStyle.Regular, WorkbenchUi.PrimaryColor);
            arrow.Text = string.Empty;
            arrow.TextAlign = ContentAlignment.MiddleCenter;
            arrow.BackColor = Color.Transparent;
            arrow.Visible = false;

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

        private Label CreateIconLabel()
        {
            var icon = new Label
            {
                AutoSize = false,
                BackColor = Color.FromArgb(235, 245, 255),
                ForeColor = WorkbenchUi.PrimaryColor,
                Font = new Font("Segoe MDL2 Assets", 17F, FontStyle.Regular),
                TextAlign = ContentAlignment.MiddleCenter,
                UseMnemonic = false
            };

            ApplyRoundedRegion(icon, 18);
            return icon;
        }

        private static Label CreateLabel(float size, FontStyle style, Color foreColor)
        {
            return new Label
            {
                AutoSize = false,
                BackColor = Color.Transparent,
                Font = new Font("Microsoft YaHei UI", size, style),
                ForeColor = foreColor,
                UseMnemonic = false
            };
        }

        private static void ApplyModuleCopy(ModuleDescriptionControls module, ModuleIconKind icon, string title, string summary, string detail)
        {
            Image? oldImage = module.Icon.Image;
            module.Icon.Image = null;
            oldImage?.Dispose();
            module.Icon.Text = GetModuleGlyph(icon);
            module.Title.Text = title;
            module.Summary.Text = summary;
            module.Detail.Text = detail;
        }

        private static string GetModuleGlyph(ModuleIconKind icon)
        {
            return icon switch
            {
                ModuleIconKind.Install => "\uE7C3",
                ModuleIconKind.Uninstall => "\uE74D",
                ModuleIconKind.Channel => "\uE8D7",
                ModuleIconKind.Repair => "\uE90F",
                ModuleIconKind.Teams => "\uE720",
                ModuleIconKind.Outlook => "\uE715",
                _ => "\uE10C"
            };
        }

        private static Bitmap CreateModuleIconBitmap(ModuleIconKind icon, Color color, Size size)
        {
            var bitmap = new Bitmap(size.Width, size.Height);
            using Graphics graphics = Graphics.FromImage(bitmap);
            graphics.Clear(Color.Transparent);
            graphics.SmoothingMode = SmoothingMode.AntiAlias;

            using var brush = new SolidBrush(color);
            using var lightBrush = new SolidBrush(Color.FromArgb(84, color));
            using var pen = new Pen(color, 2.4F)
            {
                StartCap = LineCap.Round,
                EndCap = LineCap.Round,
                LineJoin = LineJoin.Round
            };

            RectangleF r = new RectangleF(4, 4, size.Width - 8, size.Height - 8);
            switch (icon)
            {
                case ModuleIconKind.Install:
                    PointF[] top =
                    {
                        new(r.Left + r.Width * 0.50F, r.Top + r.Height * 0.06F),
                        new(r.Left + r.Width * 0.86F, r.Top + r.Height * 0.26F),
                        new(r.Left + r.Width * 0.50F, r.Top + r.Height * 0.46F),
                        new(r.Left + r.Width * 0.14F, r.Top + r.Height * 0.26F)
                    };
                    PointF[] left =
                    {
                        new(r.Left + r.Width * 0.14F, r.Top + r.Height * 0.30F),
                        new(r.Left + r.Width * 0.50F, r.Top + r.Height * 0.50F),
                        new(r.Left + r.Width * 0.50F, r.Top + r.Height * 0.94F),
                        new(r.Left + r.Width * 0.14F, r.Top + r.Height * 0.72F)
                    };
                    PointF[] right =
                    {
                        new(r.Left + r.Width * 0.86F, r.Top + r.Height * 0.30F),
                        new(r.Left + r.Width * 0.50F, r.Top + r.Height * 0.50F),
                        new(r.Left + r.Width * 0.50F, r.Top + r.Height * 0.94F),
                        new(r.Left + r.Width * 0.86F, r.Top + r.Height * 0.72F)
                    };
                    graphics.FillPolygon(lightBrush, top);
                    graphics.FillPolygon(new SolidBrush(Color.FromArgb(220, color)), left);
                    graphics.FillPolygon(brush, right);
                    break;

                case ModuleIconKind.Uninstall:
                    graphics.FillRectangle(brush, r.Left + r.Width * 0.30F, r.Top + r.Height * 0.28F, r.Width * 0.40F, r.Height * 0.06F);
                    graphics.FillRectangle(brush, r.Left + r.Width * 0.42F, r.Top + r.Height * 0.14F, r.Width * 0.16F, r.Height * 0.08F);
                    using (GraphicsPath bin = BuildRoundedPath(Rectangle.Round(new RectangleF(r.Left + r.Width * 0.27F, r.Top + r.Height * 0.38F, r.Width * 0.46F, r.Height * 0.48F)), 5))
                    {
                        graphics.FillPath(lightBrush, bin);
                        graphics.DrawPath(pen, bin);
                    }
                    graphics.DrawLine(pen, r.Left + r.Width * 0.43F, r.Top + r.Height * 0.46F, r.Left + r.Width * 0.43F, r.Top + r.Height * 0.76F);
                    graphics.DrawLine(pen, r.Left + r.Width * 0.57F, r.Top + r.Height * 0.46F, r.Left + r.Width * 0.57F, r.Top + r.Height * 0.76F);
                    break;

                case ModuleIconKind.Channel:
                    graphics.DrawArc(pen, r.Left + 2, r.Top + 4, r.Width - 6, r.Height - 8, 35, 225);
                    graphics.DrawArc(pen, r.Left + 4, r.Top + 4, r.Width - 8, r.Height - 8, 215, 225);
                    graphics.FillPolygon(brush, new[]
                    {
                        new PointF(r.Right - 3, r.Top + r.Height * 0.24F),
                        new PointF(r.Right - 11, r.Top + r.Height * 0.18F),
                        new PointF(r.Right - 9, r.Top + r.Height * 0.36F)
                    });
                    graphics.FillPolygon(brush, new[]
                    {
                        new PointF(r.Left + 3, r.Top + r.Height * 0.76F),
                        new PointF(r.Left + 11, r.Top + r.Height * 0.82F),
                        new PointF(r.Left + 9, r.Top + r.Height * 0.64F)
                    });
                    break;

                case ModuleIconKind.Repair:
                    graphics.DrawLine(pen, r.Left + r.Width * 0.26F, r.Top + r.Height * 0.76F, r.Left + r.Width * 0.70F, r.Top + r.Height * 0.32F);
                    graphics.DrawEllipse(pen, r.Left + r.Width * 0.58F, r.Top + r.Height * 0.14F, r.Width * 0.28F, r.Height * 0.28F);
                    graphics.DrawLine(pen, r.Left + r.Width * 0.24F, r.Top + r.Height * 0.78F, r.Left + r.Width * 0.15F, r.Top + r.Height * 0.69F);
                    graphics.DrawLine(pen, r.Left + r.Width * 0.35F, r.Top + r.Height * 0.86F, r.Left + r.Width * 0.27F, r.Top + r.Height * 0.78F);
                    break;

                case ModuleIconKind.Teams:
                    using (var teamsFont = new Font("Segoe UI", 17F, FontStyle.Bold, GraphicsUnit.Pixel))
                    {
                        graphics.DrawString("T", teamsFont, brush, r.Left + r.Width * 0.10F, r.Top + r.Height * 0.24F);
                    }
                    graphics.FillEllipse(lightBrush, r.Left + r.Width * 0.60F, r.Top + r.Height * 0.16F, r.Width * 0.22F, r.Height * 0.22F);
                    graphics.FillEllipse(brush, r.Left + r.Width * 0.66F, r.Top + r.Height * 0.46F, r.Width * 0.24F, r.Height * 0.30F);
                    using (GraphicsPath tCard = BuildRoundedPath(Rectangle.Round(new RectangleF(r.Left + r.Width * 0.18F, r.Top + r.Height * 0.34F, r.Width * 0.48F, r.Height * 0.42F)), 5))
                    {
                        graphics.FillPath(new SolidBrush(Color.FromArgb(210, color)), tCard);
                    }
                    break;

                case ModuleIconKind.Outlook:
                    using (GraphicsPath app = BuildRoundedPath(Rectangle.Round(new RectangleF(r.Left + r.Width * 0.16F, r.Top + r.Height * 0.26F, r.Width * 0.38F, r.Height * 0.50F)), 5))
                    {
                        graphics.FillPath(brush, app);
                    }
                    using (var outlookFont = new Font("Segoe UI", 13F, FontStyle.Bold, GraphicsUnit.Pixel))
                    using (var whiteBrush = new SolidBrush(Color.White))
                    {
                        graphics.DrawString("O", outlookFont, whiteBrush, r.Left + r.Width * 0.24F, r.Top + r.Height * 0.36F);
                    }
                    using (GraphicsPath mail = BuildRoundedPath(Rectangle.Round(new RectangleF(r.Left + r.Width * 0.44F, r.Top + r.Height * 0.34F, r.Width * 0.42F, r.Height * 0.34F)), 4))
                    {
                        graphics.FillPath(lightBrush, mail);
                        graphics.DrawPath(pen, mail);
                    }
                    graphics.DrawLine(pen, r.Left + r.Width * 0.47F, r.Top + r.Height * 0.38F, r.Left + r.Width * 0.65F, r.Top + r.Height * 0.54F);
                    graphics.DrawLine(pen, r.Left + r.Width * 0.83F, r.Top + r.Height * 0.38F, r.Left + r.Width * 0.65F, r.Top + r.Height * 0.54F);
                    break;
            }

            return bitmap;
        }

        private void LayoutControls()
        {
            if (_scrollHost.ClientSize.Width <= 0 || _scrollHost.ClientSize.Height <= 0)
            {
                return;
            }

            int viewportWidth = _scrollHost.ClientSize.Width;
            int viewportHeight = _scrollHost.ClientSize.Height;
            bool compact = viewportHeight < 690;
            int margin = compact ? 38 : 46;
            int availableWidth = Math.Max(1040, Math.Min(1600, viewportWidth - margin * 2));
            int contentLeft = Math.Max(margin, (viewportWidth - availableWidth) / 2);
            int workspaceTop = compact ? 38 : 58;
            int titleTop = compact ? 34 : 42;
            int moduleTop = compact ? 108 : 132;
            int cardGap = compact ? 22 : 30;
            int workspaceBottomPadding = compact ? 30 : 50;
            int moduleHeight = compact
                ? Math.Max(190, Math.Min(214, (viewportHeight - workspaceTop - 14 - moduleTop - cardGap - workspaceBottomPadding) / 2))
                : Math.Max(238, Math.Min(314, (viewportHeight - 298) / 2));
            bool roomy = moduleHeight > 260;
            int innerLeft = compact ? 42 : roomy ? 56 : 48;
            int innerWidth = availableWidth - innerLeft * 2;
            int columnWidth = (innerWidth - cardGap * 2) / 3;

            _workspaceCard.Location = new Point(contentLeft, workspaceTop);
            _workspaceCard.Size = new Size(availableWidth, 620);

            _lblTitle.Location = new Point(innerLeft, titleTop);
            _lblTitle.Size = new Size(innerWidth, 34);
            _lblHint.Location = new Point(innerLeft, titleTop + 42);
            _lblHint.Size = new Size(innerWidth, 28);

            int top = compact ? moduleTop : roomy ? 166 : 132;
            for (int i = 0; i < _modules.Length; i++)
            {
                int column = i % 3;
                int row = i / 3;
                int left = innerLeft + column * (columnWidth + cardGap);
                int cardTop = top + row * (moduleHeight + cardGap);

                LayoutModuleDescription(_modules[i], left, cardTop, columnWidth, moduleHeight);
            }

            int rowCount = (_modules.Length + 2) / 3;
            int contentBottom = top + rowCount * moduleHeight + (rowCount - 1) * cardGap;
            _workspaceCard.Height = contentBottom + (compact ? workspaceBottomPadding : roomy ? 58 : 50);
        }

        private static void LayoutModuleDescription(ModuleDescriptionControls module, int left, int top, int width, int height)
        {
            module.Container.Location = new Point(left, top);
            module.Container.Size = new Size(width, height);
            bool roomy = height > 260;
            bool compact = height < 230;
            int leftPad = compact ? 34 : 36;
            int iconSize = compact ? 48 : roomy ? 76 : 58;
            int iconTop = compact ? 24 : roomy ? 38 : 32;
            int titleTop = compact ? 82 : roomy ? 130 : 112;
            int titleHeight = compact ? 28 : 38;
            int summaryTop = compact ? 118 : roomy ? 176 : 146;

            ApplyLabelFont(module.Title, compact ? 14.2F : 16.2F, FontStyle.Bold);
            ApplyLabelFont(module.Summary, compact ? 11.2F : 12.7F, FontStyle.Regular);

            module.Icon.Location = new Point(leftPad, iconTop);
            module.Icon.Size = new Size(iconSize, iconSize);
            ApplyRoundedRegion(module.Icon, compact ? 16 : roomy ? 20 : 18);
            module.Title.Location = new Point(leftPad, titleTop);
            module.Title.Size = new Size(width - leftPad * 2, titleHeight);
            module.Summary.Location = new Point(leftPad, summaryTop);
            module.Summary.Size = new Size(width - leftPad * 2, compact ? height - summaryTop - 12 : roomy ? 86 : 72);
            module.Detail.Location = new Point(leftPad, 202);
            module.Detail.Size = new Size(width - 112, 24);
            module.Arrow.Location = new Point(width - 68, roomy ? height - 66 : height - 54);
            module.Arrow.Size = new Size(34, 34);
        }

        private static void ApplyLabelFont(Label label, float size, FontStyle style)
        {
            if (Math.Abs(label.Font.Size - size) < 0.01F && label.Font.Style == style)
            {
                return;
            }

            Font oldFont = label.Font;
            label.Font = new Font("Microsoft YaHei UI", size, style);
            oldFont.Dispose();
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
