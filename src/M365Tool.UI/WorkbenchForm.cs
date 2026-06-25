using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using Office365CleanupTool.Services;

namespace Office365CleanupTool
{
    public class WorkbenchForm : Form
    {
        private const int ExpandedSidebarWidth = 258;
        private const int CollapsedSidebarWidth = 78;
        private const int SidebarBottomMargin = 16;
        private const int SidebarNavGap = 8;
        private const int CollapsedNavButtonHeight = 58;

        private readonly IScriptRunner _scriptRunner = new ScriptRunner();
        private readonly Panel _sidebar;
        private readonly Panel _headerPanel;
        private readonly Panel _contentHost;
        private readonly Panel _statusBar;
        private readonly Panel _sidebarNavSurface;
        private readonly Panel _brandPanel;
        private readonly Label _lblBrandGlyph;
        private readonly Label _lblAppTitle;
        private readonly Label _lblAppSubtitle;
        private readonly Label _lblPageTitle;
        private readonly Label _lblPageDescription;
        private readonly Panel _accountPanel;
        private readonly Label _lblAccountAvatar;
        private readonly Label _lblAccountName;
        private readonly Label _lblAccountState;
        private readonly Panel _accountDivider;
        private readonly Button _btnSignIn;
        private readonly Button _btnSignOut;
        private Image? _accountAvatarImage;
        private readonly ProgressBar _progressBar;
        private readonly Button _btnToggleSidebar;
        private readonly Panel _navSelectionBar;
        private readonly Dictionary<string, Control> _pageCache = new();
        private readonly Dictionary<string, Button> _navButtons = new();
        private readonly Dictionary<string, string> _navButtonTexts = new();
        private readonly Dictionary<string, string> _navButtonIcons = new();
        private static readonly IReadOnlyDictionary<string, string> NavPngIconResourceNames = new Dictionary<string, string>
        {
            ["install"] = "M365Tool.Assets.Icons.Home.install_exact_preview_tile_128.png",
            ["uninstall"] = "M365Tool.Assets.Icons.Home.uninstall_exact_preview_tile_128.png",
            ["channel"] = "M365Tool.Assets.Icons.Home.update_channel_exact_preview_tile_128.png",
            ["repair"] = "M365Tool.Assets.Icons.Home.cleanup_repair_exact_preview_tile_128.png",
            ["teams"] = "M365Tool.Assets.Icons.Home.teams_tools_exact_preview_tile_128.png",
            ["outlook"] = "M365Tool.Assets.Icons.Home.outlook_tools_exact_preview_tile_128.png"
        };

        private static readonly Lazy<IReadOnlyDictionary<string, Image>> NavPngIconImages = new(() => LoadNavPngIconImages(1F));
        private static readonly Lazy<IReadOnlyDictionary<string, Image>> DisabledNavPngIconImages = new(() => LoadNavPngIconImages(0.38F));
        private readonly ToolTip _navToolTip = new()
        {
            ShowAlways = true,
            InitialDelay = 180,
            ReshowDelay = 100,
            AutoPopDelay = 4000
        };

        private readonly TeamsMaintenanceService _teamsMaintenanceService = new();
        private readonly OutlookDiagnosticsService _outlookDiagnosticsService = new();
        private readonly OfficeUninstallService _officeUninstallService;
        private readonly OfficeChannelService _officeChannelService = new();
        private readonly CleanupService _cleanupService;
        private readonly NetworkRepairService _networkRepairService;
        private readonly AppSettingsService _appSettingsService;
        private readonly AppSettings _appSettings;

        private bool _isSidebarCollapsed;
        private bool _sessionVerified;
        private UiLanguage _currentLanguage;
        private string _currentPageKey = "home";
        public WorkbenchForm()
        {
            AutoScaleMode = AutoScaleMode.None;
            _officeUninstallService = new OfficeUninstallService(_scriptRunner);
            _cleanupService = new CleanupService(_scriptRunner);
            _networkRepairService = new NetworkRepairService(_scriptRunner);
            _appSettingsService = new AppSettingsService();
            _appSettings = _appSettingsService.Load();
            ResetStaleValidatedProfile();
            _currentLanguage = LocalizationService.ResolveLanguage(_appSettings.LanguageMode);
            _sessionVerified = _appSettings.HasValidated21VAccount &&
                               (!string.IsNullOrWhiteSpace(_appSettings.LastValidatedAccount) ||
                                !string.IsNullOrWhiteSpace(_appSettings.LastValidatedDisplayName));

            Text = "21Vianet Office工具箱";
            Icon = AppIconProvider.GetAppIcon() ?? Icon;
            ShowIcon = true;
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(1024, 720);
            ClientSize = GetInitialClientSize();
            BackColor = Color.White;

            _sidebar = new Panel
            {
                Dock = DockStyle.Left,
                Width = ExpandedSidebarWidth,
                BackColor = Color.Transparent
            };
            _sidebar.Paint += (_, e) => PaintSidebarBackground(e.Graphics, _sidebar.ClientRectangle);

            _navSelectionBar = new Panel
            {
                Size = new Size(3, 24),
                BackColor = WorkbenchUi.PrimaryColor,
                Visible = false
            };

            _sidebarNavSurface = CreateSidebarNavSurface();
            _brandPanel = new Panel
            {
                BackColor = Color.Transparent,
                Location = Point.Empty,
                Size = Size.Empty
            };
            _brandPanel.Paint += (_, e) => PaintBrandLogo(e.Graphics, _brandPanel.ClientRectangle);
            _lblBrandGlyph = new Label
            {
                Text = "21V",
                Font = WorkbenchUi.CreateFont("Segoe UI", 23F, FontStyle.Bold),
                ForeColor = WorkbenchUi.PrimaryColor,
                BackColor = Color.Transparent,
                TextAlign = ContentAlignment.MiddleLeft,
                Size = new Size(78, 48),
                Visible = false
            };

            _headerPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 194,
                BackColor = Color.Transparent
            };
            _headerPanel.Paint += (_, e) => PaintMainCanvas(e.Graphics, _headerPanel.ClientRectangle, true);

            _statusBar = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 44,
                BackColor = Color.White,
                Visible = false
            };

            _contentHost = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent,
                Padding = Padding.Empty,
                AutoScroll = false
            };

            _btnToggleSidebar = new Button
            {
                Text = "\uE700",
                Font = WorkbenchUi.CreateIconFont(16.5F, FontStyle.Regular),
                ForeColor = WorkbenchUi.SecondaryTextColor,
                BackColor = Color.FromArgb(246, 251, 255),
                FlatStyle = FlatStyle.Flat,
                Location = new Point(22, 26),
                Size = new Size(48, 48),
                TextAlign = ContentAlignment.MiddleCenter,
                Cursor = Cursors.Hand,
                TabStop = false
            };
            _btnToggleSidebar.FlatAppearance.BorderSize = 0;
            _btnToggleSidebar.FlatAppearance.MouseOverBackColor = Color.FromArgb(241, 247, 255);
            _btnToggleSidebar.FlatAppearance.MouseDownBackColor = Color.FromArgb(232, 241, 253);
            _btnToggleSidebar.Click += (_, _) => ToggleSidebar();
            _navToolTip.SetToolTip(_btnToggleSidebar, T("展开或折叠导航", "Expand or collapse navigation"));
            ApplyRoundedRegion(_btnToggleSidebar, 24);

            _lblAppTitle = new Label
            {
                Text = "Office工具箱",
                Font = WorkbenchUi.CreateUiFont(12.2F, FontStyle.Bold),
                ForeColor = WorkbenchUi.PrimaryTextColor,
                BackColor = Color.Transparent,
                Location = new Point(104, 24),
                Size = new Size(116, 24)
            };

            _lblAppSubtitle = new Label
            {
                Text = "21Vianet Office 运维工具",
                Font = WorkbenchUi.CreateUiFont(8.4F),
                ForeColor = Color.FromArgb(106, 124, 148),
                BackColor = Color.Transparent,
                Location = new Point(104, 50),
                Size = new Size(136, 18)
            };

            _lblPageTitle = new Label
            {
                Text = "首页",
                Font = WorkbenchUi.CreateUiFont(24F, FontStyle.Bold),
                ForeColor = WorkbenchUi.PrimaryTextColor,
                BackColor = Color.Transparent,
                Location = new Point(64, 86),
                Size = new Size(520, 44)
            };

            _lblPageDescription = new Label
            {
                Text = string.Empty,
                Font = WorkbenchUi.CreateUiFont(11F),
                ForeColor = WorkbenchUi.SecondaryTextColor,
                BackColor = Color.Transparent,
                Location = new Point(66, 132),
                Size = new Size(920, 28),
                Visible = false
            };

            _accountPanel = new Panel
            {
                BackColor = Color.Transparent,
                Size = new Size(312, 74)
            };
            _accountPanel.Paint += (_, e) => PaintAccountCapsule(e.Graphics, _accountPanel.ClientRectangle);
            ApplyRoundedRegion(_accountPanel, 30);

            _lblAccountAvatar = new Label
            {
                Font = WorkbenchUi.CreateUiFont(13F, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = WorkbenchUi.PrimaryColor,
                ImageAlign = ContentAlignment.MiddleCenter,
                TextAlign = ContentAlignment.MiddleCenter,
                UseMnemonic = false
            };
            _lblAccountAvatar.Paint += (_, e) => PaintGuestAccountIcon(e.Graphics, _lblAccountAvatar.ClientRectangle);
            ApplyRoundedRegion(_lblAccountAvatar, 22);

            _lblAccountName = new Label
            {
                Font = WorkbenchUi.CreateUiFont(9.6F, FontStyle.Bold),
                ForeColor = WorkbenchUi.PrimaryTextColor,
                BackColor = Color.Transparent,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true
            };

            _lblAccountState = new Label
            {
                Font = WorkbenchUi.CreateUiFont(8.4F, FontStyle.Regular),
                ForeColor = WorkbenchUi.TertiaryTextColor,
                BackColor = Color.Transparent,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true,
                Visible = false
            };

            _accountDivider = new Panel
            {
                BackColor = Color.FromArgb(205, 220, 238),
                Size = new Size(1, 18),
                Visible = false
            };

            _btnSignIn = WorkbenchUi.CreatePrimaryButton("登录", Point.Empty, new Size(80, 38), 19);
            _btnSignIn.Font = WorkbenchUi.CreateUiFont(9F, FontStyle.Bold);
            _btnSignIn.TabStop = false;
            _btnSignIn.Click += (_, _) => HandleSignIn();

            _btnSignOut = WorkbenchUi.CreateTextButton("注销", Point.Empty, new Size(54, 30), 15);
            _btnSignOut.Font = WorkbenchUi.CreateUiFont(8.9F, FontStyle.Bold);
            _btnSignOut.TabStop = false;
            _btnSignOut.Click += (_, _) => HandleSignOut();

            _accountPanel.Controls.AddRange(new Control[] { _lblAccountAvatar, _lblAccountName, _lblAccountState, _accountDivider, _btnSignIn, _btnSignOut });

            _progressBar = new ProgressBar
            {
                Style = ProgressBarStyle.Continuous,
                Location = new Point(0, 0),
                Size = new Size(520, 12),
                Value = 0,
                Visible = false
            };

            Panel sidebarDivider = BuildDivider(DockStyle.Right, 1);
            Panel headerDivider = BuildDivider(DockStyle.Bottom, 1);
            Panel statusDivider = BuildDivider(DockStyle.Top, 1);
            sidebarDivider.BackColor = Color.Transparent;
            headerDivider.BackColor = Color.Transparent;

            _sidebar.Controls.Add(sidebarDivider);
            _sidebar.Controls.Add(_sidebarNavSurface);
            _sidebar.Controls.Add(_navSelectionBar);
            _sidebar.Controls.Add(_btnToggleSidebar);
            _brandPanel.Controls.Add(_lblBrandGlyph);
            _brandPanel.Controls.Add(_lblAppTitle);
            _brandPanel.Controls.Add(_lblAppSubtitle);
            _sidebar.Controls.Add(_brandPanel);

            _headerPanel.Controls.Add(headerDivider);
            _headerPanel.Controls.Add(_lblPageTitle);
            _headerPanel.Controls.Add(_lblPageDescription);
            _headerPanel.Controls.Add(_accountPanel);

            _statusBar.Controls.Add(statusDivider);
            _statusBar.Controls.Add(_progressBar);

            Controls.Add(_contentHost);
            Controls.Add(_statusBar);
            Controls.Add(_headerPanel);
            Controls.Add(_sidebar);

            BuildSidebar();
            ApplyLanguage();
            Resize += (_, _) => ApplyWorkbenchLayout();
            DpiChanged += (_, _) => this.BeginInvokeWhenReady(() => ApplyWorkbenchLayout());
            Shown += (_, _) => OnWorkbenchShown();
            ApplyWorkbenchLayout();
            ShowPage("home");
        }

        private static Size GetInitialClientSize()
        {
            Rectangle workingArea = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1600, 1000);
            int maxWidth = Math.Max(1024, Math.Min(2200, workingArea.Width - 80));
            int maxHeight = Math.Max(720, Math.Min(1400, workingArea.Height - 80));
            int width = Math.Clamp((int)(workingArea.Width * 0.82), 1024, maxWidth);
            int height = Math.Clamp((int)(workingArea.Height * 0.82), 720, maxHeight);
            return new Size(width, height);
        }

        private void OnWorkbenchShown()
        {
            ApplyAccessControl();
            UpdateHeaderAccountPresentation();
        }

        private static void PaintSidebarBackground(Graphics graphics, Rectangle bounds)
        {
            graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using var brush = new LinearGradientBrush(
                bounds,
                Color.White,
                Color.FromArgb(250, 253, 255),
                LinearGradientMode.ForwardDiagonal);
            graphics.FillRectangle(brush, bounds);

            using var glow = new SolidBrush(Color.FromArgb(18, 205, 232, 255));
            graphics.FillEllipse(glow, new Rectangle(-170, 88, 310, 370));
        }

        private static void PaintMainCanvas(Graphics graphics, Rectangle bounds, bool roundedTopLeft)
        {
            if (bounds.Width <= 0 || bounds.Height <= 0)
            {
                return;
            }

            graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using var brush = new LinearGradientBrush(
                bounds,
                Color.FromArgb(250, 253, 255),
                Color.FromArgb(229, 244, 255),
                LinearGradientMode.ForwardDiagonal);

            if (!roundedTopLeft)
            {
                graphics.FillRectangle(brush, bounds);
                return;
            }

            using var path = BuildRoundedTopLeftPath(new Rectangle(0, 32, bounds.Width, bounds.Height + 48), 34);
            graphics.FillPath(brush, path);
        }

        private static GraphicsPath BuildRoundedTopLeftPath(Rectangle rect, int radius)
        {
            int d = radius * 2;
            var path = new GraphicsPath();
            path.StartFigure();
            path.AddArc(new Rectangle(rect.Left, rect.Top, d, d), 180, 90);
            path.AddLine(rect.Left + radius, rect.Top, rect.Right, rect.Top);
            path.AddLine(rect.Right, rect.Top, rect.Right, rect.Bottom);
            path.AddLine(rect.Right, rect.Bottom, rect.Left, rect.Bottom);
            path.AddLine(rect.Left, rect.Bottom, rect.Left, rect.Top + radius);
            path.CloseFigure();
            return path;
        }

        private static void PaintBrandLogo(Graphics graphics, Rectangle bounds)
        {
            if (bounds.Width <= 1 || bounds.Height <= 1)
            {
                return;
            }

            graphics.SmoothingMode = SmoothingMode.AntiAlias;
            graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

            using var logoFont = new Font("Segoe UI", 20F, FontStyle.Bold, GraphicsUnit.Pixel);
            using var titleFont = new Font("Microsoft YaHei UI", 17F, FontStyle.Bold, GraphicsUnit.Pixel);
            const TextFormatFlags flags = TextFormatFlags.NoPadding | TextFormatFlags.NoClipping | TextFormatFlags.SingleLine;
            var blue = Color.FromArgb(24, 119, 242);
            var green = Color.FromArgb(28, 171, 96);

            Size logoSize = TextRenderer.MeasureText(graphics, "21V", logoFont, Size.Empty, flags);
            Size titleSize = TextRenderer.MeasureText(graphics, "Office工具箱", titleFont, Size.Empty, flags);
            int logoTop = (bounds.Height - logoSize.Height) / 2;
            int titleTop = (bounds.Height - titleSize.Height) / 2;

            TextRenderer.DrawText(graphics, "21", logoFont, new Point(0, logoTop), blue, flags);
            int vLeft = TextRenderer.MeasureText(graphics, "21", logoFont, Size.Empty, flags).Width - 2;
            TextRenderer.DrawText(graphics, "V", logoFont, new Point(vLeft, logoTop), green, flags);
            TextRenderer.DrawText(graphics, "Office工具箱", titleFont, new Point(58, titleTop), WorkbenchUi.PrimaryTextColor, flags);
        }

        private static void PaintAccountCapsule(Graphics graphics, Rectangle bounds)
        {
            if (bounds.Width <= 2 || bounds.Height <= 2)
            {
                return;
            }

            graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using var ambientShadowPath = BuildRoundedPath(new Rectangle(5, 8, bounds.Width - 10, bounds.Height - 12), 28);
            using var ambientShadow = new SolidBrush(Color.FromArgb(14, 88, 130, 185));
            graphics.FillPath(ambientShadow, ambientShadowPath);

            using var shadowPath = BuildRoundedPath(new Rectangle(2, 4, bounds.Width - 5, bounds.Height - 7), 28);
            using var shadowBrush = new SolidBrush(Color.FromArgb(12, 54, 104, 166));
            graphics.FillPath(shadowBrush, shadowPath);

            using var path = BuildRoundedPath(new Rectangle(0, 0, bounds.Width - 3, bounds.Height - 5), 28);
            using var fill = new LinearGradientBrush(
                new Rectangle(0, 0, bounds.Width, bounds.Height),
                Color.FromArgb(255, 255, 255),
                Color.FromArgb(248, 251, 255),
                LinearGradientMode.Vertical);
            using var border = new Pen(Color.FromArgb(220, 235, 250));
            graphics.FillPath(fill, path);
            graphics.DrawPath(border, path);
        }

        private void PaintGuestAccountIcon(Graphics graphics, Rectangle bounds)
        {
            if (HasAccessGranted() || bounds.Width <= 2 || bounds.Height <= 2)
            {
                return;
            }

            graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using var pen = new Pen(Color.FromArgb(90, 135, 210), 1.8F)
            {
                StartCap = LineCap.Round,
                EndCap = LineCap.Round,
                LineJoin = LineJoin.Round
            };

            RectangleF head = new(bounds.Left + bounds.Width * 0.36F, bounds.Top + bounds.Height * 0.24F, bounds.Width * 0.28F, bounds.Height * 0.28F);
            RectangleF shoulders = new(bounds.Left + bounds.Width * 0.24F, bounds.Top + bounds.Height * 0.54F, bounds.Width * 0.52F, bounds.Height * 0.28F);
            graphics.DrawEllipse(pen, head);
            graphics.DrawArc(pen, shoulders, 205, 130);
        }

        private static Panel BuildDivider(DockStyle dock, int thickness)
        {
            return new Panel
            {
                Dock = dock,
                Width = dock is DockStyle.Left or DockStyle.Right ? thickness : 0,
                Height = dock is DockStyle.Top or DockStyle.Bottom ? thickness : 0,
                BackColor = WorkbenchUi.CardBorderColor
            };
        }

        private void BuildSidebar()
        {
            const int firstTop = 254;
            const int itemStep = 82;

            AddNavButton("home", GetNavText("home"), firstTop + itemStep * 0);
            AddNavButton("install", GetNavText("install"), firstTop + itemStep * 1);
            AddNavButton("uninstall", GetNavText("uninstall"), firstTop + itemStep * 2);
            AddNavButton("channel", GetNavText("channel"), firstTop + itemStep * 3);
            AddNavButton("repair", GetNavText("repair"), firstTop + itemStep * 4);
            AddNavButton("teams", GetNavText("teams"), firstTop + itemStep * 5);
            AddNavButton("outlook", GetNavText("outlook"), firstTop + itemStep * 6);
            AddNavButton("settings", GetNavText("settings"), firstTop + itemStep * 7);
        }

        private void AddNavButton(string key, string text, int top)
        {
            var button = new Button
            {
                Text = text,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = WorkbenchUi.CreateUiFont(10.6F, FontStyle.Bold),
                ForeColor = WorkbenchUi.PrimaryTextColor,
                BackColor = Color.Transparent,
                FlatStyle = FlatStyle.Flat,
                Location = new Point(30, top),
                Size = new Size(267, 86),
                Tag = key,
                Cursor = Cursors.Hand,
                Padding = new Padding(54, 0, 0, 0),
                TabStop = false,
                UseMnemonic = false
            };

            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = Color.FromArgb(246, 251, 255);
            button.FlatAppearance.MouseDownBackColor = Color.FromArgb(241, 247, 255);
            button.Click += (_, _) => ShowPage(key);
            ApplyRoundedRegion(button, 18);

            _navButtons[key] = button;
            _navButtonTexts[key] = text;
            _navButtonIcons[key] = GetNavIcon(key);
            _navToolTip.SetToolTip(button, text);
            _sidebar.Controls.Add(button);
            button.BringToFront();
        }

        private string GetNavText(string key)
        {
            if (_currentLanguage == UiLanguage.English)
            {
                return key switch
                {
                    "home" => "Home",
                    "install" => "Install",
                    "uninstall" => "Uninstall",
                    "channel" => "Update Channel",
                    "repair" => "Cleanup & Repair",
                    "teams" => "Teams Tools",
                    "outlook" => "Outlook Tools",
                    "settings" => "About",
                    _ => "Page"
                };
            }

            return key switch
            {
                "home" => "首页",
                "install" => "安装",
                "uninstall" => "卸载",
                "channel" => "更新频道",
                "repair" => "清理与修复",
                "teams" => "Teams 工具",
                "outlook" => "Outlook 工具",
                "settings" => "关于",
                _ => "页面"
            };
        }

        private void ApplyLanguage()
        {
            Text = "21Vianet Office工具箱";
            _lblBrandGlyph.Text = "21V";
            _lblAppTitle.Text = "Office工具箱";
            _lblAppSubtitle.Text = T("21Vianet Office 运维工具", "21Vianet Office Toolbox");
            _navToolTip.SetToolTip(_btnToggleSidebar, T("展开或折叠导航", "Expand or collapse navigation"));
            _btnSignIn.Text = T("登录", "Sign in");
            _btnSignOut.Text = T("注销", "Sign out");
            _navToolTip.SetToolTip(_btnSignIn, T("登录并验证 21V 账户", "Sign in and verify 21V account"));
            _navToolTip.SetToolTip(_btnSignOut, T("注销并重置当前验证状态", "Sign out and reset current verification state"));
            UpdateHeaderAccountPresentation();

            foreach ((string key, Button button) in _navButtons)
            {
                string text = GetNavText(key);
                _navButtonTexts[key] = text;
                if (!_isSidebarCollapsed)
                {
                    button.Text = text;
                }
                _navToolTip.SetToolTip(button, text);
            }

            foreach (Control page in _pageCache.Values)
            {
                ApplyLanguageToPage(page);
            }

            UpdatePageHeaderText(_currentPageKey);
            ApplySidebarLayout();
            UpdateHeaderLayout();
            ApplyAccessControl();
        }

        private bool HasAccessGranted()
        {
            return _sessionVerified && _appSettings.HasValidated21VAccount;
        }

        private bool IsRestrictedPage(string key)
        {
            return !string.Equals(key, "home", StringComparison.OrdinalIgnoreCase);
        }

        private void ApplyAccessControl()
        {
            bool granted = HasAccessGranted();
            foreach ((string key, Button button) in _navButtons)
            {
                bool restricted = IsRestrictedPage(key);
                button.Enabled = true;
            }

            if (IsRestrictedPage(_currentPageKey) && !granted)
            {
                ShowPage("home");
                return;
            }

            UpdateNavigationVisualState();
            UpdateHeaderAccountPresentation();
        }

        private void OnLanguageModeChanged(string languageMode)
        {
            _appSettings.LanguageMode = languageMode;
            _currentLanguage = LocalizationService.ResolveLanguage(languageMode);
            ApplyLanguage();
        }

        private void OnAccessStateChanged()
        {
            RefreshSessionVerificationFromSettings();
            _currentLanguage = LocalizationService.ResolveLanguage(_appSettings.LanguageMode);
            ApplyLanguage();
            if (IsRestrictedPage(_currentPageKey) && !HasAccessGranted())
            {
                ShowPage("home");
            }
        }

        private bool PromptPortalVerification()
        {
            var dialog = new MyAccountLoginVerificationForm(_currentLanguage);
            DialogResult result = dialog.ShowDialog(this);
            bool verified = result == DialogResult.OK &&
                            dialog.IsVerified;
            if (verified)
            {
                string verifiedAccount = dialog.VerifiedAccount.Trim();
                string previousAccount = (_appSettings.LastValidatedAccount ?? string.Empty).Trim();
                string previousDisplayName = (_appSettings.LastValidatedDisplayName ?? string.Empty).Trim();
                string previousAvatarPath = (_appSettings.LastValidatedAvatarPath ?? string.Empty).Trim();
                bool sameAccount = string.Equals(previousAccount, verifiedAccount, StringComparison.OrdinalIgnoreCase);
                string fallbackUserName = ExtractUserNameFromAccount(verifiedAccount);
                string dialogUserName = (dialog.VerifiedUserName ?? string.Empty).Trim();
                bool hasReliableDialogName = dialog.HasVerifiedProfileDisplayName &&
                                             !IsInvalidCachedDisplayName(dialogUserName) &&
                                             !IsFallbackAccountDisplayName(dialogUserName, verifiedAccount);
                bool hasReliablePreviousName = sameAccount &&
                                               !IsInvalidCachedDisplayName(previousDisplayName) &&
                                               !IsFallbackAccountDisplayName(previousDisplayName, verifiedAccount);
                string verifiedUserName = hasReliableDialogName
                    ? dialogUserName
                    : hasReliablePreviousName
                        ? previousDisplayName
                        : !string.IsNullOrWhiteSpace(dialogUserName)
                            ? dialogUserName
                            : fallbackUserName;
                _sessionVerified = true;
                _appSettings.HasValidated21VAccount = true;
                _appSettings.LastValidatedAccount = verifiedAccount;
                _appSettings.LastValidatedDisplayName = verifiedUserName;
                _appSettings.LastValidatedAvatarUrl = string.Empty;
                string avatarPath = SaveValidatedAvatar(dialog.VerifiedAvatarBytes, verifiedAccount, verifiedUserName);
                if (!string.IsNullOrWhiteSpace(avatarPath))
                {
                    _appSettings.LastValidatedAvatarPath = avatarPath;
                }
                else if (sameAccount && IsUsableAvatarPath(previousAvatarPath))
                {
                    _appSettings.LastValidatedAvatarPath = previousAvatarPath;
                }
                else if (!string.Equals(previousAccount, verifiedAccount, StringComparison.OrdinalIgnoreCase))
                {
                    _appSettings.LastValidatedAvatarPath = string.Empty;
                }

                _appSettings.LastValidatedAtUtc = DateTime.UtcNow;
                _appSettingsService.Save(_appSettings);
                string avatarSaveReason = !string.IsNullOrWhiteSpace(avatarPath)
                    ? "saved"
                    : sameAccount && IsUsableAvatarPath(previousAvatarPath)
                        ? "kept-existing"
                        : "no-avatar-bytes";
                WriteAvatarSaveTrace(verifiedAccount, verifiedUserName, dialog.VerifiedAvatarBytes?.Length ?? 0, _appSettings.LastValidatedAvatarPath ?? string.Empty, avatarSaveReason);
                ResetAccountAvatarImage();
                RefreshSessionVerificationFromSettings();
                ApplyLanguage();
                ApplyAccessControl();
                ShowPage("home");
                if (!TryBeginDeferredProfileSync(dialog, verifiedAccount))
                {
                    dialog.Dispose();
                }

                return true;
            }

            dialog.Dispose();
            RefreshSessionVerificationFromSettings();
            ApplyLanguage();
            ApplyAccessControl();
            return false;
        }

        private bool TryBeginDeferredProfileSync(MyAccountLoginVerificationForm dialog, string account)
        {
            if (dialog.VerifiedAvatarBytes is not null ||
                string.IsNullOrWhiteSpace(account))
            {
                return false;
            }

            return dialog.BeginDeferredProfileSync(this, ApplyDeferredProfileSyncResult);
        }

        private void ApplyDeferredProfileSyncResult(
            string account,
            string displayName,
            byte[]? avatarBytes,
            bool hasVerifiedProfileDisplayName)
        {
            if (IsDisposed || Disposing)
            {
                return;
            }

            if (InvokeRequired)
            {
                this.BeginInvokeWhenReady(() => ApplyDeferredProfileSyncResult(
                    account,
                    displayName,
                    avatarBytes,
                    hasVerifiedProfileDisplayName));
                return;
            }

            string currentAccount = (_appSettings.LastValidatedAccount ?? string.Empty).Trim();
            if (!HasAccessGranted() ||
                !string.Equals(currentAccount, account.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                WriteAvatarSaveTrace(account, displayName, avatarBytes?.Length ?? 0, string.Empty, "deferred-ignored-stale-account");
                return;
            }

            bool updated = false;
            string currentDisplayName = (_appSettings.LastValidatedDisplayName ?? string.Empty).Trim();
            string candidateDisplayName = (displayName ?? string.Empty).Trim();
            bool hasReliableDisplayName = hasVerifiedProfileDisplayName &&
                                          !IsInvalidCachedDisplayName(candidateDisplayName) &&
                                          !IsFallbackAccountDisplayName(candidateDisplayName, account);
            if (hasReliableDisplayName &&
                !string.Equals(candidateDisplayName, currentDisplayName, StringComparison.Ordinal))
            {
                _appSettings.LastValidatedDisplayName = candidateDisplayName;
                currentDisplayName = candidateDisplayName;
                updated = true;
            }

            string avatarPath = SaveValidatedAvatar(avatarBytes, account, currentDisplayName);
            if (!string.IsNullOrWhiteSpace(avatarPath))
            {
                string previousAvatarPath = (_appSettings.LastValidatedAvatarPath ?? string.Empty).Trim();
                if (!string.Equals(previousAvatarPath, avatarPath, StringComparison.OrdinalIgnoreCase))
                {
                    _appSettings.LastValidatedAvatarPath = avatarPath;
                    DeleteCachedAvatar(previousAvatarPath);
                    updated = true;
                }
                else if (!IsUsableAvatarPath(previousAvatarPath))
                {
                    _appSettings.LastValidatedAvatarPath = avatarPath;
                    updated = true;
                }
            }

            string reason = !string.IsNullOrWhiteSpace(avatarPath)
                ? "deferred-saved"
                : updated
                    ? "deferred-updated-profile"
                    : "deferred-no-avatar";
            WriteAvatarSaveTrace(account, currentDisplayName, avatarBytes?.Length ?? 0, _appSettings.LastValidatedAvatarPath ?? string.Empty, reason);

            if (!updated)
            {
                return;
            }

            _appSettingsService.Save(_appSettings);
            ResetAccountAvatarImage();
            UpdateHeaderAccountPresentation();
        }

        private void ResetVerificationState()
        {
            _appSettings.HasValidated21VAccount = false;
            _appSettings.LastValidatedAccount = string.Empty;
            _appSettings.LastValidatedDisplayName = string.Empty;
            _appSettings.LastValidatedAvatarUrl = string.Empty;
            DeleteCachedAvatar(_appSettings.LastValidatedAvatarPath);
            _appSettings.LastValidatedAvatarPath = string.Empty;
            _appSettings.LastValidatedAtUtc = null;
            _appSettingsService.Save(_appSettings);
            ResetAccountAvatarImage();
            RefreshSessionVerificationFromSettings();
        }

        private void ResetStaleValidatedProfile()
        {
            if (!_appSettings.HasValidated21VAccount)
            {
                return;
            }

            string account = (_appSettings.LastValidatedAccount ?? string.Empty).Trim();
            string displayName = (_appSettings.LastValidatedDisplayName ?? string.Empty).Trim();
            bool invalidDisplayName = IsInvalidCachedDisplayName(displayName);
            bool placeholderAvatar = IsPlaceholderAvatarCache(_appSettings.LastValidatedAvatarPath);
            if (!invalidDisplayName)
            {
                if (placeholderAvatar)
                {
                    DeleteCachedAvatar(_appSettings.LastValidatedAvatarPath);
                    _appSettings.LastValidatedAvatarPath = string.Empty;
                    _appSettingsService.Save(_appSettings);
                }

                return;
            }

            _appSettings.HasValidated21VAccount = false;
            _appSettings.LastValidatedAccount = string.Empty;
            _appSettings.LastValidatedDisplayName = string.Empty;
            _appSettings.LastValidatedAvatarUrl = string.Empty;
            DeleteCachedAvatar(_appSettings.LastValidatedAvatarPath);
            _appSettings.LastValidatedAvatarPath = string.Empty;
            _appSettings.LastValidatedAtUtc = null;
            _appSettingsService.Save(_appSettings);
        }

        private static bool IsPlaceholderAvatarCache(string path)
        {
            try
            {
                return string.IsNullOrWhiteSpace(path) ||
                       !File.Exists(path);
            }
            catch
            {
                return true;
            }
        }

        private static bool IsInvalidCachedDisplayName(string displayName)
        {
            return string.IsNullOrWhiteSpace(displayName) ||
                   System.Text.RegularExpressions.Regex.IsMatch(
                       displayName.Trim(),
                       @"^[\{\}\[\]\(\),.:;\-_=+]+$|^(管理|创建|应用|搜索|文件|最近|已共享|收藏夹|类型|取消固定|更多|设置及其他)$",
                       System.Text.RegularExpressions.RegexOptions.CultureInvariant);
        }

        private static bool IsFallbackAccountDisplayName(string displayName, string account)
        {
            if (string.IsNullOrWhiteSpace(displayName) || string.IsNullOrWhiteSpace(account))
            {
                return false;
            }

            return string.Equals(displayName.Trim(), ExtractUserNameFromAccount(account), StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsUsableAvatarPath(string path)
        {
            try
            {
                return !string.IsNullOrWhiteSpace(path) && File.Exists(path);
            }
            catch
            {
                return false;
            }
        }

        private void RefreshSessionVerificationFromSettings()
        {
            _sessionVerified = _appSettings.HasValidated21VAccount &&
                               (!string.IsNullOrWhiteSpace(_appSettings.LastValidatedAccount) ||
                                !string.IsNullOrWhiteSpace(_appSettings.LastValidatedDisplayName));
        }

        private static string ExtractUserNameFromAccount(string account)
        {
            if (string.IsNullOrWhiteSpace(account))
            {
                return string.Empty;
            }

            string trimmed = account.Trim();
            int at = trimmed.IndexOf('@');
            return at > 0 ? trimmed[..at] : trimmed;
        }

        private string T(string zh, string en) => LocalizationService.Localize(_currentLanguage, zh, en);

        private void ToggleSidebar()
        {
            if (ShouldForceCollapsedSidebar())
            {
                _isSidebarCollapsed = true;
                ApplyWorkbenchLayout(updateSidebarMode: false);
                return;
            }

            _isSidebarCollapsed = !_isSidebarCollapsed;
            ApplyWorkbenchLayout(updateSidebarMode: false);
        }

        private float GetShellScale() =>
            WorkbenchUi.GetAdaptiveUiScale(ClientSize.Width, ClientSize.Height);

        private bool ShouldForceCollapsedSidebar() =>
            ClientSize.Width < 1280;

        private static int GetSidebarWidth(bool isCollapsed, float scale) =>
            WorkbenchUi.Scale(isCollapsed ? CollapsedSidebarWidth : ExpandedSidebarWidth, scale);

        private int GetHeaderBaseHeight() =>
            string.Equals(_currentPageKey, "home", StringComparison.OrdinalIgnoreCase) ? 170 : 194;

        private void ApplyWorkbenchLayout(bool updateSidebarMode = true)
        {
            if (updateSidebarMode)
            {
                if (ShouldForceCollapsedSidebar())
                {
                    _isSidebarCollapsed = true;
                }
                else if (ClientSize.Width > 1380)
                {
                    _isSidebarCollapsed = false;
                }
            }

            float shellScale = GetShellScale();
            int sidebarWidth = GetSidebarWidth(_isSidebarCollapsed, shellScale);
            int headerBaseHeight = GetHeaderBaseHeight();
            float scale = WorkbenchUi.GetAdaptiveUiScale(
                Math.Max(1, ClientSize.Width - sidebarWidth),
                Math.Max(1, ClientSize.Height - WorkbenchUi.Scale(headerBaseHeight, shellScale)));
            _headerPanel.Height = WorkbenchUi.Scale(headerBaseHeight, scale);
            ApplySidebarLayout();
            UpdateStatusLayout();
            UpdateHeaderLayout();
        }

        private void ApplySidebarLayout()
        {
            float scale = GetShellScale();
            int S(int value) => WorkbenchUi.Scale(value, scale);

            int expandedSidebarWidth = S(ExpandedSidebarWidth);
            int collapsedSidebarWidth = S(CollapsedSidebarWidth);
            int sidebarBottomMargin = S(SidebarBottomMargin);
            int sidebarNavGap = S(SidebarNavGap);
            int collapsedNavButtonHeight = S(CollapsedNavButtonHeight);

            _sidebar.Width = _isSidebarCollapsed ? collapsedSidebarWidth : expandedSidebarWidth;

            _brandPanel.Visible = false;
            _lblBrandGlyph.Visible = false;
            _lblAppTitle.Visible = false;
            _lblAppSubtitle.Visible = false;

            _btnToggleSidebar.Size = WorkbenchUi.ScaleSize(new Size(48, 48), scale);
            WorkbenchUi.ApplyIconFont(_btnToggleSidebar, 16.5F, FontStyle.Regular, scale);
            ApplyRoundedRegion(_btnToggleSidebar, S(24));
            int toggleLeft = _isSidebarCollapsed
                ? (collapsedSidebarWidth - _btnToggleSidebar.Width) / 2
                : S(18);
            _btnToggleSidebar.Location = new Point(toggleLeft, S(26));

            string[] navOrder = { "home", "install", "uninstall", "channel", "repair", "teams", "outlook", "settings" };
            int sidebarHeight = _sidebar.ClientSize.Height;
            bool compactNav = sidebarHeight < S(980);
            int expandedNavHeight = S(compactNav ? 64 : 86);
            int expandedFirstTop = S(compactNav ? 154 : 184);
            int expandedItemStep = S(compactNav ? 68 : 82);
            if (compactNav)
            {
                int bottomLimit = Math.Max(
                    _btnToggleSidebar.Bottom + sidebarNavGap + expandedNavHeight,
                    sidebarHeight - sidebarBottomMargin);
                int minimumGap = sidebarHeight < S(760) ? S(4) : sidebarNavGap;
                int minimumFirstTop = _btnToggleSidebar.Bottom + (sidebarHeight < S(760) ? S(24) : S(48));
                int requiredHeight = navOrder.Length * expandedNavHeight + (navOrder.Length - 1) * minimumGap;
                if (bottomLimit - expandedFirstTop < requiredHeight)
                {
                    expandedFirstTop = Math.Max(minimumFirstTop, bottomLimit - requiredHeight);
                }

                int availableHeight = Math.Max(expandedNavHeight, bottomLimit - expandedFirstTop);
                int fittedStep = navOrder.Length > 1
                    ? (availableHeight - expandedNavHeight) / (navOrder.Length - 1)
                    : expandedItemStep;
                if (fittedStep < expandedNavHeight + minimumGap)
                {
                    int fittedHeight = (availableHeight - (navOrder.Length - 1) * minimumGap) / navOrder.Length;
                    expandedNavHeight = Math.Max(S(48), Math.Min(expandedNavHeight, fittedHeight));
                    fittedStep = navOrder.Length > 1
                        ? (availableHeight - expandedNavHeight) / (navOrder.Length - 1)
                        : expandedItemStep;
                }

                expandedItemStep = Math.Min(expandedItemStep, Math.Max(expandedNavHeight, fittedStep));
            }

            int collapsedFirstTop = S(154);
            int collapsedSettingsTop = Math.Max(
                _btnToggleSidebar.Bottom + sidebarNavGap,
                sidebarHeight - sidebarBottomMargin - collapsedNavButtonHeight);
            int collapsedPrimaryCount = navOrder.Length - 1;
            int collapsedAvailableStep = collapsedPrimaryCount > 1
                ? (collapsedSettingsTop - sidebarNavGap - collapsedFirstTop - collapsedNavButtonHeight) / (collapsedPrimaryCount - 1)
                : collapsedNavButtonHeight + sidebarNavGap;
            int collapsedItemStep = Math.Max(collapsedNavButtonHeight, Math.Min(S(68), collapsedAvailableStep));

            for (int i = 0; i < navOrder.Length; i++)
            {
                string key = navOrder[i];
                if (!_navButtons.TryGetValue(key, out Button? button))
                {
                    continue;
                }

                int top = _isSidebarCollapsed
                    ? i == navOrder.Length - 1
                        ? collapsedSettingsTop
                        : collapsedFirstTop + i * collapsedItemStep
                    : expandedFirstTop + i * expandedItemStep;
                button.Location = new Point(_isSidebarCollapsed ? S(12) : S(20), top);
                button.Size = new Size(
                    _isSidebarCollapsed ? S(54) : _sidebar.Width - S(40),
                    _isSidebarCollapsed ? collapsedNavButtonHeight : expandedNavHeight);

                if (_isSidebarCollapsed)
                {
                    button.TextAlign = ContentAlignment.MiddleCenter;
                    button.Padding = Padding.Empty;
                    ApplyCollapsedNavIcon(button, key, isEnabled: true, scale);
                }
                else
                {
                    button.TextAlign = ContentAlignment.MiddleLeft;
                    button.Image = null;
                    button.Padding = new Padding(S(40), 0, 0, 0);
                    WorkbenchUi.ApplyUiFont(button, 10.6F, FontStyle.Bold, scale);
                    button.Text = _navButtonTexts[key];
                }
            }

            PositionSidebarBottomButtons();
            UpdateSidebarChrome();
            UpdateNavigationVisualState();
        }

        private void PositionSidebarBottomButtons()
        {
            if (!_navButtons.TryGetValue("settings", out Button? settingsButton))
            {
                return;
            }

            float scale = GetShellScale();
            int sidebarBottomMargin = WorkbenchUi.Scale(SidebarBottomMargin, scale);
            int sidebarNavGap = WorkbenchUi.Scale(SidebarNavGap, scale);
            int targetTop = _sidebar.ClientSize.Height - sidebarBottomMargin - settingsButton.Height;
            int maxVisibleTop = Math.Max(0, targetTop);
            if (_navButtons.TryGetValue("outlook", out Button? outlookButton))
            {
                int belowOutlook = outlookButton.Bottom + sidebarNavGap;
                if (targetTop >= belowOutlook)
                {
                    settingsButton.Top = maxVisibleTop;
                }
                else
                {
                    settingsButton.Top = Math.Min(settingsButton.Top, maxVisibleTop);
                }

                return;
            }

            settingsButton.Top = maxVisibleTop;
        }

        private void UpdateSelectionBarPosition()
        {
            _navSelectionBar.Visible = false;
        }

        private static string GetNavIcon(string key)
        {
            return key switch
            {
                "home" => "\uE80F",
                "install" => "\uE7C3",
                "uninstall" => "\uE74D",
                "channel" => "\uE8D7",
                "repair" => "\uE90F",
                "teams" => "\uE720",
                "outlook" => "\uE715",
                "settings" => "\uE946",
                _ => "\uE10C"
            };
        }

        private void ShowPage(string key)
        {
            if (IsRestrictedPage(key) && !HasAccessGranted())
            {
                key = "home";
                MessageBox.Show(
                    T("请先点击右上角“登录”完成 21V 账户验证，再使用其他功能。", "Click Sign in at the top right to complete 21V account verification before using other pages."),
                    T("访问受限", "Access Restricted"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }

            _currentPageKey = key;
            UpdatePageHeaderText(key);
            ApplyWorkbenchLayout();
            UpdateNavigationVisualState();

            _contentHost.SuspendLayout();
            _contentHost.Controls.Clear();
            Control page = GetOrCreatePage(key);
            page.Dock = DockStyle.Fill;
            _contentHost.Controls.Add(page);
            _contentHost.ResumeLayout();
            ActiveControl = null;

            ApplyLanguageToPage(page);

            if (page is WorkbenchInstallationControl installPage)
            {
                installPage.PrepareForDisplay();
                this.BeginInvokeWhenReady(installPage.PrepareForDisplay);
            }
            else if (page is WorkbenchHomeControl homePage)
            {
                homePage.PrepareForDisplay();
            }
            else if (page is WorkbenchTeamsControl teamsPage)
            {
                teamsPage.PrepareForDisplay();
            }
            else if (page is WorkbenchOutlookControl outlookPage)
            {
                outlookPage.PrepareForDisplay();
            }
            else if (page is WorkbenchRepairControl repairPage)
            {
                repairPage.PrepareForDisplay();
            }
            else if (page is WorkbenchUninstallControl uninstallPage)
            {
                uninstallPage.PrepareForDisplay();
            }
            else if (page is WorkbenchChannelControl channelPage)
            {
                channelPage.PrepareForDisplay();
            }
            else if (page is WorkbenchSettingsControl settingsPage)
            {
                settingsPage.ApplyLanguage(_currentLanguage);
                settingsPage.PrepareForDisplay();
            }

            if (_isSidebarCollapsed)
            {
                ApplySidebarLayout();
            }

            ApplyAccessControl();
        }

        private void UpdateStatusLayout()
        {
            _progressBar.Width = Math.Max(420, _statusBar.ClientSize.Width - 180);
            _progressBar.Left = (_statusBar.ClientSize.Width - _progressBar.Width) / 2;
            _progressBar.Top = (_statusBar.ClientSize.Height - _progressBar.Height) / 2;
        }

        private void UpdatePageHeaderText(string key)
        {
            (string title, string description) = GetPageHeader(key);
            _lblPageTitle.Text = title;
            bool showDescription = string.Equals(key, "home", StringComparison.OrdinalIgnoreCase) &&
                                   !string.IsNullOrWhiteSpace(description);
            _lblPageDescription.Text = showDescription ? description : string.Empty;
            _lblPageDescription.Visible = showDescription;
        }

        private void UpdateHeaderLayout()
        {
            if (_headerPanel.ClientSize.Width <= 0)
            {
                return;
            }

            float scale = WorkbenchUi.GetAdaptiveUiScale(
                _headerPanel.ClientSize.Width,
                Math.Max(1, ClientSize.Height - GetHeaderBaseHeight()));
            int S(int value) => WorkbenchUi.Scale(value, scale);

            WorkbenchUi.ApplyUiFont(_lblPageTitle, 24F, FontStyle.Bold, scale);
            WorkbenchUi.ApplyUiFont(_lblPageDescription, 11F, FontStyle.Regular, scale);
            WorkbenchUi.ApplyUiFont(_lblAccountAvatar, 13F, FontStyle.Bold, scale);
            WorkbenchUi.ApplyUiFont(_lblAccountName, 9.6F, FontStyle.Bold, scale);
            WorkbenchUi.ApplyUiFont(_lblAccountState, 8.4F, FontStyle.Regular, scale);
            WorkbenchUi.ApplyUiFont(_btnSignIn, 9F, FontStyle.Bold, scale);
            WorkbenchUi.ApplyUiFont(_btnSignOut, 8.9F, FontStyle.Bold, scale);

            int rightMargin = S(48);
            bool isHomePage = string.Equals(_currentPageKey, "home", StringComparison.OrdinalIgnoreCase);
            int top = S(isHomePage ? 58 : 88);
            int panelHeight = S(62);
            int dividerGap = S(12);
            int dividerHeight = S(18);
            bool isSignInMode = _btnSignIn.Visible;
            int paddingLeft = S(isSignInMode ? 12 : 14);
            int paddingRight = S(isSignInMode ? 12 : 14);
            int avatarSize = S(isSignInMode ? 36 : 42);
            int avatarTextGap = S(isSignInMode ? 12 : 14);
            Button activeButton = isSignInMode ? _btnSignIn : _btnSignOut;
            Button inactiveButton = isSignInMode ? _btnSignOut : _btnSignIn;
            int buttonWidth = Math.Clamp(
                TextRenderer.MeasureText(activeButton.Text, activeButton.Font).Width + S(isSignInMode ? 26 : 18),
                S(isSignInMode ? 76 : 48),
                S(isSignInMode ? 96 : 64));
            int buttonHeight = S(isSignInMode ? 36 : 30);
            activeButton.Size = new Size(buttonWidth, buttonHeight);
            int measuredNameWidth = TextRenderer.MeasureText(_lblAccountName.Text, _lblAccountName.Font).Width + S(4);
            int textWidth = isSignInMode ? Math.Clamp(measuredNameWidth, S(52), S(86)) : Math.Clamp(measuredNameWidth, S(64), S(172));
            int signedInWidth = paddingLeft + avatarSize + avatarTextGap + textWidth + dividerGap + 1 + dividerGap + activeButton.Width + paddingRight;
            int signInWidth = paddingLeft + avatarSize + avatarTextGap + textWidth + S(14) + activeButton.Width + paddingRight;
            int accountWidth = isSignInMode ? signInWidth : signedInWidth;
            _accountPanel.Size = new Size(accountWidth, panelHeight);
            _accountPanel.Location = new Point(_headerPanel.ClientSize.Width - accountWidth - rightMargin, top);
            ApplyRoundedRegion(_accountPanel, S(29));
            _accountPanel.Invalidate();

            _lblAccountAvatar.Visible = true;
            _lblAccountAvatar.Location = new Point(paddingLeft, (panelHeight - avatarSize - S(4)) / 2);
            _lblAccountAvatar.Size = new Size(avatarSize, avatarSize);
            ApplyRoundedRegion(_lblAccountAvatar, avatarSize / 2);

            activeButton.Visible = true;
            int buttonTop = (panelHeight - activeButton.Height - S(4)) / 2;
            activeButton.Location = new Point(accountWidth - paddingRight - activeButton.Width, buttonTop);
            inactiveButton.Location = new Point(accountWidth - paddingRight - inactiveButton.Width, buttonTop);

            int textLeft = _lblAccountAvatar.Right + avatarTextGap;
            _lblAccountName.Visible = true;
            _lblAccountName.Location = new Point(textLeft, (panelHeight - S(24) - S(4)) / 2);
            _lblAccountName.Size = new Size(textWidth, S(24));
            _lblAccountState.Location = new Point(textLeft, _lblAccountName.Bottom);
            _lblAccountState.Size = new Size(textWidth, 0);
            _lblAccountState.Visible = false;
            _accountDivider.Visible = !isSignInMode;
            _accountDivider.Location = new Point(_lblAccountName.Right + dividerGap, (panelHeight - dividerHeight - S(4)) / 2);
            _accountDivider.Size = new Size(1, dividerHeight);

            int pageMargin = WorkbenchUi.GetAdaptivePageMargin(_headerPanel.ClientSize.Width);
            int pageContentWidth = WorkbenchUi.GetAdaptiveContentWidth(
                _headerPanel.ClientSize.Width,
                320,
                GetCurrentPageContentMaxWidth(),
                pageMargin,
                scrollbarAllowance: SystemInformation.VerticalScrollBarWidth);
            int pageLeft = WorkbenchUi.GetAdaptiveContentLeft(_headerPanel.ClientSize.Width, pageContentWidth, pageMargin);
            int pageWidth = Math.Max(S(320), Math.Min(pageContentWidth, _accountPanel.Left - pageLeft - S(48)));
            _lblPageTitle.Location = new Point(pageLeft, S(isHomePage ? 70 : 86));
            _lblPageDescription.Location = new Point(pageLeft + S(2), S(isHomePage ? 116 : 132));
            _lblPageTitle.Size = new Size(pageWidth, S(isHomePage ? 42 : 52));
            _lblPageDescription.Size = new Size(pageWidth, S(32));
        }

        private int GetCurrentPageContentMaxWidth()
        {
            return _currentPageKey switch
            {
                "home" or "install" => WorkbenchUi.WideContentMaxWidth,
                "settings" => WorkbenchUi.ReadingContentMaxWidth,
                _ => WorkbenchUi.DefaultContentMaxWidth
            };
        }

        private void UpdateHeaderAccountPresentation()
        {
            string userName = GetHeaderUserName();
            bool isVerified = HasAccessGranted();

            _lblAccountName.Text = string.IsNullOrWhiteSpace(userName)
                ? T("未登录", "Not signed in")
                : userName;
            _lblAccountName.ForeColor = isVerified
                ? WorkbenchUi.PrimaryTextColor
                : Color.FromArgb(92, 111, 136);
            _lblAccountState.Text = string.Empty;
            _lblAccountState.Visible = false;
            UpdateAccountAvatar(userName);
            _btnSignIn.Visible = !isVerified;
            _btnSignIn.Enabled = !isVerified;
            _btnSignOut.Visible = isVerified;
            _btnSignOut.Enabled = isVerified;
            UpdateHeaderLayout();
        }

        private void UpdateAccountAvatar(string userName)
        {
            bool isVerified = HasAccessGranted();
            if (!isVerified)
            {
                ResetAccountAvatarImage();
            }

            if (isVerified)
            {
                _accountAvatarImage ??= TryLoadValidatedAvatar(_appSettings.LastValidatedAvatarPath, 42);
            }

            if (_accountAvatarImage is not null)
            {
                _lblAccountAvatar.Text = string.Empty;
                _lblAccountAvatar.Image = _accountAvatarImage;
                _lblAccountAvatar.BackColor = Color.Transparent;
                return;
            }

            string initial = isVerified && !string.IsNullOrWhiteSpace(userName) ? userName.Trim()[0].ToString().ToUpperInvariant() : string.Empty;
            _lblAccountAvatar.Image = null;
            _lblAccountAvatar.Text = initial;
            _lblAccountAvatar.BackColor = isVerified ? WorkbenchUi.PrimaryColor : Color.FromArgb(242, 248, 255);
            _lblAccountAvatar.ForeColor = Color.White;
            _lblAccountAvatar.Invalidate();
        }

        private void ResetAccountAvatarImage()
        {
            if (_accountAvatarImage is not null)
            {
                _lblAccountAvatar.Image = null;
                _accountAvatarImage.Dispose();
                _accountAvatarImage = null;
            }
        }

        private static Image? TryLoadValidatedAvatar(string path, int size)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                {
                    return null;
                }

                using var source = Image.FromFile(path);
                return CreateCircularAvatar(source, size);
            }
            catch
            {
                return null;
            }
        }

        private static string SaveValidatedAvatar(byte[]? avatarBytes, string account, string displayName)
        {
            if (avatarBytes is null || avatarBytes.Length < 128)
            {
                return string.Empty;
            }

            try
            {
                using var input = new MemoryStream(avatarBytes);
                using var source = Image.FromStream(input);
                if (source.Width < 16 || source.Height < 16)
                {
                    return string.Empty;
                }

                using var normalized = new Bitmap(source);
                string identity = string.IsNullOrWhiteSpace(account) ? displayName : account;
                string directory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "M365Tool",
                    "AccountAvatar");
                Directory.CreateDirectory(directory);

                string fileName = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(identity.Trim().ToLowerInvariant()))) + ".png";
                string path = Path.Combine(directory, fileName);
                using var output = new MemoryStream();
                normalized.Save(output, System.Drawing.Imaging.ImageFormat.Png);
                File.WriteAllBytes(path, output.ToArray());
                return path;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static void WriteAvatarSaveTrace(string account, string displayName, int sourceBytes, string path, string reason)
        {
            try
            {
                string directory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "M365Tool",
                    "Logs");
                Directory.CreateDirectory(directory);
                long fileBytes = !string.IsNullOrWhiteSpace(path) && File.Exists(path)
                    ? new FileInfo(path).Length
                    : 0;
                string line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss}\tavatar-save\tAccount={account}\tDisplay={displayName}\tSourceBytes={sourceBytes}\tPath={path}\tFileBytes={fileBytes}\tReason={reason}{Environment.NewLine}";
                File.AppendAllText(Path.Combine(directory, "AuthProfileTrace.log"), line);
            }
            catch
            {
                // Profile trace is best-effort only.
            }
        }

        private static void DeleteCachedAvatar(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }

            try
            {
                string avatarDirectory = Path.GetFullPath(Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "M365Tool",
                    "AccountAvatar"));
                string target = Path.GetFullPath(path);
                if (target.StartsWith(avatarDirectory, StringComparison.OrdinalIgnoreCase) &&
                    File.Exists(target))
                {
                    File.Delete(target);
                }
            }
            catch
            {
                // Avatar cleanup is best-effort only.
            }
        }

        private static Image CreateCircularAvatar(Image source, int size)
        {
            var bitmap = new Bitmap(size, size);
            using Graphics graphics = Graphics.FromImage(bitmap);
            graphics.SmoothingMode = SmoothingMode.AntiAlias;
            graphics.Clear(Color.Transparent);

            float scale = Math.Max(size / (float)source.Width, size / (float)source.Height);
            int width = (int)Math.Ceiling(source.Width * scale);
            int height = (int)Math.Ceiling(source.Height * scale);
            int left = (size - width) / 2;
            int top = (size - height) / 2;

            using var clip = new GraphicsPath();
            clip.AddEllipse(0, 0, size - 1, size - 1);
            graphics.SetClip(clip);
            graphics.DrawImage(source, new Rectangle(left, top, width, height));
            graphics.ResetClip();
            using var border = new Pen(Color.FromArgb(230, 242, 250, 255), 1F);
            graphics.DrawEllipse(border, 0.5F, 0.5F, size - 1.5F, size - 1.5F);
            return bitmap;
        }

        private string GetHeaderUserName()
        {
            string displayName = (_appSettings.LastValidatedDisplayName ?? string.Empty).Trim();
            if (!string.IsNullOrWhiteSpace(displayName))
            {
                return displayName;
            }

            string account = (_appSettings.LastValidatedAccount ?? string.Empty).Trim();
            if (!string.IsNullOrWhiteSpace(account))
            {
                int at = account.IndexOf('@');
                return (at > 0 ? account[..at] : account).Trim();
            }

            return (_appSettings.LastValidatedDisplayName ?? string.Empty).Trim();
        }

        private void UpdateNavigationVisualState()
        {
            bool granted = HasAccessGranted();
            float scale = GetShellScale();
            foreach ((string navKey, Button navButton) in _navButtons)
            {
                bool isActive = string.Equals(navKey, _currentPageKey, StringComparison.OrdinalIgnoreCase);
                bool isRestricted = IsRestrictedPage(navKey);
                bool isEnabled = !isRestricted || granted;

                navButton.Enabled = true;
                if (!isEnabled)
                {
                    navButton.BackColor = Color.Transparent;
                    navButton.ForeColor = Color.FromArgb(176, 188, 202);
                    navButton.Cursor = Cursors.Default;
                    navButton.FlatAppearance.BorderSize = 0;
                    navButton.FlatAppearance.MouseOverBackColor = Color.Transparent;
                    navButton.FlatAppearance.MouseDownBackColor = Color.Transparent;
                    if (_isSidebarCollapsed)
                    {
                        ApplyCollapsedNavIcon(navButton, navKey, isEnabled: false, scale);
                    }
                    else
                    {
                        WorkbenchUi.ApplyUiFont(navButton, 10.6F, FontStyle.Bold, scale);
                        navButton.Text = _navButtonTexts[navKey];
                        navButton.Image = null;
                    }

                    continue;
                }

                navButton.Cursor = Cursors.Hand;
                navButton.FlatAppearance.BorderSize = 0;
                navButton.FlatAppearance.MouseOverBackColor = Color.FromArgb(252, 254, 255);
                navButton.FlatAppearance.MouseDownBackColor = Color.FromArgb(240, 246, 255);
                navButton.BackColor = isActive ? Color.FromArgb(229, 242, 255) : Color.Transparent;
                navButton.ForeColor = isActive ? WorkbenchUi.PrimaryColor : WorkbenchUi.SecondaryTextColor;
                if (_isSidebarCollapsed)
                {
                    ApplyCollapsedNavIcon(navButton, navKey, isEnabled: true, scale);
                }
                else
                {
                    WorkbenchUi.ApplyUiFont(navButton, 10.6F, FontStyle.Bold, scale);
                    navButton.Text = _navButtonTexts[navKey];
                    navButton.Image = null;
                }
            }

            UpdateSelectionBarPosition();
        }

        private static void ApplyCollapsedNavIcon(Button button, string key, bool isEnabled, float scale)
        {
            if (TryGetNavPngIcon(key, isEnabled, out Image? image))
            {
                WorkbenchUi.ApplyUiFont(button, 1F, FontStyle.Regular, scale);
                button.Text = string.Empty;
                button.Image = image;
                button.ImageAlign = ContentAlignment.MiddleCenter;
                return;
            }

            button.Image = null;
            button.ImageAlign = ContentAlignment.MiddleCenter;
            WorkbenchUi.ApplyIconFont(button, 13.5F, FontStyle.Regular, scale);
            button.Text = GetNavIcon(key);
        }

        private static bool TryGetNavPngIcon(string key, bool isEnabled, out Image? image)
        {
            IReadOnlyDictionary<string, Image> images = isEnabled
                ? NavPngIconImages.Value
                : DisabledNavPngIconImages.Value;
            return images.TryGetValue(key, out image);
        }

        private static IReadOnlyDictionary<string, Image> LoadNavPngIconImages(float alpha)
        {
            var images = new Dictionary<string, Image>();
            var assembly = typeof(WorkbenchForm).Assembly;

            foreach (KeyValuePair<string, string> resource in NavPngIconResourceNames)
            {
                using Stream stream = assembly.GetManifestResourceStream(resource.Value)
                    ?? throw new InvalidOperationException($"Missing sidebar navigation icon resource: {resource.Value}");
                using Image source = Image.FromStream(stream);
                images[resource.Key] = CreateScaledNavIconImage(source, 34, alpha);
            }

            return images;
        }

        private static Image CreateScaledNavIconImage(Image source, int size, float alpha)
        {
            var bitmap = new Bitmap(size, size);
            using Graphics graphics = Graphics.FromImage(bitmap);
            graphics.Clear(Color.Transparent);
            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphics.SmoothingMode = SmoothingMode.AntiAlias;
            graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

            var colorMatrix = new ColorMatrix
            {
                Matrix33 = alpha
            };
            using var attributes = new ImageAttributes();
            attributes.SetColorMatrix(colorMatrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap);
            graphics.DrawImage(
                source,
                new Rectangle(0, 0, size, size),
                0,
                0,
                source.Width,
                source.Height,
                GraphicsUnit.Pixel,
                attributes);

            return bitmap;
        }

        private static void SetNavButtonImage(Button button, string glyph, Color color)
        {
            Image? oldImage = button.Image;
            button.Image = CreateGlyphBitmap(glyph, color, 22);
            oldImage?.Dispose();
        }

        private static Image CreateGlyphBitmap(string glyph, Color color, int size)
        {
            var bitmap = new Bitmap(size, size);
            using Graphics graphics = Graphics.FromImage(bitmap);
            graphics.Clear(Color.Transparent);
            graphics.SmoothingMode = SmoothingMode.AntiAlias;
            DrawModernNavIcon(graphics, glyph, color, new Rectangle(2, 2, size - 4, size - 4));
            return bitmap;
        }

        private static void DrawModernNavIcon(Graphics graphics, string glyph, Color color, Rectangle rect)
        {
            using var pen = new Pen(color, 1.9F)
            {
                StartCap = LineCap.Round,
                EndCap = LineCap.Round,
                LineJoin = LineJoin.Round
            };
            using var brush = new SolidBrush(Color.FromArgb(28, color));

            RectangleF r = rect;
            switch (glyph)
            {
                case "\uE80F": // Home
                    PointF[] roof =
                    {
                        new(r.Left + r.Width * 0.18F, r.Top + r.Height * 0.48F),
                        new(r.Left + r.Width * 0.50F, r.Top + r.Height * 0.18F),
                        new(r.Left + r.Width * 0.82F, r.Top + r.Height * 0.48F)
                    };
                    graphics.DrawLines(pen, roof);
                    using (GraphicsPath body = BuildRoundedPath(Rectangle.Round(new RectangleF(r.Left + r.Width * 0.26F, r.Top + r.Height * 0.45F, r.Width * 0.48F, r.Height * 0.40F)), 4))
                    {
                        graphics.FillPath(brush, body);
                        graphics.DrawPath(pen, body);
                    }
                    break;
                case "\uE7C3": // Install package
                    PointF[] cube =
                    {
                        new(r.Left + r.Width * 0.50F, r.Top + r.Height * 0.16F),
                        new(r.Left + r.Width * 0.80F, r.Top + r.Height * 0.33F),
                        new(r.Left + r.Width * 0.80F, r.Top + r.Height * 0.68F),
                        new(r.Left + r.Width * 0.50F, r.Top + r.Height * 0.84F),
                        new(r.Left + r.Width * 0.20F, r.Top + r.Height * 0.68F),
                        new(r.Left + r.Width * 0.20F, r.Top + r.Height * 0.33F)
                    };
                    graphics.FillPolygon(brush, cube);
                    graphics.DrawPolygon(pen, cube);
                    graphics.DrawLine(pen, cube[0], new PointF(r.Left + r.Width * 0.50F, r.Top + r.Height * 0.50F));
                    graphics.DrawLine(pen, cube[1], new PointF(r.Left + r.Width * 0.50F, r.Top + r.Height * 0.50F));
                    graphics.DrawLine(pen, cube[5], new PointF(r.Left + r.Width * 0.50F, r.Top + r.Height * 0.50F));
                    break;
                case "\uE74D": // Trash
                    graphics.DrawLine(pen, r.Left + r.Width * 0.30F, r.Top + r.Height * 0.28F, r.Left + r.Width * 0.70F, r.Top + r.Height * 0.28F);
                    graphics.DrawLine(pen, r.Left + r.Width * 0.42F, r.Top + r.Height * 0.18F, r.Left + r.Width * 0.58F, r.Top + r.Height * 0.18F);
                    using (GraphicsPath bin = BuildRoundedPath(Rectangle.Round(new RectangleF(r.Left + r.Width * 0.32F, r.Top + r.Height * 0.34F, r.Width * 0.36F, r.Height * 0.46F)), 3))
                    {
                        graphics.FillPath(brush, bin);
                        graphics.DrawPath(pen, bin);
                    }
                    graphics.DrawLine(pen, r.Left + r.Width * 0.45F, r.Top + r.Height * 0.42F, r.Left + r.Width * 0.45F, r.Top + r.Height * 0.72F);
                    graphics.DrawLine(pen, r.Left + r.Width * 0.56F, r.Top + r.Height * 0.42F, r.Left + r.Width * 0.56F, r.Top + r.Height * 0.72F);
                    break;
                case "\uE8D7": // Channel/update
                    graphics.DrawArc(pen, r.Left + 3, r.Top + 4, r.Width - 8, r.Height - 8, 35, 220);
                    graphics.DrawArc(pen, r.Left + 5, r.Top + 4, r.Width - 8, r.Height - 8, 215, 220);
                    graphics.DrawLine(pen, r.Right - 6, r.Top + r.Height * 0.35F, r.Right - 3, r.Top + r.Height * 0.20F);
                    graphics.DrawLine(pen, r.Right - 6, r.Top + r.Height * 0.35F, r.Right - 17, r.Top + r.Height * 0.34F);
                    break;
                case "\uE90F": // Repair
                    graphics.DrawLine(pen, r.Left + r.Width * 0.25F, r.Top + r.Height * 0.74F, r.Left + r.Width * 0.68F, r.Top + r.Height * 0.31F);
                    graphics.DrawEllipse(pen, r.Left + r.Width * 0.57F, r.Top + r.Height * 0.17F, r.Width * 0.24F, r.Height * 0.24F);
                    graphics.DrawLine(pen, r.Left + r.Width * 0.24F, r.Top + r.Height * 0.80F, r.Left + r.Width * 0.16F, r.Top + r.Height * 0.72F);
                    break;
                case "\uE720": // Teams
                    graphics.DrawString("T", new Font("Segoe UI", rect.Height * 0.72F, FontStyle.Bold, GraphicsUnit.Pixel), new SolidBrush(color), r.Left + r.Width * 0.18F, r.Top + r.Height * 0.15F);
                    graphics.FillEllipse(brush, r.Left + r.Width * 0.62F, r.Top + r.Height * 0.22F, r.Width * 0.18F, r.Height * 0.18F);
                    graphics.DrawEllipse(pen, r.Left + r.Width * 0.62F, r.Top + r.Height * 0.22F, r.Width * 0.18F, r.Height * 0.18F);
                    break;
                case "\uE715": // Outlook
                    using (GraphicsPath envelope = BuildRoundedPath(Rectangle.Round(new RectangleF(r.Left + r.Width * 0.16F, r.Top + r.Height * 0.30F, r.Width * 0.68F, r.Height * 0.42F)), 4))
                    {
                        graphics.FillPath(brush, envelope);
                        graphics.DrawPath(pen, envelope);
                    }
                    graphics.DrawLine(pen, r.Left + r.Width * 0.18F, r.Top + r.Height * 0.34F, r.Left + r.Width * 0.50F, r.Top + r.Height * 0.55F);
                    graphics.DrawLine(pen, r.Left + r.Width * 0.82F, r.Top + r.Height * 0.34F, r.Left + r.Width * 0.50F, r.Top + r.Height * 0.55F);
                    break;
                case "\uE946": // About/settings
                    graphics.DrawEllipse(pen, r.Left + r.Width * 0.28F, r.Top + r.Height * 0.22F, r.Width * 0.44F, r.Height * 0.44F);
                    graphics.DrawLine(pen, r.Left + r.Width * 0.50F, r.Top + r.Height * 0.42F, r.Left + r.Width * 0.50F, r.Top + r.Height * 0.55F);
                    graphics.DrawLine(pen, r.Left + r.Width * 0.50F, r.Top + r.Height * 0.74F, r.Left + r.Width * 0.50F, r.Top + r.Height * 0.76F);
                    break;
                default:
                    graphics.DrawEllipse(pen, r);
                    break;
            }
        }

        private void UpdateSidebarChrome()
        {
            _brandPanel.Visible = false;

            _sidebarNavSurface.Visible = false;
            _sidebarNavSurface.SendToBack();
        }

        private static Panel CreateSidebarNavSurface()
        {
            var panel = new Panel
            {
                BackColor = WorkbenchUi.SidebarNavSurfaceColor
            };

            void UpdateRegion()
            {
                if (panel.Width <= 1 || panel.Height <= 1)
                {
                    return;
                }

                using var path = BuildRoundedPath(new Rectangle(0, 0, panel.Width, panel.Height), 22);
                panel.Region = new Region(path);
            }

            panel.SizeChanged += (_, _) =>
            {
                UpdateRegion();
                panel.Invalidate();
            };

            panel.Paint += (_, e) =>
            {
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                using var path = BuildRoundedPath(new Rectangle(0, 0, panel.Width - 1, panel.Height - 1), 22);
                using var pen = new Pen(Color.FromArgb(233, 241, 250));
                e.Graphics.DrawPath(pen, path);
            };

            UpdateRegion();
            return panel;
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

        private void HandleSignIn()
        {
            ActiveControl = null;
            PromptPortalVerification();
        }

        private void HandleSignOut()
        {
            ActiveControl = null;
            if (!HasAccessGranted())
            {
                return;
            }

            DialogResult confirm = MessageBox.Show(
                T("注销将重置当前 21V 账户验证状态。是否继续？", "Signing out will reset the current 21V account verification state. Continue?"),
                T("确认注销", "Confirm Sign out"),
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (confirm != DialogResult.Yes)
            {
                return;
            }

            ResetVerificationState();
            ApplyLanguage();
            ShowPage("home");
        }

        private void ApplyLanguageToPage(Control page)
        {
            switch (page)
            {
                case WorkbenchHomeControl homePage:
                    homePage.ApplyLanguage(_currentLanguage);
                    break;
                case WorkbenchInstallationControl installPage:
                    installPage.ApplyLanguage(_currentLanguage);
                    break;
                case WorkbenchUninstallControl uninstallPage:
                    uninstallPage.ApplyLanguage(_currentLanguage);
                    break;
                case WorkbenchChannelControl channelPage:
                    channelPage.ApplyLanguage(_currentLanguage);
                    break;
                case WorkbenchRepairControl repairPage:
                    repairPage.ApplyLanguage(_currentLanguage);
                    break;
                case WorkbenchTeamsControl teamsPage:
                    teamsPage.ApplyLanguage(_currentLanguage);
                    break;
                case WorkbenchOutlookControl outlookPage:
                    outlookPage.ApplyLanguage(_currentLanguage);
                    break;
                case WorkbenchSettingsControl settingsPage:
                    settingsPage.ApplyLanguage(_currentLanguage);
                    break;
            }
        }

        private Control GetOrCreatePage(string key)
        {
            if (_pageCache.TryGetValue(key, out Control? cached))
            {
                return cached;
            }

            Control page = key switch
            {
                "home" => BuildHomePage(),
                "install" => new WorkbenchInstallationControl(_currentLanguage),
                "teams" => new WorkbenchTeamsControl(_teamsMaintenanceService),
                "outlook" => new WorkbenchOutlookControl(_outlookDiagnosticsService),
                "uninstall" => new WorkbenchUninstallControl(_officeUninstallService),
                "channel" => new WorkbenchChannelControl(_officeChannelService),
                "repair" => new WorkbenchRepairControl(_cleanupService, _networkRepairService),
                "settings" => BuildSettingsPage(),
                _ => BuildPlaceholderPage("页面开发中", "当前页面尚未接入到新工作台")
            };

            _pageCache[key] = page;
            return page;
        }

        private Control BuildHomePage()
        {
            var homePage = new WorkbenchHomeControl(_appSettingsService, _appSettings, _currentLanguage);
            homePage.ModuleSelected += (_, pageKey) => ShowPage(pageKey);
            return homePage;
        }

        private Control BuildSettingsPage()
        {
            var settings = new WorkbenchSettingsControl(_appSettingsService, _appSettings, _currentLanguage);
            settings.LanguageModeChanged += (_, mode) => OnLanguageModeChanged(mode);
            return settings;
        }

        private static Control BuildPlaceholderPage(string title, string description)
        {
            var panel = new Panel
            {
                BackColor = Color.White,
                Dock = DockStyle.Fill,
                Padding = new Padding(24)
            };

            var card = new Panel
            {
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Location = new Point(24, 24),
                Size = new Size(860, 180),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };

            var lblTitle = new Label
            {
                Text = title,
                Font = WorkbenchUi.CreateUiFont(16F, FontStyle.Bold),
                ForeColor = Color.FromArgb(32, 32, 32),
                Location = new Point(24, 24),
                Size = new Size(360, 32)
            };

            var lblDescription = new Label
            {
                Text = description,
                Font = WorkbenchUi.CreateUiFont(10F),
                ForeColor = Color.FromArgb(96, 96, 96),
                Location = new Point(24, 68),
                Size = new Size(760, 52)
            };

            card.Controls.Add(lblTitle);
            card.Controls.Add(lblDescription);
            panel.Controls.Add(card);
            return panel;
        }

        private (string Title, string Description) GetPageHeader(string key)
        {
            if (_currentLanguage == UiLanguage.English)
            {
                return key switch
                {
                    "home" => ("Home", "Welcome to 21Vianet Office Toolbox"),
                    "install" => ("Install", "Install M365 Apps, Project, or Visio online with pre-checks"),
                    "uninstall" => ("Uninstall", "Run official SaRA / Get Help scenario to uninstall Office with logs"),
                    "channel" => ("Update Channel", "Switch Office update channel or target a specific build"),
                    "repair" => ("Cleanup & Repair", "Run activation cleanup and network/activation environment repair"),
                    "teams" => ("Teams Tools", "Handle new Teams cache/sign-in records and Outlook meeting add-in issues"),
                    "outlook" => ("Outlook Tools", "Run Outlook full scan, offline scan, and calendar scan"),
                    "settings" => ("About", "Language settings and support channels"),
                    _ => ("Page", "This page is not connected yet")
                };
            }

            return key switch
            {
                "home" => ("首页", "欢迎使用 21Vianet Office工具箱"),
                "install" => ("安装", "在线安装 M365 应用、Project 或 Visio，并保留必要的安装前检查"),
                "uninstall" => ("卸载", "调用微软官方 SaRA / Get Help 场景执行 Office 卸载并保留日志"),
                "channel" => ("更新频道", "切换 Office 更新频道，或按目标版本执行更新"),
                "repair" => ("清理与修复", "集中处理激活清理和网络与激活环境修复"),
                "teams" => ("Teams 工具", "处理新版 Teams 的缓存、登录记录和 Outlook 会议插件问题"),
                "outlook" => ("Outlook 工具", "执行 Outlook 全量扫描、离线扫描与日历扫描"),
                "settings" => ("关于", "查看语言设置、验证状态与反馈支持入口"),
                _ => ("页面", "当前页面尚未接入到新工作台")
            };
        }
    }
}
