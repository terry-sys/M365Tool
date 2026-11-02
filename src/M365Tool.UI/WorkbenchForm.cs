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
        private sealed class AccountAvatarLabel : Label
        {
            protected override void OnPaintBackground(PaintEventArgs e)
            {
                e.Graphics.Clear(Color.White);
                if (Image != null) return;
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                using var fill = new SolidBrush(Color.FromArgb(233, 239, 252));
                e.Graphics.FillEllipse(fill, 0, 0, Width - 1, Height - 1);
            }
        }

        private const int ExpandedSidebarWidth = 216;
        private const int CollapsedSidebarWidth = 68;

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
        private readonly Panel _navSelectionBar;
        private readonly Dictionary<string, Control> _pageCache = new();
        private readonly Dictionary<string, Button> _navButtons = new();
        private readonly Dictionary<string, string> _navButtonTexts = new();
        private readonly Dictionary<string, string> _navButtonIcons = new();
        private static readonly IReadOnlyDictionary<string, string> NavPngIconResourceNames = new Dictionary<string, string>
        {
            ["home"] = "M365Tool.Assets.Icons.FluentOutline.home.png",
            ["settings"] = "M365Tool.Assets.Icons.FluentOutline.about.png",
            ["menu"] = "M365Tool.Assets.Icons.FluentOutline.menu.png",
            ["install"] = "M365Tool.Assets.Icons.FluentOutline.install.png",
            ["uninstall"] = "M365Tool.Assets.Icons.FluentOutline.uninstall.png",
            ["channel"] = "M365Tool.Assets.Icons.FluentOutline.channel.png",
            ["repair"] = "M365Tool.Assets.Icons.FluentOutline.repair.png",
            ["teams"] = "M365Tool.Assets.Icons.FluentOutline.teams.png",
            ["outlook"] = "M365Tool.Assets.Icons.FluentOutline.outlook.png"
        };

        private static readonly Dictionary<(int Size, bool Enabled), IReadOnlyDictionary<string, Image>> NavPngIconImages = new();
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

        private readonly System.Windows.Forms.Timer _signInPreparationTimer = new() { Interval = 750 };
        private MyAccountLoginVerificationForm? _preparedSignInDialog;
        private UiLanguage _preparedSignInLanguage;
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
            _sessionVerified = _appSettings.HasCurrentVerification &&
                               (!string.IsNullOrWhiteSpace(_appSettings.LastValidatedAccount) ||
                                !string.IsNullOrWhiteSpace(_appSettings.LastValidatedDisplayName));

            Text = "21Vianet Office工具箱";
            Icon = AppIconProvider.GetAppIcon() ?? Icon;
            ShowIcon = false;
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(1024, 720);
            ClientSize = GetInitialClientSize();
            BackColor = Color.White;

            _sidebar = new Panel
            {
                Dock = DockStyle.Left,
                Width = ExpandedSidebarWidth,
                BackColor = WorkbenchUi.SidebarBackgroundColor
            };
            _sidebar.Paint += (_, e) => { PaintSidebarBackground(e.Graphics, _sidebar.ClientRectangle); PaintSidebarGroups(e.Graphics); };

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
            _headerPanel.BackColor = Color.White;
            _headerPanel.Paint += (_, e) => e.Graphics.DrawLine(Pens.Gainsboro, 0, _headerPanel.Height - 1, _headerPanel.Width, _headerPanel.Height - 1);

            _statusBar = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 32,
                BackColor = WorkbenchUi.CanvasColor,
                Visible = true
            };

            _contentHost = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent,
                Padding = Padding.Empty,
                AutoScroll = false
            };

            _lblAppTitle = new Label
            {
                Text = "Office 工具箱",
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
                TextAlign = ContentAlignment.MiddleLeft,
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
            _accountPanel.BackColor = Color.White;

            _lblAccountAvatar = new AccountAvatarLabel
            {
                Font = WorkbenchUi.CreateUiFont(13F, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = WorkbenchUi.PrimaryColor,
                ImageAlign = ContentAlignment.MiddleCenter,
                TextAlign = ContentAlignment.MiddleCenter,
                UseMnemonic = false
            };
            _lblAccountAvatar.Paint += (_, e) => PaintGuestAccountIcon(e.Graphics, _lblAccountAvatar.ClientRectangle);

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

            _btnSignIn = WorkbenchUi.CreatePrimaryButton("登录", Point.Empty, new Size(80, 32), 7);
            _btnSignIn.Font = WorkbenchUi.CreateUiFont(9F, FontStyle.Bold);
            _btnSignIn.TabStop = true;
            _btnSignIn.Click += (_, _) => HandleSignIn();

            _btnSignOut = WorkbenchUi.CreateTextButton("注销", Point.Empty, new Size(54, 30), 15);
            _btnSignOut.Font = WorkbenchUi.CreateUiFont(8.9F, FontStyle.Bold);
            _btnSignOut.TabStop = true;
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
            sidebarDivider.BackColor = WorkbenchUi.CardBorderColor;
            headerDivider.BackColor = WorkbenchUi.CardBorderColor;

            _sidebar.Controls.Add(sidebarDivider);
            _sidebar.Controls.Add(_sidebarNavSurface);
            _sidebar.Controls.Add(_navSelectionBar);
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
            _statusBar.Paint += (_, e) =>
            {
                bool busy = MaintenanceCoordinator.Shared.IsBusy;
                using var font = WorkbenchUi.CreateUiFont(8.5F);
                using var dot = new SolidBrush(busy ? WorkbenchUi.PrimaryColor : WorkbenchUi.SuccessColor);
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.FillEllipse(dot, 34, (_statusBar.Height - 5) / 2, 5, 5);
                TextRenderer.DrawText(e.Graphics, busy ? T("维护任务运行中", "Maintenance task running") : T("就绪", "Ready"), font,
                    new Rectangle(47, 0, 260, _statusBar.Height), WorkbenchUi.SecondaryTextColor,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
            };

            Controls.Add(_contentHost);
            Controls.Add(_statusBar);
            Controls.Add(_headerPanel);
            Controls.Add(_sidebar);

            bool lastBusy = MaintenanceCoordinator.Shared.IsBusy;
            var statusRefresh = new System.Windows.Forms.Timer { Interval = 1000 };
            statusRefresh.Tick += (_, _) =>
            {
                bool busy = MaintenanceCoordinator.Shared.IsBusy;
                if (busy != lastBusy) { lastBusy = busy; _statusBar.Invalidate(); }
            };
            Shown += (_, _) => statusRefresh.Start();
            Disposed += (_, _) => statusRefresh.Dispose();

            _signInPreparationTimer.Tick += async (_, _) =>
            {
                _signInPreparationTimer.Stop();
                await PrepareSignInDialogAsync();
            };
            Disposed += (_, _) =>
            {
                _signInPreparationTimer.Dispose();
                _preparedSignInDialog?.Dispose();
                _preparedSignInDialog = null;
            };

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
            ScheduleSignInPreparation();
        }

        private static void PaintSidebarBackground(Graphics graphics, Rectangle bounds)
        {
            using var brush = new SolidBrush(WorkbenchUi.SidebarBackgroundColor);
            graphics.FillRectangle(brush, bounds);
        }

        private void PaintSidebarGroups(Graphics graphics)
        {
            if (_isSidebarCollapsed || _navButtons.Count == 0) return;
            float scale = GetShellScale();
            int S(int value) => WorkbenchUi.Scale(value, scale);
            using var font = WorkbenchUi.CreateUiFont(9F * scale);
            foreach (var group in new[] { ("home", T("工作台", "Workspace")),
                ("install", T("应用管理", "App Management")), ("repair", T("诊断与维护", "Diagnostics & Maintenance")) })
            {
                if (!_navButtons.TryGetValue(group.Item1, out var button)) continue;
                TextRenderer.DrawText(graphics, group.Item2, font,
                    new Rectangle(S(29), button.Top - S(30), _sidebar.Width - S(40), S(22)),
                    WorkbenchUi.SecondaryTextColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix);
            }
            using var pen = new Pen(WorkbenchUi.CardBorderColor);
            int dividerTop = _navButtons["settings"].Top - S(24);
            graphics.DrawLine(pen, S(22), dividerTop, _sidebar.Width - S(22), dividerTop);
        }

        private static void PaintBrandLogo(Graphics graphics, Rectangle bounds)
        {
            if (bounds.Width < 36 || bounds.Height < 36) return;
            Icon? icon = AppIconProvider.GetAppIcon();
            if (icon != null) graphics.DrawIcon(icon, new Rectangle(0, 0, 36, 36));
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
            var button = new WorkbenchNavigationButton
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
                TabStop = true,
                UseMnemonic = false
            };

            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = Color.FromArgb(246, 251, 255);
            button.FlatAppearance.MouseDownBackColor = Color.FromArgb(241, 247, 255);
            button.Click += (_, _) => ShowPage(key);
            ApplyRoundedRegion(button, 7);

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
            _lblAppTitle.Text = "Office 工具箱";
            _lblAppSubtitle.Text = T("21Vianet Office 运维工具", "21Vianet Office Toolbox");
            _btnSignIn.Text = T("登录", "Sign in");
            _btnSignOut.Text = T("注销", "Sign out");
            _navToolTip.SetToolTip(_btnSignIn, T("登录并验证 21V 账户", "Sign in and verify 21V account"));
            _navToolTip.SetToolTip(_btnSignOut, T("注销并重置当前验证状态", "Sign out and reset current verification state"));
            UpdateHeaderAccountPresentation();

            foreach ((string key, Button button) in _navButtons)
            {
                string text = GetNavText(key);
                _navButtonTexts[key] = text;
                button.AccessibleName = text;
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
            return _sessionVerified && _appSettings.HasCurrentVerification;
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

        private void ScheduleSignInPreparation()
        {
            if (!IsDisposed && !Disposing && !HasAccessGranted()) _signInPreparationTimer.Start();
        }

        private async Task PrepareSignInDialogAsync()
        {
            if (IsDisposed || Disposing || HasAccessGranted() || _preparedSignInDialog != null) return;
            var dialog = new MyAccountLoginVerificationForm(_currentLanguage);
            _preparedSignInDialog = dialog;
            _preparedSignInLanguage = _currentLanguage;
            try { await dialog.PrepareAsync(); }
            catch
            {
                // A clicked dialog owns its own failure UI; an unused warm-up fails silently.
                if (ReferenceEquals(_preparedSignInDialog, dialog))
                {
                    _preparedSignInDialog = null;
                    dialog.Dispose();
                }
            }
        }

        private bool PromptPortalVerification()
        {
            _signInPreparationTimer.Stop();
            var dialog = _preparedSignInDialog;
            _preparedSignInDialog = null;
            if (dialog != null && (dialog.IsDisposed || _preparedSignInLanguage != _currentLanguage))
            {
                dialog.Dispose();
                dialog = null;
            }
            dialog ??= new MyAccountLoginVerificationForm(_currentLanguage);
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
                _appSettings.VerificationVersion = AppSettings.CurrentVerificationVersion;
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
            ScheduleSignInPreparation();
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
            _sessionVerified = _appSettings.HasCurrentVerification &&
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

        private float GetShellScale() =>
            WorkbenchUi.GetAdaptiveUiScale(ClientSize.Width, Math.Max(1, ClientSize.Height - 76));

        private bool ShouldForceCollapsedSidebar() =>
            ClientSize.Width < 1100;

        private static int GetSidebarWidth(bool isCollapsed, float scale) =>
            WorkbenchUi.Scale(isCollapsed ? CollapsedSidebarWidth : ExpandedSidebarWidth, scale);

        private int GetHeaderBaseHeight() =>
            76;

        private void ApplyWorkbenchLayout()
        {
            _isSidebarCollapsed = ShouldForceCollapsedSidebar();

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
            _sidebar.Width = GetSidebarWidth(_isSidebarCollapsed, scale);
            _brandPanel.Visible = true;
            _brandPanel.SetBounds(_isSidebarCollapsed ? (_sidebar.Width - S(36)) / 2 : S(22), S(25),
                S(_isSidebarCollapsed ? 36 : 172), S(40));
            _lblBrandGlyph.Visible = false;
            _lblAppTitle.Visible = _lblAppSubtitle.Visible = !_isSidebarCollapsed;
            WorkbenchUi.ApplyUiFont(_lblAppTitle, 11.25F, FontStyle.Bold, scale);
            WorkbenchUi.ApplyUiFont(_lblAppSubtitle, 8.5F, FontStyle.Regular, scale);
            _lblAppTitle.SetBounds(S(48), 0, S(124), S(24));
            _lblAppSubtitle.SetBounds(S(49), S(24), S(100), S(18));
            _lblAppSubtitle.Text = "21Vianet";
            string[] order = { "home", "install", "uninstall", "channel", "repair", "teams", "outlook", "settings" };
            int[] tops = { 122, 222, 270, 318, 420, 468, 516 };
            for (int i = 0; i < order.Length; i++)
            {
                Button button = _navButtons[order[i]];
                int top = i == 7 ? _sidebar.Height - S(78) : S(_isSidebarCollapsed ? 108 + i * 52 : tops[i]);
                button.SetBounds(S(_isSidebarCollapsed ? 10 : 14), top,
                    _sidebar.Width - S(_isSidebarCollapsed ? 20 : 28), S(43));
                if (_isSidebarCollapsed) ApplyCollapsedNavIcon(button, order[i], true, scale);
                else { button.Text = _navButtonTexts[order[i]]; ApplyExpandedNavIcon(button, order[i], true, scale); }
            }
            UpdateSidebarChrome();
            UpdateNavigationVisualState();
            _sidebar.Invalidate();
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
            string group = key switch
            {
                "home" => T("工作台", "Workspace"),
                "install" or "uninstall" or "channel" => T("应用管理", "App Management"),
                "settings" => T("工具箱", "Toolbox"),
                _ => T("诊断与维护", "Diagnostics & Maintenance")
            };
            _lblPageTitle.Text = $"{group}    /    {title}";
            bool showDescription = string.Equals(key, "home", StringComparison.OrdinalIgnoreCase) &&
                                   !string.IsNullOrWhiteSpace(description);
            _lblPageDescription.Text = showDescription ? description : string.Empty;
            _lblPageDescription.Visible = false;
        }

        private void UpdateHeaderLayout()
        {
            if (_headerPanel.ClientSize.Width <= 0) return;
            float scale = GetShellScale();
            int S(int value) => WorkbenchUi.Scale(value, scale);
            WorkbenchUi.ApplyUiFont(_lblPageTitle, 9.75F, FontStyle.Regular, scale);
            WorkbenchUi.ApplyUiFont(_lblAccountAvatar, 9F, FontStyle.Bold, scale);
            WorkbenchUi.ApplyUiFont(_lblAccountName, 9.4F, FontStyle.Bold, scale);
            WorkbenchUi.ApplyUiFont(_lblAccountState, 8.25F, FontStyle.Regular, scale);
            WorkbenchUi.ApplyUiFont(_btnSignIn, 9F, FontStyle.Bold, scale);
            WorkbenchUi.ApplyUiFont(_btnSignOut, 9F, FontStyle.Regular, scale);
            bool isVerified = HasAccessGranted();
            Button active = isVerified ? _btnSignOut : _btnSignIn;
            const TextFormatFlags textFlags = TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix;
            int nameWidth = TextRenderer.MeasureText(_lblAccountName.Text, _lblAccountName.Font, Size.Empty, textFlags).Width;
            int stateWidth = TextRenderer.MeasureText(_lblAccountState.Text, _lblAccountState.Font, Size.Empty, textFlags).Width;
            int textWidth = Math.Max(Math.Clamp(nameWidth, S(100), S(172)), stateWidth);
            int buttonWidth = Math.Max(S(isVerified ? 54 : 80),
                TextRenderer.MeasureText(active.Text, active.Font, Size.Empty, textFlags).Width + S(24));
            int accountWidth = S(60) + textWidth + S(24) + buttonWidth;
            _accountPanel.SetBounds(_headerPanel.Width - accountWidth - S(26), S(8), accountWidth, S(60));
            _lblAccountAvatar.SetBounds(0, S(12), S(36), S(36));
            _lblAccountName.SetBounds(S(47), S(8), textWidth, S(24));
            _lblAccountState.SetBounds(S(47), S(32), textWidth, S(19));
            _lblAccountState.Visible = true;
            _lblAccountState.AutoEllipsis = false;
            _accountDivider.SetBounds(S(60) + textWidth, S(17), 1, S(27));
            _accountDivider.Visible = isVerified;
            active.SetBounds(accountWidth - buttonWidth, S(14), buttonWidth, S(32));
            _lblPageTitle.SetBounds(S(34), 0, Math.Max(S(120), _accountPanel.Left - S(60)), _headerPanel.Height);
            _lblPageDescription.Visible = false;
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

            _lblAccountName.Text = isVerified && !string.IsNullOrWhiteSpace(userName)
                ? userName : T("未登录", "Not signed in");
            _lblAccountName.ForeColor = WorkbenchUi.PrimaryTextColor;
            _lblAccountState.Text = isVerified ? T("●  21V 已验证", "●  21V verified") : T("登录以使用全部功能", "Sign in to access all tools");
            _lblAccountState.Visible = true;
            _lblAccountState.ForeColor = isVerified ? WorkbenchUi.SuccessColor : WorkbenchUi.SecondaryTextColor;
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
                _accountAvatarImage ??= TryLoadValidatedAvatar(_appSettings.LastValidatedAvatarPath, 36);
            }

            if (_accountAvatarImage is not null)
            {
                _lblAccountAvatar.Text = string.Empty;
                _lblAccountAvatar.Image = _accountAvatarImage;
                _lblAccountAvatar.BackColor = Color.Transparent;
                return;
            }

            string[] nameParts = userName.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            string initial = isVerified && nameParts.Length > 0
                ? string.Concat(nameParts.Take(2).Select(part => char.ToUpperInvariant(part[0]))) : string.Empty;
            _lblAccountAvatar.Image = null;
            _lblAccountAvatar.Text = initial;
            _lblAccountAvatar.BackColor = Color.Transparent;
            _lblAccountAvatar.ForeColor = WorkbenchUi.PrimaryColor;
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
                        WorkbenchUi.ApplyUiFont(navButton, 10.1F, isActive ? FontStyle.Bold : FontStyle.Regular, scale);
                        navButton.Text = _navButtonTexts[navKey];
                        ApplyExpandedNavIcon(navButton, navKey, isEnabled: false, scale);
                    }

                    continue;
                }

                navButton.Cursor = Cursors.Hand;
                navButton.FlatAppearance.BorderSize = 0;
                navButton.FlatAppearance.MouseOverBackColor = isActive ? Color.FromArgb(225, 234, 253) : Color.FromArgb(243, 245, 248);
                navButton.FlatAppearance.MouseDownBackColor = isActive ? Color.FromArgb(214, 226, 250) : Color.FromArgb(233, 237, 243);
                navButton.BackColor = isActive ? WorkbenchUi.NavigationSelectedColor : Color.Transparent;
                navButton.ForeColor = isActive ? WorkbenchUi.PrimaryColor : WorkbenchUi.PrimaryTextColor;
                if (_isSidebarCollapsed)
                {
                    ApplyCollapsedNavIcon(navButton, navKey, isEnabled: true, scale);
                }
                else
                {
                    WorkbenchUi.ApplyUiFont(navButton, 10.1F, isActive ? FontStyle.Bold : FontStyle.Regular, scale);
                    navButton.Text = _navButtonTexts[navKey];
                    ApplyExpandedNavIcon(navButton, navKey, isEnabled: true, scale);
                    WorkbenchUi.ApplyUiFont(navButton, 10.1F, isActive ? FontStyle.Bold : FontStyle.Regular, scale);
                }
            }

            UpdateSelectionBarPosition();
        }

        private static void ApplyCollapsedNavIcon(Button button, string key, bool isEnabled, float scale)
        {
            if (button is WorkbenchNavigationButton navigation) navigation.UiScale = scale;
            button.Padding = Padding.Empty;
            button.TextImageRelation = TextImageRelation.Overlay;
            if (TryGetNavPngIcon(key, isEnabled, WorkbenchUi.Scale(24, scale), out Image? image))
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

        private static void ApplyExpandedNavIcon(Button button, string key, bool isEnabled, float scale)
        {
            WorkbenchUi.ApplyUiFont(button, 10.1F, FontStyle.Regular, scale);
            if (button is WorkbenchNavigationButton navigation) navigation.UiScale = scale;
            TryGetNavPngIcon(key, isEnabled, WorkbenchUi.Scale(24, scale), out Image? image);
            button.Image = image;
            button.ImageAlign = ContentAlignment.MiddleLeft;
            button.TextImageRelation = TextImageRelation.ImageBeforeText;
            button.Padding = new Padding(WorkbenchUi.Scale(16, scale), 0, WorkbenchUi.Scale(12, scale), 0);
        }

        private static bool TryGetNavPngIcon(string key, bool isEnabled, int size, out Image? image)
        {
            if (!NavPngIconImages.TryGetValue((size, isEnabled), out IReadOnlyDictionary<string, Image>? images))
            {
                images = LoadNavPngIconImages(size, isEnabled ? 1F : 0.38F);
                NavPngIconImages[(size, isEnabled)] = images;
            }
            return images.TryGetValue(key, out image);
        }

        private static IReadOnlyDictionary<string, Image> LoadNavPngIconImages(int size, float alpha)
        {
            var images = new Dictionary<string, Image>();
            var assembly = typeof(WorkbenchForm).Assembly;

            foreach (KeyValuePair<string, string> resource in NavPngIconResourceNames)
            {
                using Stream stream = assembly.GetManifestResourceStream(resource.Value)
                    ?? throw new InvalidOperationException($"Missing sidebar navigation icon resource: {resource.Value}");
                using Image source = Image.FromStream(stream);
                images[resource.Key] = CreateScaledNavIconImage(source, size, alpha);
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
            _brandPanel.Visible = true;

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
            ScheduleSignInPreparation();
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
