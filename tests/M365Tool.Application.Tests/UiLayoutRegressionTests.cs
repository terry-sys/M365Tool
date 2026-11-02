using System.Drawing;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Windows.Forms;
using Office365CleanupTool;
using Office365CleanupTool.Services;
using Xunit;

namespace M365Tool.Application.Tests;

public class UiLayoutRegressionTests
{
    public static IEnumerable<object[]> PageCases()
    {
        foreach (string page in new[] { "home", "install", "uninstall", "channel", "repair", "teams", "outlook", "settings" })
        foreach (var size in new[] { new Size(750, 450), new Size(750, 820), new Size(800, 820), new Size(980, 820), new Size(1000, 650), new Size(1000, 820), new Size(1100, 820), new Size(1280, 820), new Size(1400, 820), new Size(1760, 1100) })
        foreach (UiLanguage language in new[] { UiLanguage.Chinese, UiLanguage.English })
            yield return new object[] { page, size.Width, size.Height, language };
    }

    [Theory]
    [MemberData(nameof(PageCases))]
    public void PagesKeepControlsReadableAndActionsReachable(string page, int width, int height, UiLanguage language) => RunSta(() =>
    {
        using var control = CreatePage(page);
        control.Size = new Size(width, height);
        control.GetType().GetMethod("ApplyLanguage")!.Invoke(control, new object[] { language });
        control.GetType().GetMethod("LayoutControls", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(control, null);
        var problems = new List<string>();
        CheckLayout(control, problems);
        Assert.True(problems.Count == 0, string.Join(Environment.NewLine, problems));
    });

    [Theory]
    [InlineData(UiLanguage.Chinese)]
    [InlineData(UiLanguage.English)]
    public void NetworkRepairDialogButtonsFitTheirCaptions(UiLanguage language) => RunSta(() =>
    {
        using var dialog = new NetworkRepairOptionsForm(new NetworkRepairService(new NeverRun()), language);
        foreach (Button button in dialog.Controls.OfType<Button>())
            Assert.True(TextRenderer.MeasureText(button.Text, button.Font).Width + 16 <= button.Width,
                $"Caption does not fit: {button.Text}");
    });

    [Fact]
    public void OutlineIconsCoverAllNavigationEntriesAndScaledStates() => RunSta(() =>
    {
        var getIcon = typeof(WorkbenchForm).GetMethod("TryGetNavPngIcon", BindingFlags.Static | BindingFlags.NonPublic)!;
        foreach (string key in new[] { "home", "install", "uninstall", "channel", "repair", "teams", "outlook", "settings", "menu" })
        foreach (int size in new[] { 24, 28, 32, 42 })
        foreach (bool enabled in new[] { false, true })
        {
            object?[] args = { key, enabled, size, null };
            Assert.True((bool)getIcon.Invoke(null, args)!, $"Missing icon: {key}");
            var icon = Assert.IsType<Bitmap>(args[3]);
            Assert.Equal(new Size(size, size), icon.Size);
            Assert.Equal(0, icon.GetPixel(0, 0).A);
            Assert.True(Enumerable.Range(0, size).Any(y => Enumerable.Range(0, size).Any(x => icon.GetPixel(x, y).A > 0)), $"Blank icon: {key}");
        }
    });

    [Theory]
    [InlineData(1F)]
    [InlineData(1.25F)]
    [InlineData(1.32F)]
    public void ExpandedNavigationKeepsAlignedColumnsAndReadableSpacing(float scale) => RunSta(() =>
    {
        var apply = typeof(WorkbenchForm).GetMethod("ApplyExpandedNavIcon", BindingFlags.Static | BindingFlags.NonPublic)!;
        string[] keys = { "home", "install", "uninstall", "channel", "repair", "teams", "outlook", "settings" };
        string[][] labels =
        {
            new[] { "首页", "安装", "卸载", "更新频道", "清理与修复", "Teams 工具", "Outlook 工具", "关于" },
            new[] { "Home", "Install", "Uninstall", "Update Channel", "Cleanup & Repair", "Teams Tools", "Outlook Tools", "About" }
        };
        foreach (string[] language in labels)
        {
            int? textColumn = null;
            for (int i = 0; i < keys.Length; i++)
            {
                using var button = new WorkbenchNavigationButton
                {
                    Text = language[i], Size = WorkbenchUi.ScaleSize(new Size(218, 64), scale),
                    Font = WorkbenchUi.CreateUiFont(10.6F * scale, FontStyle.Bold),
                    BackColor = Color.White, ForeColor = WorkbenchUi.SecondaryTextColor
                };
                apply.Invoke(null, new object[] { button, keys[i], true, scale });
                (Rectangle icon, Rectangle text) = button.GetContentBounds();
                textColumn ??= text.Left;
                Assert.Equal(textColumn.Value, text.Left);
                Assert.Equal(WorkbenchUi.Scale(12, scale), text.Left - icon.Right);
                Assert.InRange(Math.Abs(icon.Top + icon.Height / 2F - (text.Top + text.Height / 2F)), 0, 1);
                Size measured = TextRenderer.MeasureText(button.Text, button.Font, Size.Empty,
                    TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding);
                Assert.True(measured.Width <= text.Width, $"Navigation label clipped: {button.Text}");
                using var bitmap = new Bitmap(button.Width, button.Height);
                button.DrawToBitmap(bitmap, button.ClientRectangle);
                Rectangle iconInk = GetInkBounds(bitmap, icon);
                Rectangle textInk = GetInkBounds(bitmap, text);
                Assert.True(textInk.Left - iconInk.Right >= WorkbenchUi.Scale(12, scale));
                Assert.True(Math.Abs(iconInk.Top + iconInk.Height / 2F - (textInk.Top + textInk.Height / 2F)) <= 2,
                    $"Visible alignment differs for {button.Text}: icon {iconInk}, text {textInk}.");
            }
        }
    });

    private static Rectangle GetInkBounds(Bitmap bitmap, Rectangle area)
    {
        var pixels = new List<Point>();
        for (int y = area.Top; y < area.Bottom; y++)
        for (int x = area.Left; x < area.Right; x++)
        {
            Color color = bitmap.GetPixel(x, y);
            if (color.A > 0 && (color.R < 180 || color.G < 180 || color.B < 180)) pixels.Add(new Point(x, y));
        }
        Assert.NotEmpty(pixels);
        return Rectangle.FromLTRB(pixels.Min(p => p.X), pixels.Min(p => p.Y), pixels.Max(p => p.X) + 1, pixels.Max(p => p.Y) + 1);
    }

    [Fact]
    public void FieldLabelsHaveReadableContrastOnLightCards()
    {
        using var label = WorkbenchUi.CreateFieldLabel("Target Channel", Point.Empty);
        static double Luminance(Color color)
        {
            static double Channel(byte value)
            {
                double linear = value / 255D;
                return linear <= 0.04045 ? linear / 12.92 : Math.Pow((linear + 0.055) / 1.055, 2.4);
            }
            return 0.2126 * Channel(color.R) + 0.7152 * Channel(color.G) + 0.0722 * Channel(color.B);
        }
        double contrast = (Luminance(Color.White) + 0.05) / (Luminance(label.ForeColor) + 0.05);
        Assert.True(contrast >= 4.5, $"Field label contrast is only {contrast:F2}:1.");
    }

    [Fact]
    public void NavigationPaintingPreservesInteractionAndStateFeedback() => RunSta(() =>
    {
        using var button = new WorkbenchNavigationButton
        {
            Text = "首页", AccessibleName = "首页", Size = new Size(218, 64),
            BackColor = Color.FromArgb(229, 242, 255)
        };
        button.FlatAppearance.MouseOverBackColor = Color.FromArgb(221, 237, 255);
        button.FlatAppearance.MouseDownBackColor = Color.FromArgb(207, 229, 255);
        static Color RenderBackground(WorkbenchNavigationButton control)
        {
            using var image = new Bitmap(control.Width, control.Height);
            control.DrawToBitmap(image, control.ClientRectangle);
            return image.GetPixel(control.Width - 20, control.Height / 2);
        }
        Assert.Equal(button.BackColor.ToArgb(), RenderBackground(button).ToArgb());
        typeof(WorkbenchNavigationButton).GetMethod("OnMouseEnter", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(button, new object[] { EventArgs.Empty });
        Assert.Equal(button.FlatAppearance.MouseOverBackColor.ToArgb(), RenderBackground(button).ToArgb());
        typeof(WorkbenchNavigationButton).GetMethod("OnMouseDown", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(button, new object[] { new MouseEventArgs(MouseButtons.Left, 1, 10, 10, 0) });
        Assert.Equal(button.FlatAppearance.MouseDownBackColor.ToArgb(), RenderBackground(button).ToArgb());
        int clicks = 0;
        button.Click += (_, _) => clicks++;
        button.PerformClick();
        Assert.Equal(1, clicks);
        Assert.True(button.TabStop);
        Assert.Equal("首页", button.AccessibleName);
    });

    [Fact]
    public void TransparentNavigationRepaintsTheParentAndClearsPreviousFrames() => RunSta(() =>
    {
        using var parent = new Panel { Size = new Size(280, 180), BackColor = Color.White };
        parent.Paint += (_, e) => e.Graphics.FillRectangle(Brushes.AliceBlue, 0, 0, parent.Width, parent.Height);
        using var button = new WorkbenchNavigationButton
        {
            Text = "很长的旧标题", Location = new Point(20, 30), Size = new Size(218, 64), BackColor = Color.Transparent
        };
        parent.Controls.Add(button);
        var paint = typeof(WorkbenchNavigationButton).GetMethod("OnPaint", BindingFlags.Instance | BindingFlags.NonPublic)!;
        using var image = new Bitmap(button.Width, button.Height);
        using var graphics = Graphics.FromImage(image);
        graphics.Clear(Color.Black);
        paint.Invoke(button, new object[] { new PaintEventArgs(graphics, button.ClientRectangle) });
        Assert.Equal(Color.AliceBlue.ToArgb(), image.GetPixel(195, 32).ToArgb());
        button.Text = "首页";
        paint.Invoke(button, new object[] { new PaintEventArgs(graphics, button.ClientRectangle) });
        using var fresh = new Bitmap(button.Width, button.Height);
        using var freshGraphics = Graphics.FromImage(fresh);
        freshGraphics.Clear(Color.Magenta);
        paint.Invoke(button, new object[] { new PaintEventArgs(freshGraphics, button.ClientRectangle) });
        for (int y = 0; y < image.Height; y++)
        for (int x = 0; x < image.Width; x++)
            Assert.Equal(fresh.GetPixel(x, y), image.GetPixel(x, y));
    });

    [Fact]
    public void ExpandingNavigationRestoresTheLabelFont() => RunSta(() =>
    {
        using var button = new WorkbenchNavigationButton { Text = "Teams 工具", Size = new Size(218, 64) };
        var flags = BindingFlags.Static | BindingFlags.NonPublic;
        typeof(WorkbenchForm).GetMethod("ApplyCollapsedNavIcon", flags)!.Invoke(null, new object[] { button, "teams", true, 1F });
        button.Text = "Teams 工具";
        typeof(WorkbenchForm).GetMethod("ApplyExpandedNavIcon", flags)!.Invoke(null, new object[] { button, "teams", true, 1F });
        Assert.InRange(button.Font.SizeInPoints, 10.0F, 10.2F);
    });

    [Fact]
    public void EnglishSignInDialogStartsWithEnglishStatus() => RunSta(() =>
    {
        using var dialog = new MyAccountLoginVerificationForm(UiLanguage.English);
        var status = (Label)typeof(MyAccountLoginVerificationForm).GetField("_statusLabel", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(dialog)!;
        Assert.Equal("Initializing 21V sign-in page...", status.Text);
    });

    [Theory]
    [InlineData(1024, 728)]
    [InlineData(1920, 1040)]
    public void SignInDialogFitsTheAvailableWorkingArea(int width, int height) => RunSta(() =>
    {
        using var dialog = new MyAccountLoginVerificationForm(UiLanguage.English);
        dialog.FitToWorkingArea(new Rectangle(0, 0, width, height));
        Assert.InRange(dialog.Width, 1, width);
        Assert.InRange(dialog.Height, 1, height);
        Assert.Equal(dialog.Size, dialog.MinimumSize);
        Assert.Equal(dialog.Size, dialog.MaximumSize);
    });

    [Theory]
    [InlineData(Keys.Enter)]
    [InlineData(Keys.Space)]
    public void HomeModuleCardsCanBeActivatedWithTheKeyboard(Keys key) => RunSta(() =>
    {
        using var home = new WorkbenchHomeControl(new AppSettingsService(), new AppSettings(), UiLanguage.English);
        var selected = new List<string>();
        home.ModuleSelected += (_, page) => selected.Add(page);
        var workspace = (Panel)typeof(WorkbenchHomeControl).GetField("_workspaceCard", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(home)!;
        foreach (Panel card in workspace.Controls.OfType<Panel>().Where(panel => panel.AccessibleRole == AccessibleRole.PushButton))
        {
            Assert.True(card.TabStop);
            Assert.True(card.CanSelect);
            card.GetType().GetMethod("OnKeyDown", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(card, new object[] { new KeyEventArgs(key) });
        }
        Assert.Equal(new[] { "install", "uninstall", "channel", "repair", "teams", "outlook" }, selected);
    });

    [Theory]
    [InlineData("install", "_txtResultDetails", "Installation result")]
    [InlineData("uninstall", "_txtResult", "View uninstall results")]
    [InlineData("channel", "_txtResult", "View channel switch")]
    [InlineData("repair", "_txtResult", "View cleanup")]
    [InlineData("teams", "_txtResult", "View Teams")]
    [InlineData("outlook", "_txtResult", "View Outlook")]
    public void LanguageChangesPreserveExecutionDetails(string page, string fieldName, string text) => RunSta(() =>
    {
        using var control = CreatePage(page);
        var results = (TextBox)control.GetType().GetField(fieldName, BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(control)!;
        string detail = $"Failure details: {text}; exit code 123";
        results.Text = detail;
        control.GetType().GetMethod("ApplyLanguage")!.Invoke(control, new object[] { UiLanguage.English });
        Assert.Equal(detail, results.Text);
    });

    [Theory]
    [InlineData("channel", "_settingsCard")]
    [InlineData("uninstall", "_targetCard")]
    [InlineData("home", "_workspaceCard")]
    [InlineData("install", "_basicCard")]
    [InlineData("repair", "_cardsHost")]
    [InlineData("teams", "_cardsHost")]
    [InlineData("outlook", "_cardsHost")]
    [InlineData("settings", "_settingsCard")]
    public void RelayoutWhileScrolledPreservesLogicalSectionPosition(string page, string firstCard) => RunSta(() =>
    {
        using var control = (ScrollableControl)CreatePage(page);
        control.Size = new Size(1000, 450);
        var card = (Panel)control.GetType().GetField(firstCard, BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(control)!;
        var layout = control.GetType().GetMethod("LayoutControls", BindingFlags.Instance | BindingFlags.NonPublic)!;
        layout.Invoke(control, null);
        control.CreateControl();
        var scrollHost = control.AutoScroll ? control : (ScrollableControl)control.GetType().GetField("_scrollHost", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(control)!;
        int logicalTop = card.Top;
        var extent = scrollHost.DisplayRectangle.Size;
        scrollHost.AutoScrollPosition = new Point(0, 150);
        Assert.True(scrollHost.AutoScrollPosition.Y < 0);
        for (int i = 0; i < 3; i++) layout.Invoke(control, null);
        Assert.Equal(logicalTop, card.Top - (card.Parent == scrollHost ? scrollHost.AutoScrollPosition.Y : 0));
        Assert.Equal(extent, scrollHost.DisplayRectangle.Size);
    });

    private static Control CreatePage(string page)
    {
        var runner = new NeverRun();
        return page switch
        {
            "home" => new WorkbenchHomeControl(new AppSettingsService(), new AppSettings(), UiLanguage.Chinese),
            "install" => new WorkbenchInstallationControl(UiLanguage.Chinese),
            "uninstall" => new WorkbenchUninstallControl(new OfficeUninstallService(runner)),
            "channel" => new WorkbenchChannelControl(new OfficeChannelService()),
            "repair" => new WorkbenchRepairControl(new CleanupService(runner), new NetworkRepairService(runner)),
            "teams" => new WorkbenchTeamsControl(new TeamsMaintenanceService()),
            "outlook" => new WorkbenchOutlookControl(new OutlookDiagnosticsService()),
            "settings" => new WorkbenchSettingsControl(new AppSettingsService(), new AppSettings(), UiLanguage.Chinese),
            _ => throw new ArgumentException(page)
        };
    }

    private static void CheckLayout(Control parent, List<string> problems)
    {
        Control[] children = parent.Controls.Cast<Control>().Where(control => control.Visible).ToArray();
        for (int i = 0; i < children.Length; i++)
        {
            Control control = children[i];
            if (!(parent is ScrollableControl scroll && scroll.AutoScroll) && !parent.ClientRectangle.Contains(control.Bounds))
                problems.Add($"Outside parent: {control.GetType().Name} {control.Text}");
            for (int j = i + 1; j < children.Length; j++)
                if (control.Bounds.IntersectsWith(children[j].Bounds))
                    problems.Add($"Overlapping: {control.Text} / {children[j].Text}");
            if (control is Button button)
            {
                Assert.True(button.TabStop, $"Button skipped by Tab: {button.Text}");
                Size textSize = TextRenderer.MeasureText(button.Text, button.Font, Size.Empty, TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix);
                if (textSize.Width > button.Width || textSize.Height > button.Height)
                    problems.Add($"Button caption clipped: {button.Text}");
            }
            if (control is TextBox { ReadOnly: true } textBox)
                Assert.True(textBox.TabStop, "Read-only results must be reachable for keyboard scrolling/copying.");
            if (control is Label { AutoSize: false, AutoEllipsis: false } label && !string.IsNullOrWhiteSpace(label.Text))
            {
                Size textSize = TextRenderer.MeasureText(label.Text, label.Font, new Size(Math.Max(1, label.Width), int.MaxValue), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix);
                if (textSize.Height > label.Height + 1)
                    problems.Add($"Label clipped: {label.Text} needs {textSize.Height}, has {label.Height}");
            }
            CheckLayout(control, problems);
        }
    }

    private static void RunSta(Action action)
    {
        Exception? error = null;
        var thread = new Thread(() =>
        {
            try { action(); }
            catch (Exception ex) { error = ex; }
        }) { IsBackground = true };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        Assert.True(thread.Join(TimeSpan.FromSeconds(15)), "UI layout test timed out.");
        if (error != null) ExceptionDispatchInfo.Capture(error).Throw();
    }

    private sealed class NeverRun : IScriptRunner
    {
        public Task<ScriptExecutionResult> RunAsync(ScriptExecutionRequest request) => throw new InvalidOperationException("Maintenance is disabled in UI tests.");
        public void StartDetached(ScriptExecutionRequest request) => throw new InvalidOperationException("Maintenance is disabled in UI tests.");
    }
}
