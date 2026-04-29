using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace Office365CleanupTool
{
    internal enum WorkbenchTone
    {
        Default,
        Info,
        Success,
        Warning,
        Danger
    }

    internal static class WorkbenchUi
    {
        private sealed class SmoothButton : Button
        {
            private bool _hovered;
            private bool _pressed;

            public Color NormalBackColor { get; set; }
            public Color HoverBackColor { get; set; }
            public Color PressedBackColor { get; set; }
            public Color BorderColor { get; set; }
            public int CornerRadius { get; set; } = 7;

            public SmoothButton()
            {
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.UserPaint, true);
                FlatStyle = FlatStyle.Flat;
                FlatAppearance.BorderSize = 0;
                UseVisualStyleBackColor = false;
                TabStop = false;
            }

            protected override void OnMouseEnter(EventArgs e)
            {
                _hovered = true;
                Invalidate();
                base.OnMouseEnter(e);
            }

            protected override void OnMouseLeave(EventArgs e)
            {
                _hovered = false;
                _pressed = false;
                Invalidate();
                base.OnMouseLeave(e);
            }

            protected override void OnMouseDown(MouseEventArgs mevent)
            {
                if (mevent.Button == MouseButtons.Left)
                {
                    _pressed = true;
                    Invalidate();
                }

                base.OnMouseDown(mevent);
            }

            protected override void OnMouseUp(MouseEventArgs mevent)
            {
                _pressed = false;
                Invalidate();
                base.OnMouseUp(mevent);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                using (var parentFill = new SolidBrush(ResolveParentBackColor()))
                {
                    e.Graphics.FillRectangle(parentFill, ClientRectangle);
                }

                Rectangle rect = new Rectangle(0, 0, Width - 1, Height - 1);
                Color backColor = Enabled
                    ? (_pressed ? PressedBackColor : _hovered ? HoverBackColor : NormalBackColor)
                    : DisabledSurfaceColor;
                Color borderColor = Enabled ? BorderColor : DisabledBorderColor;
                Color textColor = Enabled ? ForeColor : TertiaryTextColor;

                using GraphicsPath path = BuildRoundedPath(rect, CornerRadius);
                using var fill = new SolidBrush(backColor);
                using var border = new Pen(borderColor);
                e.Graphics.FillPath(fill, path);
                e.Graphics.DrawPath(border, path);

                TextRenderer.DrawText(
                    e.Graphics,
                    Text,
                    Font,
                    rect,
                    textColor,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix);
            }

            private Color ResolveParentBackColor()
            {
                Control? current = Parent;
                while (current is not null)
                {
                    if (current.BackColor != Color.Transparent)
                    {
                        return current.BackColor;
                    }

                    current = current.Parent;
                }

                return PageSurfaceColor;
            }
        }

        public static Color PrimaryColor => Color.FromArgb(24, 119, 242);
        public static Color PrimaryHoverColor => Color.FromArgb(18, 103, 220);
        public static Color PrimaryPressedColor => Color.FromArgb(14, 86, 194);
        public static Color PrimaryTextColor => Color.FromArgb(28, 43, 64);
        public static Color SecondaryTextColor => Color.FromArgb(86, 105, 130);
        public static Color TertiaryTextColor => Color.FromArgb(126, 145, 170);
        public static Color CanvasColor => Color.White;
        public static Color PageSurfaceColor => Color.White;
        public static Color SidebarBackgroundColor => Color.White;
        public static Color SidebarNavSurfaceColor => Color.White;
        public static Color CardBorderColor => Color.FromArgb(231, 240, 250);
        public static Color ControlBorderColor => Color.FromArgb(225, 236, 249);
        public static Color MutedSurfaceColor => Color.FromArgb(247, 251, 255);
        public static Color InsetSurfaceColor => Color.FromArgb(250, 253, 255);
        public static Color NavigationSelectedColor => Color.FromArgb(255, 255, 255);
        public static Color DisabledSurfaceColor => Color.FromArgb(243, 245, 248);
        public static Color DisabledBorderColor => Color.FromArgb(226, 232, 240);
        public static Color SuccessColor => Color.FromArgb(31, 122, 77);
        public static Color WarningColor => Color.FromArgb(161, 92, 0);
        public static Color DangerColor => Color.FromArgb(180, 35, 24);

        public static Panel CreateCard(Point location, Size size, Color? backColor = null)
        {
            return CreateSurfacePanel(location, size, backColor ?? Color.White, CardBorderColor, 26);
        }

        public static Panel CreateInsetPanel(Point location, Size size)
        {
            return CreateSurfacePanel(location, size, InsetSurfaceColor, ControlBorderColor, 22);
        }

        public static GroupBox CreateGroup(string title, Point location, Size size)
        {
            return new GroupBox
            {
                Text = title,
                Font = new Font("Microsoft YaHei UI", 10.5F, FontStyle.Bold),
                Location = location,
                Size = size,
                BackColor = Color.White
            };
        }

        public static Label CreateSectionTitle(string text, Point location, int width = 220)
        {
            return new Label
            {
                Text = text,
                Font = new Font("Microsoft YaHei UI", 11F, FontStyle.Bold),
                ForeColor = PrimaryTextColor,
                BackColor = Color.Transparent,
                Location = location,
                Size = new Size(width, 26),
                UseMnemonic = false
            };
        }

        public static Label CreateFieldLabel(string text, Point location, int width = 120)
        {
            return new Label
            {
                Text = text,
                Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold),
                ForeColor = TertiaryTextColor,
                BackColor = Color.Transparent,
                Location = location,
                Size = new Size(width, 24),
                UseMnemonic = false
            };
        }

        public static Label CreateTagLabel(string text, Point location, int width = 72, WorkbenchTone tone = WorkbenchTone.Info)
        {
            var label = new Label
            {
                Text = text,
                Location = location,
                Size = new Size(width, 18),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft YaHei UI", 7.5F, FontStyle.Bold),
                UseMnemonic = false
            };

            ApplyStatusStyle(label, tone);
            return label;
        }

        public static ComboBox CreateComboBox(Point location, Size size)
        {
            return new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Microsoft YaHei UI", 9.75F),
                ForeColor = PrimaryTextColor,
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Location = location,
                Size = size
            };
        }

        public static CheckBox CreateCheckBox(string text, Point location, int width = 160)
        {
            return new CheckBox
            {
                Text = text,
                Font = new Font("Microsoft YaHei UI", 9.5F),
                ForeColor = PrimaryTextColor,
                BackColor = Color.Transparent,
                Location = location,
                Size = new Size(width, 26)
            };
        }

        public static TextBox CreateReadonlyTextBox(Point location, Size size, string text = "")
        {
            return new TextBox
            {
                BackColor = MutedSurfaceColor,
                ForeColor = PrimaryTextColor,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Microsoft YaHei UI", 9.25F),
                Location = location,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Size = size,
                Text = text,
                TabStop = false
            };
        }

        public static Button CreatePrimaryButton(string text, Point location, Size size, int cornerRadius = 7)
        {
            var button = new SmoothButton
            {
                Text = text,
                Location = location,
                Size = size,
                ForeColor = Color.White,
                Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold),
                UseMnemonic = false,
                NormalBackColor = PrimaryColor,
                HoverBackColor = PrimaryHoverColor,
                PressedBackColor = PrimaryPressedColor,
                BorderColor = PrimaryColor,
                CornerRadius = cornerRadius
            };

            return button;
        }

        public static Button CreateSecondaryButton(string text, Point location, Size size, int cornerRadius = 7)
        {
            var button = new SmoothButton
            {
                Text = text,
                Location = location,
                Size = size,
                ForeColor = PrimaryTextColor,
                Font = new Font("Microsoft YaHei UI", 9.25F, FontStyle.Bold),
                UseMnemonic = false,
                NormalBackColor = Color.White,
                HoverBackColor = Color.FromArgb(248, 251, 255),
                PressedBackColor = Color.FromArgb(240, 246, 253),
                BorderColor = ControlBorderColor,
                CornerRadius = cornerRadius
            };

            return button;
        }

        public static Button CreateDangerButton(string text, Point location, Size size)
        {
            var button = new Button
            {
                Text = text,
                Location = location,
                Size = size,
                BackColor = DangerColor,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei UI", 9.75F, FontStyle.Bold),
                FlatAppearance =
                {
                    BorderSize = 0,
                    MouseOverBackColor = Color.FromArgb(154, 25, 25),
                    MouseDownBackColor = Color.FromArgb(128, 20, 20)
                },
                UseVisualStyleBackColor = false,
                TabStop = false,
                UseMnemonic = false
            };

            ApplyRoundedRegion(button, 10);
            return button;
        }

        public static Button CreateTextButton(string text, Point location, Size size, int cornerRadius = 7)
        {
            var button = new SmoothButton
            {
                Text = text,
                Location = location,
                Size = size,
                ForeColor = PrimaryColor,
                Font = new Font("Microsoft YaHei UI", 9.25F, FontStyle.Bold),
                UseMnemonic = false,
                NormalBackColor = Color.Transparent,
                HoverBackColor = Color.FromArgb(239, 244, 252),
                PressedBackColor = Color.FromArgb(231, 238, 250),
                BorderColor = Color.Transparent,
                CornerRadius = cornerRadius
            };

            return button;
        }

        public static void StyleInputTextBox(TextBox textBox)
        {
            textBox.BorderStyle = BorderStyle.FixedSingle;
            textBox.BackColor = Color.White;
            textBox.ForeColor = PrimaryTextColor;
            textBox.Font = new Font("Microsoft YaHei UI", 9.5F);
        }

        public static void StyleCheckedListBox(CheckedListBox listBox)
        {
            listBox.BorderStyle = BorderStyle.None;
            listBox.BackColor = Color.White;
            listBox.ForeColor = PrimaryTextColor;
            listBox.Font = new Font("Microsoft YaHei UI", 9.5F);
            listBox.ItemHeight = 24;
            listBox.IntegralHeight = false;
        }

        public static void ApplyStatusStyle(Label label, WorkbenchTone tone)
        {
            label.Font = new Font("Microsoft YaHei UI", 7.5F, FontStyle.Bold);
            label.Padding = new Padding(6, 1, 6, 1);
            label.TextAlign = ContentAlignment.MiddleCenter;
            label.BorderStyle = BorderStyle.None;

            switch (tone)
            {
                case WorkbenchTone.Info:
                    label.BackColor = Color.FromArgb(245, 249, 255);
                    label.ForeColor = Color.FromArgb(87, 120, 191);
                    break;
                case WorkbenchTone.Success:
                    label.BackColor = Color.FromArgb(244, 250, 246);
                    label.ForeColor = SuccessColor;
                    break;
                case WorkbenchTone.Warning:
                    label.BackColor = Color.FromArgb(255, 249, 239);
                    label.ForeColor = WarningColor;
                    break;
                case WorkbenchTone.Danger:
                    label.BackColor = Color.FromArgb(253, 243, 243);
                    label.ForeColor = DangerColor;
                    break;
                default:
                    label.BackColor = Color.FromArgb(247, 249, 252);
                    label.ForeColor = TertiaryTextColor;
                    break;
            }

            ApplyRoundedRegion(label, 6);
        }

        public static Panel CreateActionCard(
            string title,
            string description,
            Point location,
            out Button actionButton,
            bool primary,
            string buttonText,
            Size? size = null)
        {
            var panel = CreateCard(location, size ?? new Size(308, 168));

            var lblTitle = new Label
            {
                Text = title,
                Font = new Font("Microsoft YaHei UI", 11F, FontStyle.Bold),
                ForeColor = PrimaryTextColor,
                Location = new Point(18, 16),
                Size = new Size(240, 24),
                UseMnemonic = false,
                BackColor = Color.Transparent
            };

            var lblDescription = new Label
            {
                Text = description,
                Font = new Font("Microsoft YaHei UI", 9.25F),
                ForeColor = SecondaryTextColor,
                Location = new Point(18, 46),
                Size = new Size(272, 54),
                UseMnemonic = false,
                BackColor = Color.Transparent
            };

            actionButton = primary
                ? CreatePrimaryButton(buttonText, new Point(18, 108), new Size(136, 40))
                : CreateSecondaryButton(buttonText, new Point(18, 108), new Size(136, 40));

            panel.Controls.AddRange(new Control[] { lblTitle, lblDescription, actionButton });
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

                using GraphicsPath path = BuildRoundedPath(new Rectangle(0, 0, control.Width, control.Height), radius);
                control.Region = new Region(path);
            }

            control.SizeChanged += (_, _) => UpdateRegion();
            UpdateRegion();
        }

        private static Panel CreateSurfacePanel(Point location, Size size, Color backColor, Color borderColor, int radius)
        {
            var panel = new Panel
            {
                Location = location,
                Size = size,
                BackColor = backColor,
                BorderStyle = BorderStyle.None,
                Margin = Padding.Empty
            };

            void UpdateCardRegion()
            {
                if (panel.Width <= 1 || panel.Height <= 1)
                {
                    return;
                }

                using GraphicsPath path = BuildRoundedPath(new Rectangle(0, 0, panel.Width, panel.Height), radius);
                panel.Region = new Region(path);
            }

            panel.SizeChanged += (_, _) =>
            {
                UpdateCardRegion();
                panel.Invalidate();
            };

            panel.Paint += (_, e) =>
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                using GraphicsPath shadowPath = BuildRoundedPath(new Rectangle(0, 2, panel.Width - 1, panel.Height - 3), radius);
                using var shadowBrush = new SolidBrush(Color.FromArgb(10, 38, 63, 105));
                e.Graphics.FillPath(shadowBrush, shadowPath);
                using GraphicsPath path = BuildRoundedPath(new Rectangle(0, 0, panel.Width - 1, panel.Height - 1), radius);
                using var pen = new Pen(borderColor);
                e.Graphics.DrawPath(pen, path);
            };

            UpdateCardRegion();
            return panel;
        }

        private static GraphicsPath BuildRoundedPath(Rectangle rect, int radius)
        {
            int r = Math.Max(1, radius);
            int d = r * 2;
            var path = new GraphicsPath();

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
    }
}
