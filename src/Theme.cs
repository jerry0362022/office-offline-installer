using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Windows.Forms;

namespace O365Installer
{
    static class Theme
    {
        public static bool Dark;
        static bool loaded;

        public static Color Bg { get { return Dark ? C(0x16, 0x19, 0x1E) : C(0xF5, 0xF6, 0xF8); } }
        public static Color Card { get { return Dark ? C(0x1E, 0x23, 0x2A) : Color.White; } }
        public static Color Border { get { return Dark ? C(0x2B, 0x31, 0x3A) : C(0xE3, 0xE6, 0xEB); } }
        public static Color Text { get { return Dark ? C(0xE6, 0xE9, 0xEE) : C(0x1F, 0x23, 0x28); } }
        public static Color SubText { get { return Dark ? C(0x9A, 0xA4, 0xB2) : C(0x6E, 0x76, 0x81); } }
        public static Color Field { get { return Dark ? C(0x23, 0x29, 0x32) : Color.White; } }
        public static Color Surface { get { return Dark ? C(0x2A, 0x2F, 0x37) : C(0xEE, 0xF1, 0xF5); } }
        public static Color Accent { get { return Dark ? C(0x3B, 0x82, 0xF6) : C(0x2B, 0x6C, 0xF6); } }
        public static Color AccentHover { get { return Dark ? C(0x5B, 0x97, 0xF8) : C(0x1F, 0x5A, 0xE0); } }
        public static Color AccentText { get { return Color.White; } }
        public static Color LogBg { get { return Dark ? C(0x1B, 0x1F, 0x26) : C(0xFB, 0xFC, 0xFD); } }
        public static Color LogText { get { return Dark ? C(0xD7, 0xDB, 0xE0) : C(0x30, 0x37, 0x40); } }
        public static Color Ok { get { return Dark ? C(0x3D, 0xD6, 0x8C) : C(0x16, 0xA3, 0x4A); } }
        public static Color Warn { get { return Dark ? C(0xF5, 0xB5, 0x44) : C(0xB4, 0x53, 0x09); } }
        public static Color Fail { get { return Dark ? C(0xF8, 0x71, 0x71) : C(0xDC, 0x26, 0x26); } }

        static Color C(int r, int g, int b) { return Color.FromArgb(r, g, b); }

        static string FilePath()
        {
            string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeOfflineInstaller");
            try { Directory.CreateDirectory(dir); } catch { }
            return Path.Combine(dir, "theme.txt");
        }

        public static Font UiFont(float size, FontStyle style)
        {
            return new Font("Microsoft YaHei UI", size, style);
        }

        public static void Load()
        {
            if (loaded) return;
            loaded = true;
            try
            {
                string p = FilePath();
                if (!File.Exists(p)) return;
                Dark = File.ReadAllText(p).Trim() == "dark";
            }
            catch { }
        }

        public static void Save()
        {
            try { File.WriteAllText(FilePath(), Dark ? "dark" : "light"); } catch { }
        }

        public static void Toggle()
        {
            Dark = !Dark;
            Save();
        }

        public static void Apply(Control root)
        {
            Form f = root as Form;
            if (f != null)
            {
                f.BackColor = Bg;
                f.ForeColor = Text;
            }

            SideNav sn = root as SideNav;
            if (sn != null)
            {
                sn.BackColor = Card;
                sn.Invalidate();
            }

            FlatProgress fp = root as FlatProgress;
            if (fp != null) fp.Invalidate();

            FlatCheck fc = root as FlatCheck;
            if (fc != null) fc.Invalidate();

            CardPanel cp = root as CardPanel;
            if (cp != null)
            {
                cp.BackColor = Card;
                cp.ForeColor = Text;
                cp.Invalidate();
            }

            CardGroup cg = root as CardGroup;
            if (cg != null)
            {
                cg.BackColor = Card;
                cg.ForeColor = Text;
                cg.Invalidate();
            }

            FlatButton fb = root as FlatButton;
            if (fb != null) fb.ApplyTheme();

            Button b = root as Button;
            if (b != null && fb == null)
            {
                b.FlatStyle = FlatStyle.Flat;
                b.BackColor = Surface;
                b.ForeColor = Text;
                b.FlatAppearance.BorderColor = Border;
                b.FlatAppearance.BorderSize = 1;
            }

            Panel pn = root as Panel;
            if (pn != null && !(pn is CardPanel) && !(pn is SideNav))
            {
                pn.BackColor = Bg;
            }

            TextBox tb = root as TextBox;
            if (tb != null)
            {
                tb.BackColor = tb.ReadOnly ? LogBg : Field;
                tb.ForeColor = tb.ReadOnly ? LogText : Text;
                if (!tb.ReadOnly) tb.BorderStyle = BorderStyle.FixedSingle;
            }
            RichTextBox rt = root as RichTextBox;
            if (rt != null)
            {
                rt.BackColor = LogBg;
                rt.ForeColor = LogText;
                rt.BorderStyle = BorderStyle.None;
            }
            ComboBox cb = root as ComboBox;
            if (cb != null)
            {
                cb.BackColor = Field;
                cb.ForeColor = Text;
                cb.FlatStyle = FlatStyle.Flat;
                cb.Invalidate();
            }
            ListView lv = root as ListView;
            if (lv != null)
            {
                lv.BackColor = Card;
                lv.ForeColor = Text;
            }

            Label lb = root as Label;
            if (lb != null)
            {
                string role = lb.Tag as string;
                if (role == "sub") lb.ForeColor = SubText;
                else if (role == "warn") lb.ForeColor = Warn;
                else if (role == "ok") lb.ForeColor = Ok;
                else if (role == "fail") lb.ForeColor = Fail;
                else if (role != "fixed") lb.ForeColor = Text;
            }

            CheckBox ck = root as CheckBox;
            if (ck != null) ck.ForeColor = Text;
            RadioButton rb = root as RadioButton;
            if (rb != null) rb.ForeColor = Text;

            foreach (Control c in root.Controls) Apply(c);
        }

        public static GraphicsPath Round(Rectangle r, int radius)
        {
            GraphicsPath p = new GraphicsPath();
            int d = radius * 2;
            if (d <= 0) { p.AddRectangle(r); return p; }
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }
    }

    class NavEntry
    {
        public string Glyph;
        public string Text;
        public NavEntry(string g, string t) { Glyph = g; Text = t; }
    }

    class SideNav : Panel
    {
        public readonly List<NavEntry> Items = new List<NavEntry>();
        public int SelectedIndex = 0;
        public event EventHandler SelectedIndexChanged;
        int hover = -1;

        public SideNav()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            Width = 212;
            BackColor = Theme.Card;
        }

        public void Add(string glyph, string text)
        {
            Items.Add(new NavEntry(glyph, text));
            Invalidate();
        }

        public void SetText(int index, string text)
        {
            if (index < 0 || index >= Items.Count) return;
            if (string.IsNullOrEmpty(text) || Items[index].Text == text) return;
            Items[index].Text = text;
            Invalidate();
        }

        int Top0 { get { return 92; } }
        int ItemH { get { return 44; } }

        int HitTest(int y)
        {
            int i = (y - Top0) / ItemH;
            if (y < Top0 || i < 0 || i >= Items.Count) return -1;
            return i;
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            int i = HitTest(e.Y);
            if (i != hover) { hover = i; Invalidate(); Cursor = i >= 0 ? Cursors.Hand : Cursors.Default; }
            base.OnMouseMove(e);
        }

        protected override void OnMouseLeave(EventArgs e) { hover = -1; Invalidate(); base.OnMouseLeave(e); }

        protected override void OnMouseClick(MouseEventArgs e)
        {
            int i = HitTest(e.Y);
            if (i >= 0 && i != SelectedIndex)
            {
                SelectedIndex = i;
                Invalidate();
                if (SelectedIndexChanged != null) SelectedIndexChanged(this, EventArgs.Empty);
            }
            base.OnMouseClick(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(BackColor);

            TextRenderer.DrawText(g, "Office", new Font("Microsoft YaHei UI", 12F, FontStyle.Bold), new Point(20, 26), Theme.Text, TextFormatFlags.NoPadding);
            TextRenderer.DrawText(g, "Offline Installer", new Font("Microsoft YaHei UI", 8F), new Point(21, 50), Theme.SubText, TextFormatFlags.NoPadding);

            Font icon = new Font("Segoe MDL2 Assets", 11F);
            Font text = new Font("Microsoft YaHei UI", 9.5F);

            for (int i = 0; i < Items.Count; i++)
            {
                Rectangle r = new Rectangle(12, Top0 + i * ItemH, Width - 24, 38);
                bool sel = i == SelectedIndex;

                if (sel)
                {
                    using (GraphicsPath p = Theme.Round(r, 8))
                    using (SolidBrush b = new SolidBrush(Theme.Dark ? Color.FromArgb(0x25, 0x2C, 0x38) : Color.FromArgb(0xE8, 0xEF, 0xFE)))
                        g.FillPath(b, p);
                }
                else if (i == hover)
                {
                    using (GraphicsPath p = Theme.Round(r, 8))
                    using (SolidBrush b = new SolidBrush(Theme.Dark ? Color.FromArgb(0x24, 0x29, 0x31) : Color.FromArgb(0xF2, 0xF4, 0xF7)))
                        g.FillPath(b, p);
                }

                Color fg = sel ? Theme.Accent : (Theme.Dark ? Color.FromArgb(0xC3, 0xC9, 0xD2) : Color.FromArgb(0x4A, 0x51, 0x5B));
                TextRenderer.DrawText(g, Items[i].Glyph, icon, new Point(r.X + 14, r.Y + 10), fg);
                TextRenderer.DrawText(g, Items[i].Text, text, new Rectangle(r.X + 44, r.Y, r.Width - 50, r.Height), fg,
                    TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
            }

            using (Pen pen = new Pen(Theme.Border)) g.DrawLine(pen, Width - 1, 0, Width - 1, Height);
        }
    }

    class FlatCheck : CheckBox
    {
        public FlatCheck()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            AutoSize = false;
            Height = 24;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Parent != null ? Parent.BackColor : Theme.Card);

            Rectangle box = new Rectangle(0, (Height - 18) / 2, 18, 18);
            using (GraphicsPath p = Theme.Round(box, 4))
            using (SolidBrush b = new SolidBrush(Checked ? Theme.Accent : (Theme.Dark ? Color.FromArgb(0x23, 0x29, 0x32) : Color.White)))
            using (Pen pen = new Pen(Checked ? Theme.Accent : Theme.Border))
            {
                g.FillPath(b, p);
                g.DrawPath(pen, p);
            }

            if (Checked)
            {
                using (Pen pen = new Pen(Theme.AccentText, 1.8f))
                {
                    pen.StartCap = LineCap.Round;
                    pen.EndCap = LineCap.Round;
                    g.DrawLines(pen, new Point[] {
                        new Point(box.X + 4, box.Y + 9),
                        new Point(box.X + 8, box.Y + 13),
                        new Point(box.X + 14, box.Y + 5) });
                }
            }

            TextRenderer.DrawText(g, Text, Font, new Rectangle(26, 0, Width - 26, Height), Theme.Text,
                TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        }
    }

    class FlatCombo : ComboBox
    {
        public FlatCombo()
        {
            DrawMode = DrawMode.OwnerDrawFixed;
            FlatStyle = FlatStyle.Flat;
            DropDownStyle = ComboBoxStyle.DropDownList;
        }

        protected override void OnDrawItem(DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            bool sel = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            Color bg = sel ? Theme.Accent : Theme.Field;
            Color fg = sel ? Theme.AccentText : Theme.Text;
            using (SolidBrush b = new SolidBrush(bg)) e.Graphics.FillRectangle(b, e.Bounds);
            string t = Items[e.Index].ToString();
            TextRenderer.DrawText(e.Graphics, t, Font, new Rectangle(e.Bounds.X + 6, e.Bounds.Y, e.Bounds.Width - 8, e.Bounds.Height),
                fg, TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            if (m.Msg != 0x000F) return;

            using (Graphics g = Graphics.FromHwnd(Handle))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle r = new Rectangle(0, 0, Width, Height);
                using (SolidBrush b = new SolidBrush(Theme.Field)) g.FillRectangle(b, r);
                using (Pen pen = new Pen(Theme.Border))
                using (GraphicsPath p = Theme.Round(new Rectangle(0, 0, Width - 1, Height - 1), 4))
                {
                    g.DrawPath(pen, p);
                }

                string t = Text;
                TextRenderer.DrawText(g, t, Font, new Rectangle(8, 2, Width - 30, Height - 4), Theme.Text,
                    TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);

                Point[] arrow = new Point[] {
                    new Point(Width - 17, Height / 2 - 2),
                    new Point(Width - 11, Height / 2 - 2),
                    new Point(Width - 14, Height / 2 + 3) };
                using (SolidBrush ab = new SolidBrush(Theme.SubText)) g.FillPolygon(ab, arrow);
            }
        }
    }

    class FlatProgress : Control
    {
        int val;
        public int Maximum = 100;

        public FlatProgress()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        }

        public int Value
        {
            get { return val; }
            set { val = value; Invalidate(); }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Parent != null ? Parent.BackColor : Theme.Bg);

            int h = 6;
            int top = (Height - h) / 2;
            Rectangle track = new Rectangle(0, top, Width - 1, h);
            using (GraphicsPath p = Theme.Round(track, h / 2))
            using (SolidBrush b = new SolidBrush(Theme.Dark ? Color.FromArgb(0x2B, 0x31, 0x3A) : Color.FromArgb(0xE6, 0xEA, 0xF0)))
            {
                g.FillPath(b, p);
            }

            int w = Maximum <= 0 ? 0 : (int)((long)track.Width * val / Maximum);
            if (w > 2)
            {
                if (w < h) w = h;
                Rectangle fill = new Rectangle(0, top, w, h);
                using (GraphicsPath p = Theme.Round(fill, h / 2))
                using (SolidBrush b = new SolidBrush(Theme.Accent))
                {
                    g.FillPath(b, p);
                }
            }
        }
    }

    class CardPanel : Panel
    {
        public CardPanel()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Theme.Bg);
            Rectangle r = new Rectangle(1, 1, Width - 2, Height - 2);
            using (GraphicsPath p = Theme.Round(r, 8))
            using (SolidBrush b = new SolidBrush(BackColor))
            using (Pen pen = new Pen(Theme.Border))
            {
                g.FillPath(b, p);
                g.DrawPath(pen, p);
            }
        }
    }

    class CardGroup : GroupBox
    {
        public const int HeaderHeight = 40;

        bool collapsed;
        bool allowCollapse;
        int expanded;

        public event EventHandler CollapsedChanged;

        public CardGroup()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        }

        public bool AllowCollapse
        {
            get { return allowCollapse; }
            set
            {
                if (allowCollapse == value) return;
                allowCollapse = value;
                if (!allowCollapse && collapsed) Collapsed = false;
                Invalidate();
            }
        }

        public int ExpandedHeight
        {
            get { return expanded > HeaderHeight ? expanded : Height; }
            set { if (value > HeaderHeight) expanded = value; }
        }

        public bool Collapsed
        {
            get { return collapsed; }
            set
            {
                if (collapsed == value) return;
                if (!collapsed && expanded <= HeaderHeight && Height > HeaderHeight) expanded = Height;
                collapsed = value;
                ApplyCollapse();
                Invalidate();
                if (CollapsedChanged != null) CollapsedChanged(this, EventArgs.Empty);
            }
        }

        void ApplyCollapse()
        {
            foreach (Control c in Controls) c.Visible = !collapsed;
            Height = collapsed ? HeaderHeight : (expanded > HeaderHeight ? expanded : HeaderHeight);
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            if (allowCollapse) Cursor = e.Y <= HeaderHeight ? Cursors.Hand : Cursors.Default;
            base.OnMouseMove(e);
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            if (allowCollapse) Cursor = Cursors.Default;
            base.OnMouseLeave(e);
        }

        protected override void OnMouseClick(MouseEventArgs e)
        {
            if (allowCollapse && e.Button == MouseButtons.Left && e.Y <= HeaderHeight)
            {
                Collapsed = !collapsed;
                return;
            }
            base.OnMouseClick(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Theme.Bg);

            Rectangle r = new Rectangle(0, 0, Width - 1, Height - 1);
            using (GraphicsPath p = Theme.Round(r, 8))
            using (SolidBrush b = new SolidBrush(BackColor))
            using (Pen pen = new Pen(Theme.Border))
            {
                g.FillPath(b, p);
                g.DrawPath(pen, p);
            }

            TextRenderer.DrawText(g, Text, new Font("Microsoft YaHei UI", 9F, FontStyle.Bold),
                new Point(14, 9), Theme.Text, TextFormatFlags.NoPadding);

            if (!allowCollapse) return;

            int ax = Width - 26;
            int ay = 17;
            Point[] arrow;
            if (collapsed)
            {
                arrow = new Point[] {
                    new Point(ax - 3, ay - 5),
                    new Point(ax + 2, ay),
                    new Point(ax - 3, ay + 5) };
            }
            else
            {
                arrow = new Point[] {
                    new Point(ax - 5, ay - 3),
                    new Point(ax, ay + 2),
                    new Point(ax + 5, ay - 3) };
            }
            using (Pen ap = new Pen(Theme.SubText, 1.6f))
            {
                ap.StartCap = LineCap.Round;
                ap.EndCap = LineCap.Round;
                ap.LineJoin = LineJoin.Round;
                g.DrawLines(ap, arrow);
            }
        }
    }

    class FlatButton : Button
    {
        public bool Primary;
        bool hover, down;

        public FlatButton()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            Font = new Font("Microsoft YaHei UI", 9F);
            Cursor = Cursors.Hand;
        }

        public void ApplyTheme()
        {
            BackColor = Parent != null ? Parent.BackColor : Theme.Card;
            ForeColor = Theme.Text;
            Invalidate();
        }

        protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { hover = false; down = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs e) { down = true; Invalidate(); base.OnMouseDown(e); }
        protected override void OnMouseUp(MouseEventArgs e) { down = false; Invalidate(); base.OnMouseUp(e); }
        protected override void OnEnabledChanged(EventArgs e) { Invalidate(); base.OnEnabledChanged(e); }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Parent != null ? Parent.BackColor : Theme.Card);

            Color fill;
            Color text;
            if (!Enabled)
            {
                fill = Theme.Dark ? Color.FromArgb(0x2A, 0x2F, 0x37) : Color.FromArgb(0xEC, 0xEF, 0xF3);
                text = Theme.Dark ? Color.FromArgb(0x6B, 0x74, 0x82) : Color.FromArgb(0xA0, 0xA7, 0xB3);
            }
            else if (Primary)
            {
                fill = hover ? Theme.AccentHover : Theme.Accent;
                if (down) fill = ControlPaint.Dark(fill, 0.08f);
                text = Theme.AccentText;
            }
            else
            {
                fill = Theme.Surface;
                if (hover) fill = ControlPaint.Dark(fill, Theme.Dark ? -0.06f : 0.04f);
                if (down) fill = ControlPaint.Dark(fill, Theme.Dark ? -0.1f : 0.08f);
                text = Theme.Text;
            }

            Rectangle r = new Rectangle(0, 0, Width - 1, Height - 1);
            using (GraphicsPath p = Theme.Round(r, 6))
            using (SolidBrush b = new SolidBrush(fill))
            {
                g.FillPath(b, p);
                if (!Primary && Enabled)
                {
                    using (Pen pen = new Pen(Theme.Border)) g.DrawPath(pen, p);
                }
            }

            TextRenderer.DrawText(g, Text, Font, r, text,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        }
    }
}
