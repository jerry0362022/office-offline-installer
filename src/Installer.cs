using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyVersion("1.0.1.0")]
[assembly: AssemblyFileVersion("1.0.1.0")]

namespace O365Installer
{
    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
                    | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
            }
            catch { }

            if (args.Length > 0 && args[0] == "--selftest")
            {
                SelfTest();
                return;
            }

            if (args.Length > 0 && args[0] == "--lict")
            {
                StringBuilder lb = new StringBuilder();
                try
                {
                    lb.AppendLine("ospp=" + (OfficeLicense.FindOsppVbs() ?? "NOT FOUND"));
                    lb.AppendLine("--- dstatus ---");
                    lb.AppendLine(OfficeLicense.Run("/dstatus"));
                }
                catch (Exception ex)
                {
                    lb.AppendLine("ERR " + ex.Message);
                }
                File.WriteAllText(Path.Combine(Path.GetTempPath(), "o365_lict.txt"), lb.ToString(), Encoding.UTF8);
                return;
            }

            if (args.Length > 0 && args[0] == "--metrics")
            {
                Application.EnableVisualStyles();
                MainForm mf = new MainForm(true);
                StringBuilder mb = new StringBuilder();
                mb.AppendLine("working=" + Screen.PrimaryScreen.WorkingArea.Width + "x" + Screen.PrimaryScreen.WorkingArea.Height);
                mb.AppendLine("bounds=" + Screen.PrimaryScreen.Bounds.Width + "x" + Screen.PrimaryScreen.Bounds.Height);
                mb.AppendLine("client=" + mf.ClientSize.Width + "x" + mf.ClientSize.Height);
                mb.AppendLine("outer=" + mf.Width + "x" + mf.Height);
                File.WriteAllText(Path.Combine(Path.GetTempPath(), "o365_metrics.txt"), mb.ToString(), Encoding.UTF8);
                return;
            }

            if (args.Length > 0 && args[0] == "--dlgtest")
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                using (UninstallForm uf = new UninstallForm())
                {
                    using (LicenseForm lf = new LicenseForm())
                    {
                        File.WriteAllText(Path.Combine(Path.GetTempPath(), "o365_dlgtest.txt"), "uninstall+license form built ok", Encoding.UTF8);
                    }
                }
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            bool uiTest = args.Length > 0 && args[0] == "--uitest";

            if (!IsAdmin() && !uiTest)
            {
                try
                {
                    ProcessStartInfo psi = new ProcessStartInfo(Assembly.GetExecutingAssembly().Location);
                    psi.UseShellExecute = true;
                    psi.Verb = "runas";
                    psi.WorkingDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                    Process.Start(psi);
                }
                catch
                {
                    MessageBox.Show("需要管理员权限才能安装 Office，请在弹窗中点击\"是\"。", "Microsoft 365 安装器",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                return;
            }

            bool uiTest2 = args.Length > 0 && args[0] == "--uitest";

            bool createdNew = false;
            Mutex single = new Mutex(true, "Global\\O365Installer_SingleRun", out createdNew);
            if (!createdNew && !uiTest2)
            {
                MessageBox.Show("安装器已经在运行，请勿重复打开。", "Microsoft 365 安装器",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (!uiTest2 && !ConsentForm.Ask()) return;

            Application.Run(new MainForm(uiTest2));
            GC.KeepAlive(single);
        }

        internal static bool IsAdmin()
        {
            WindowsIdentity id = WindowsIdentity.GetCurrent();
            WindowsPrincipal p = new WindowsPrincipal(id);
            return p.IsInRole(WindowsBuiltInRole.Administrator);
        }

        static void SelfTest()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("admin=" + IsAdmin());
            sb.AppendLine("64bitOS=" + Environment.Is64BitOperatingSystem);
            sb.AppendLine("temp=" + Path.GetTempPath());
            sb.AppendLine("--- installed office ---");
            List<string> found = MainForm.DetectOffice();
            if (found.Count == 0) sb.AppendLine("none");
            foreach (string f in found) sb.AppendLine(f);
            sb.AppendLine("--- xml ---");
            List<string> excl = new List<string>();
            excl.Add("Outlook"); excl.Add("OneNote");
            sb.AppendLine(MainForm.BuildConfigXml("O365ProPlusRetail", "zh-cn", "64", "E:\\", excl, Path.Combine(Path.GetTempPath(), "O365_Setup", "logs"), "Current"));
            sb.AppendLine("--- remove xml ---");
            sb.AppendLine(OfficeRemover.BuildRemoveC2RXml(Path.Combine(Path.GetTempPath(), "O365_Setup", "logs")));
            sb.AppendLine("--- url ---");
            sb.AppendLine(MainForm.BuildUrl("O365ProPlusRetail", "zh-cn"));
            File.WriteAllText(Path.Combine(Path.GetTempPath(), "o365_exe_selftest.txt"), sb.ToString(), Encoding.UTF8);
        }
    }

    class CorruptImageException : Exception
    {
        public CorruptImageException(string msg) : base(msg) { }
    }

    class MainForm : Form
    {

        const string CdnBase = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/";
        const string DefaultUrl = CdnBase + "zh-cn/O365ProPlusRetail.img";

        static readonly string[] LangCodes = new string[]
        {
            "zh-cn", "zh-tw", "en-us",
            "ar-sa", "bg-bg", "cs-cz", "da-dk", "de-de", "el-gr", "en-gb",
            "es-es", "es-mx", "et-ee", "fi-fi", "fr-ca", "fr-fr", "he-il",
            "hi-in", "hr-hr", "hu-hu", "id-id", "it-it", "ja-jp", "kk-kz",
            "ko-kr", "lt-lt", "lv-lv", "ms-my", "nb-no", "nl-nl", "pl-pl",
            "pt-br", "pt-pt", "ro-ro", "ru-ru", "sk-sk", "sl-si", "sr-latn-rs",
            "sv-se", "th-th", "tr-tr", "uk-ua", "vi-vn"
        };

        internal static string BuildUrl(string product, string lang)
        {
            return CdnBase + lang + "/" + product + ".img";
        }
        static readonly string WorkDir = Path.Combine(Path.GetTempPath(), "O365_Setup");
        static readonly string LogDir = Path.Combine(WorkDir, "logs");
        static readonly string[] AppNames = new string[] { "Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Teams", "OneDrive", "Access", "Publisher" };


        static readonly Dictionary<string, string> LangNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "zh-cn", "简体中文" },
            { "zh-tw", "繁體中文" },
            { "en-us", "英语（美国）" },
            { "en-gb", "英语（英国）" },
            { "ja-jp", "日语" },
            { "ko-kr", "韩语" },
            { "de-de", "德语" },
            { "fr-fr", "法语" },
            { "fr-ca", "法语（加拿大）" },
            { "es-es", "西班牙语" },
            { "es-mx", "西班牙语（墨西哥）" },
            { "ru-ru", "俄语" },
            { "pt-br", "葡萄牙语（巴西）" },
            { "pt-pt", "葡萄牙语（葡萄牙）" },
            { "it-it", "意大利语" },
            { "nl-nl", "荷兰语" },
            { "pl-pl", "波兰语" },
            { "tr-tr", "土耳其语" },
            { "ar-sa", "阿拉伯语" },
            { "he-il", "希伯来语" },
            { "th-th", "泰语" },
            { "vi-vn", "越南语" },
            { "id-id", "印尼语" },
            { "ms-my", "马来语" },
            { "hi-in", "印地语" },
            { "cs-cz", "捷克语" },
            { "da-dk", "丹麦语" },
            { "el-gr", "希腊语" },
            { "fi-fi", "芬兰语" },
            { "hu-hu", "匈牙利语" },
            { "nb-no", "挪威语" },
            { "ro-ro", "罗马尼亚语" },
            { "sk-sk", "斯洛伐克语" },
            { "sl-si", "斯洛文尼亚语" },
            { "sv-se", "瑞典语" },
            { "uk-ua", "乌克兰语" },
            { "bg-bg", "保加利亚语" },
            { "hr-hr", "克罗地亚语" },
            { "sr-latn-rs", "塞尔维亚语" },
            { "et-ee", "爱沙尼亚语" },
            { "lv-lv", "拉脱维亚语" },
            { "lt-lt", "立陶宛语" },
            { "kk-kz", "哈萨克语" },
            { "az-latn-az", "阿塞拜疆语" },
            { "uz-latn-uz", "乌兹别克语" },
            { "mk-mk", "马其顿语" },
            { "sq-al", "阿尔巴尼亚语" },
            { "bs-latn-ba", "波斯尼亚语" },
            { "ga-ie", "爱尔兰语" },
            { "mt-mt", "马耳他语" },
            { "is-is", "冰岛语" },
            { "hy-am", "亚美尼亚语" },
            { "ka-ge", "格鲁吉亚语" },
            { "fa-ir", "波斯语" },
            { "prs-af", "达利语（阿富汗）" },
            { "fil-ph", "菲律宾语" },
            { "mi-nz", "毛利语" },
            { "cy-gb", "威尔士语" },
            { "gd-gb", "苏格兰盖尔语" },
            { "eu-es", "巴斯克语" },
            { "ca-es", "加泰罗尼亚语" },
            { "gl-es", "加利西亚语" }
        };

        static string LangLabel(string code)
        {
            string name;
            if (!LangNames.TryGetValue(code, out name))
            {
                try
                {
                    CultureInfo ci = CultureInfo.GetCultureInfo(code);
                    name = ci.DisplayName;
                    if (string.IsNullOrEmpty(name) || string.Equals(name, code, StringComparison.OrdinalIgnoreCase))
                        name = ci.NativeName;
                }
                catch { name = code; }
            }
            if (string.IsNullOrEmpty(name)) name = code;
            return name + "  (" + code + ")";
        }

        static string CodeFromLabel(string label)
        {
            if (string.IsNullOrEmpty(label)) return "zh-cn";
            int p = label.LastIndexOf('(');
            if (p >= 0 && label.EndsWith(")")) return label.Substring(p + 1, label.Length - p - 2).Trim();
            return label.Trim();
        }

        ComboBox cbProduct, cbLang, cbArch;
        TextBox tbUrl;
        CheckBox chkKeep;
        RadioButton rbImg, rbOdt;
        Label lblHint, lblStatus, lblUrl, lblSourcePath;
        ProgressBar bar;
        RichTextBox log;
        Button btnStart, btnCancel, btnLog;
        Dictionary<string, CheckBox> checks = new Dictionary<string, CheckBox>();
        CancellationTokenSource cts;
        bool busy;
        bool applying;
        List<ProductDef> catalog = Catalog.All();

        public MainForm(bool uiTest)
        {
            BuildUi();

            if (uiTest)
            {
                System.Windows.Forms.Timer t = new System.Windows.Forms.Timer();
                t.Interval = 1200;
                t.Tick += delegate
                {
                    try
                    {
                        Rectangle vs = SystemInformation.VirtualScreen;
                        Bitmap b = new Bitmap(vs.Width, vs.Height);
                        using (Graphics g = Graphics.FromImage(b))
                        {
                            g.CopyFromScreen(vs.Location, new Point(0, 0), vs.Size);
                        }
                        b.Save(Path.Combine(Path.GetTempPath(), "o365exe_ui.png"), ImageFormat.Png);
                        b.Dispose();
                    }
                    catch { }
                    t.Stop();
                    Close();
                };
                t.Start();
            }
        }

        void BuildUi()
        {
            Font f = new Font("Microsoft YaHei UI", 9F);
            Text = "Microsoft 365 离线安装器";
            ClientSize = new Size(744, 606);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            Font = f;
            BackColor = Color.FromArgb(248, 249, 251);

            Label title = new Label();
            title.Text = "Microsoft 365 离线安装器";
            title.Font = new Font("Microsoft YaHei UI", 15F, FontStyle.Bold);
            title.Location = new Point(16, 12);
            title.AutoSize = true;
            Controls.Add(title);

            Label sub = new Label();
            sub.Text = "从官方离线镜像安装，全程自动，无需联网下载安装文件";
            sub.ForeColor = Color.FromArgb(120, 125, 135);
            sub.Location = new Point(18, 42);
            sub.AutoSize = true;
            Controls.Add(sub);

            GroupBox g1 = new GroupBox();
            g1.Text = "安装选项";
            g1.Location = new Point(16, 68);
            g1.Size = new Size(336, 168);
            Controls.Add(g1);

            AddLabel(g1, "产品", 14, 30);
            cbProduct = new ComboBox();
            cbProduct.DropDownStyle = ComboBoxStyle.DropDownList;
            cbProduct.Location = new Point(60, 26);
            cbProduct.Size = new Size(258, 24);
            cbProduct.DropDownWidth = 340;
            cbProduct.MaxDropDownItems = 20;
            foreach (ProductDef p in catalog) cbProduct.Items.Add(p.Name);
            cbProduct.SelectedIndex = 0;
            g1.Controls.Add(cbProduct);

            AddLabel(g1, "语言", 14, 66);
            cbLang = new ComboBox();
            cbLang.DropDownStyle = ComboBoxStyle.DropDownList;
            cbLang.Location = new Point(60, 62);
            cbLang.Size = new Size(150, 24);
            cbLang.DropDownWidth = 240;
            g1.Controls.Add(cbLang);
            FillLanguages();

            AddLabel(g1, "位数", 220, 66);
            cbArch = new ComboBox();
            cbArch.DropDownStyle = ComboBoxStyle.DropDownList;
            cbArch.Location = new Point(258, 62);
            cbArch.Size = new Size(60, 24);
            cbArch.Items.Add("64");
            cbArch.Items.Add("32");
            cbArch.SelectedIndex = 0;
            g1.Controls.Add(cbArch);

            lblHint = new Label();
            lblHint.ForeColor = Color.FromArgb(150, 120, 40);
            lblHint.Font = new Font("Microsoft YaHei UI", 8F);
            lblHint.Location = new Point(14, 98);
            lblHint.Size = new Size(306, 44);
            lblHint.Text = HintText();
            g1.Controls.Add(lblHint);

            cbProduct.SelectedIndexChanged += delegate
            {
                lblHint.Text = HintText();
                if (!applying)
                {
                    ToggleMode();
                    AutoFillUrl();
                }
            };
            cbLang.SelectedIndexChanged += delegate { if (!applying) AutoFillUrl(); };

            GroupBox g2 = new GroupBox();
            g2.Text = "安装源";
            g2.Location = new Point(364, 68);
            g2.Size = new Size(364, 168);
            Controls.Add(g2);

            rbImg = new RadioButton();
            rbImg.Text = "离线镜像 (.img)";
            rbImg.Location = new Point(14, 22);
            rbImg.AutoSize = true;
            rbImg.Checked = true;
            g2.Controls.Add(rbImg);

            rbOdt = new RadioButton();
            rbOdt.Text = "ODT 在线下载";
            rbOdt.Location = new Point(168, 22);
            rbOdt.AutoSize = true;
            g2.Controls.Add(rbOdt);

            lblUrl = new Label();
            lblUrl.Text = "镜像链接 (.img)";
            lblUrl.Location = new Point(14, 52);
            lblUrl.AutoSize = true;
            g2.Controls.Add(lblUrl);

            tbUrl = new TextBox();
            tbUrl.Location = new Point(14, 70);
            tbUrl.Size = new Size(336, 24);
            tbUrl.Text = DefaultUrl;
            g2.Controls.Add(tbUrl);

            lblSourcePath = new Label();
            lblSourcePath.Text = "安装源将下载到: " + Path.Combine(WorkDir, "source");
            lblSourcePath.ForeColor = Color.FromArgb(120, 125, 135);
            lblSourcePath.Font = new Font("Microsoft YaHei UI", 8F);
            lblSourcePath.Location = new Point(14, 52);
            lblSourcePath.Size = new Size(336, 44);
            lblSourcePath.Visible = false;
            g2.Controls.Add(lblSourcePath);

            chkKeep = new CheckBox();
            chkKeep.Text = "完成后保留安装源（便于给其他电脑用）";
            chkKeep.Location = new Point(14, 104);
            chkKeep.AutoSize = true;
            g2.Controls.Add(chkKeep);

            Label l5 = new Label();
            l5.Text = "临时文件: " + WorkDir;
            l5.ForeColor = Color.FromArgb(120, 125, 135);
            l5.Font = new Font("Microsoft YaHei UI", 8F);
            l5.Location = new Point(14, 136);
            l5.Size = new Size(336, 18);
            l5.AutoEllipsis = true;
            g2.Controls.Add(l5);

            rbImg.CheckedChanged += delegate { if (!applying) ToggleMode(); };
            rbOdt.CheckedChanged += delegate { if (!applying) ToggleMode(); };
            ToggleMode();

            GroupBox g3 = new GroupBox();
            g3.Text = "要安装的组件（取消勾选 = 不安装）";
            g3.Location = new Point(16, 246);
            g3.Size = new Size(712, 96);
            Controls.Add(g3);

            string[] on = new string[] { "Word", "Excel", "PowerPoint" };
            for (int i = 0; i < AppNames.Length; i++)
            {
                int col = i % 5;
                int row = i / 5;
                CheckBox c = new CheckBox();
                c.Text = AppNames[i];
                c.AutoSize = true;
                c.CheckAlign = ContentAlignment.MiddleLeft;
                c.Location = new Point(14 + col * 140, 26 + row * 34);
                foreach (string o in on) if (o == AppNames[i]) c.Checked = true;
                g3.Controls.Add(c);
                checks[AppNames[i]] = c;
            }

            bar = new ProgressBar();
            bar.Location = new Point(16, 354);
            bar.Size = new Size(712, 14);
            bar.Maximum = 100;
            Controls.Add(bar);

            lblStatus = new Label();
            lblStatus.Text = "就绪";
            lblStatus.Location = new Point(16, 372);
            lblStatus.Size = new Size(712, 20);
            lblStatus.ForeColor = Color.FromArgb(80, 85, 95);
            Controls.Add(lblStatus);

            log = new RichTextBox();
            log.Location = new Point(16, 396);
            log.Size = new Size(712, 148);
            log.ReadOnly = true;
            log.BackColor = Color.FromArgb(252, 252, 253);
            log.BorderStyle = BorderStyle.FixedSingle;
            log.Font = new Font("Consolas", 9F);
            log.ScrollBars = RichTextBoxScrollBars.Vertical;
            Controls.Add(log);

            btnStart = new Button();
            btnStart.Text = "开始安装";
            btnStart.Location = new Point(534, 556);
            btnStart.Size = new Size(104, 34);
            btnStart.BackColor = Color.FromArgb(0, 120, 212);
            btnStart.ForeColor = Color.White;
            btnStart.FlatStyle = FlatStyle.Flat;
            btnStart.FlatAppearance.BorderSize = 0;
            btnStart.Click += OnStartClick;
            Controls.Add(btnStart);

            btnCancel = new Button();
            btnCancel.Text = "取消";
            btnCancel.Location = new Point(646, 556);
            btnCancel.Size = new Size(82, 34);
            btnCancel.Enabled = false;
            btnCancel.Click += delegate
            {
                if (cts != null) cts.Cancel();
                lblStatus.Text = "正在取消...";
            };
            Controls.Add(btnCancel);

            btnLog = new Button();
            btnLog.Text = "打开日志目录";
            btnLog.Location = new Point(16, 556);
            btnLog.Size = new Size(110, 34);
            btnLog.Click += delegate
            {
                string d = Directory.Exists(LogDir) ? LogDir : WorkDir;
                if (Directory.Exists(d)) Process.Start("explorer.exe", "\"" + d + "\"");
            };
            Controls.Add(btnLog);

            Button btnLicense = new Button();
            btnLicense.Text = "授权管理";
            btnLicense.Location = new Point(136, 556);
            btnLicense.Size = new Size(110, 34);
            btnLicense.Click += delegate
            {
                using (LicenseForm lf = new LicenseForm())
                {
                    lf.ShowDialog(this);
                }
            };
            Controls.Add(btnLicense);

            Button btnUninstall = new Button();
            btnUninstall.Text = "卸载旧版";
            btnUninstall.Location = new Point(256, 556);
            btnUninstall.Size = new Size(110, 34);
            btnUninstall.Click += delegate
            {
                using (UninstallForm uf = new UninstallForm())
                {
                    uf.ShowDialog(this);
                }
            };
            Controls.Add(btnUninstall);

            Button btnAbout = new Button();
            btnAbout.Text = "声明";
            btnAbout.Location = new Point(366, 556);
            btnAbout.Size = new Size(80, 34);
            btnAbout.Click += delegate
            {
                using (ConsentForm cf = new ConsentForm(false)) cf.ShowDialog(this);
            };
            Controls.Add(btnAbout);

            int avail = Screen.PrimaryScreen.WorkingArea.Height;
            int over = Height - avail + 6;
            if (over > 0)
            {
                int newLogH = log.Height - over;
                if (newLogH < 72) newLogH = 72;
                int shift = log.Height - newLogH;
                if (shift > 0)
                {
                    log.Height = newLogH;
                    btnStart.Top -= shift;
                    btnCancel.Top -= shift;
                    btnLog.Top -= shift;
                    btnLicense.Top -= shift;
                    btnUninstall.Top -= shift;
                    btnAbout.Top -= shift;
                    ClientSize = new Size(ClientSize.Width, ClientSize.Height - shift);
                }
            }

            FormClosing += delegate(object s, FormClosingEventArgs e)
            {
                if (busy)
                {
                    DialogResult r = MessageBox.Show("安装正在进行，确定要退出吗？", "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (r != DialogResult.Yes) e.Cancel = true;
                    else if (cts != null) cts.Cancel();
                }
            };
        }

        string HintText()
        {
            ProductDef p = SelectedProduct;
            if (p == null) return "";
            return "产品 ID: " + p.Id + "    更新通道: " + p.Channel + Environment.NewLine
                + (p.Img ? "该产品有微软官方离线镜像，可选镜像或 ODT 下载。" : "该产品仅支持 ODT 在线下载。");
        }

        ProductDef SelectedProduct
        {
            get
            {
                if (catalog.Count == 0) return null;
                int i = cbProduct == null ? 0 : cbProduct.SelectedIndex;
                if (i < 0 || i >= catalog.Count) i = 0;
                return catalog[i];
            }
        }

        void ToggleMode()
        {
            if (rbImg == null || rbOdt == null || tbUrl == null) return;
            ProductDef p = SelectedProduct;
            bool canImg = p != null && p.Img;

            applying = true;
            rbImg.Enabled = canImg;
            if (!canImg) rbOdt.Checked = true;
            applying = false;

            bool useImg = rbImg.Checked && canImg;
            lblUrl.Visible = useImg;
            tbUrl.Visible = useImg;
            lblSourcePath.Visible = !useImg;

            bool multi = p != null && !p.SingleApp;
            foreach (CheckBox c in checks.Values) c.Enabled = multi && !busy;

            if (useImg) AutoFillUrl();
        }

        void FillLanguages()
        {
            applying = true;
            cbLang.Items.Clear();
            foreach (string code in LangCodes) cbLang.Items.Add(LangLabel(code));
            cbLang.SelectedIndex = 0;
            applying = false;
        }

        string SelectedLang
        {
            get
            {
                if (cbLang == null) return "zh-cn";
                return CodeFromLabel(cbLang.Text);
            }
        }

        void AutoFillUrl()
        {
            if (cbProduct == null || tbUrl == null) return;
            ProductDef p = SelectedProduct;
            if (p == null || !p.Img) return;
            string url = BuildUrl(p.Id, SelectedLang);
            if (tbUrl.Text != url) tbUrl.Text = url;
        }

        static void AddLabel(Control parent, string text, int x, int y)
        {
            Label l = new Label();
            l.Text = text;
            l.Location = new Point(x, y);
            l.AutoSize = true;
            parent.Controls.Add(l);
        }

        static string PsQuote(string s) { return "'" + s.Replace("'", "''") + "'"; }

        void Report(int pct, string status)
        {
            if (InvokeRequired) { BeginInvoke(new Action<int, string>(Report), pct, status); return; }
            if (pct >= 0) bar.Value = Math.Max(0, Math.Min(100, pct));
            lblStatus.Text = status;
        }

        void ReportText(string status)
        {
            if (InvokeRequired) { BeginInvoke(new Action<string>(ReportText), status); return; }
            lblStatus.Text = status;
        }

        void Log(string msg)
        {
            Log(msg, Color.FromArgb(45, 48, 55));
        }

        void Log(string msg, Color c)
        {
            if (InvokeRequired) { BeginInvoke(new Action<string, Color>(Log), msg, c); return; }
            log.SelectionStart = log.TextLength;
            log.SelectionLength = 0;
            log.SelectionColor = c;
            log.AppendText("[" + DateTime.Now.ToString("HH:mm:ss") + "] " + msg + Environment.NewLine);
            log.SelectionColor = Color.FromArgb(45, 48, 55);
            log.ScrollToCaret();
        }

        void SetBusy(bool b)
        {
            busy = b;
            btnStart.Enabled = !b;
            btnCancel.Enabled = b;
            cbProduct.Enabled = !b;
            cbLang.Enabled = !b;
            cbArch.Enabled = !b;
            tbUrl.Enabled = !b;
            chkKeep.Enabled = !b;
            foreach (CheckBox c in checks.Values) c.Enabled = !b;
            if (!b) bar.Value = 0;
        }

        async void OnStartClick(object sender, EventArgs e)
        {
            string url = tbUrl.Text.Trim();
            if (!url.StartsWith("http"))
            {
                MessageBox.Show("镜像链接不正确。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!checks["Word"].Checked && !checks["Excel"].Checked && !checks["PowerPoint"].Checked)
            {
                DialogResult r = MessageBox.Show("没有勾选任何核心组件，确定继续？", "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (r != DialogResult.Yes) return;
            }

            SetBusy(true);
            log.Clear();
            cts = new CancellationTokenSource();
            try
            {
                await InstallAsync(url, cts.Token);
            }
            catch (OperationCanceledException)
            {
                Log("已取消", Color.FromArgb(190, 120, 0));
                ReportText("已取消");
            }
            catch (Exception ex)
            {
                Log("出错: " + ex.Message, Color.FromArgb(190, 40, 40));
                ReportText("已停止: " + ex.Message);
                MessageBox.Show(ex.Message, "安装未完成", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                SetBusy(false);
                cts = null;
            }
        }
        async Task InstallAsync(string url, CancellationToken ct)
        {
            ProductDef prod = SelectedProduct;
            if (prod == null) throw new Exception("请先选择要安装的产品。");

            string lang = SelectedLang;
            string arch = cbArch.SelectedItem.ToString();
            bool keepSource = chkKeep.Checked;
            bool useImg = rbImg.Checked && prod.Img;

            List<string> exclude = new List<string>();
            if (!prod.SingleApp)
            {
                foreach (string a in AppNames) if (!checks[a].Checked) exclude.Add(a);
            }

            Directory.CreateDirectory(WorkDir);
            Directory.CreateDirectory(LogDir);

            Report(2, "检查系统环境...");
            Log("产品: " + prod.Name + "  (" + prod.Id + ")");
            Log("通道: " + prod.Channel + "   语言: " + lang + "   " + arch + " 位   方式: " + (useImg ? "离线镜像" : "ODT 下载"));
            if (arch == "64" && !Environment.Is64BitOperatingSystem) throw new Exception("当前系统不是 64 位，无法安装 64 位 Office。");

            long needGB = 13;
            if (prod.SingleApp || prod.Id.StartsWith("Visio") || prod.Id.StartsWith("Project")) needGB = 6;
            long freeGB = new DriveInfo("C").AvailableFreeSpace / (1024L * 1024 * 1024);
            Log("C 盘可用空间: " + freeGB + " GB，需要约 " + needGB + " GB");
            if (freeGB < needGB) throw new Exception("C 盘空间不足，至少需要 " + needGB + " GB，当前 " + freeGB + " GB。");

            Report(5, "检测已安装的 Office...");
            List<string> old = DetectOffice();
            if (old.Count > 0)
            {
                foreach (string o in old) Log("发现: " + o, Color.FromArgb(190, 120, 0));
                Log("提示：安装前如需卸载旧版，请点主界面「卸载旧版」按钮单独处理（本流程不会动已有 Office）", Color.FromArgb(0, 120, 212));
            }
            else
            {
                Log("未发现已安装的 Office");
            }

            if (useImg)
                await InstallFromImage(url, prod, lang, arch, exclude, keepSource, ct);
            else
                await InstallFromOdt(prod, lang, arch, exclude, keepSource, ct);
        }

        async Task InstallFromImage(string url, ProductDef prod, string lang, string arch, List<string> exclude, bool keepSource, CancellationToken ct)
        {
            string img = Path.Combine(WorkDir, prod.Id + "-" + lang + "-x" + arch + ".img");
            string meta = img + ".meta";

            for (int pass = 1; pass <= 2; pass++)
            {
                try
                {
                    await InstallFromImageOnce(url, prod, lang, arch, exclude, keepSource, ct);
                    return;
                }
                catch (CorruptImageException ex)
                {
                    Log("镜像校验失败: " + ex.Message, Color.FromArgb(190, 40, 40));
                    try
                    {
                        if (File.Exists(img)) File.Delete(img);
                        if (File.Exists(meta)) File.Delete(meta);
                        Log("已自动删除损坏的镜像", Color.FromArgb(190, 120, 0));
                    }
                    catch
                    {
                        Log("镜像删除失败，请手动删除: " + img, Color.FromArgb(190, 40, 40));
                        throw;
                    }
                    if (pass == 2)
                        throw new Exception("镜像重新下载后仍然校验失败，请检查网络或更换安装源。");
                    Log("开始重新下载镜像...", Color.FromArgb(0, 120, 212));
                }
            }
        }

        async Task InstallFromImageOnce(string url, ProductDef prod, string lang, string arch, List<string> exclude, bool keepSource, CancellationToken ct)
        {

            string img = Path.Combine(WorkDir, prod.Id + "-" + lang + "-x" + arch + ".img");
            string meta = img + ".meta";
            string cfgPath = Path.Combine(WorkDir, "configuration.xml");
            bool mounted = false;
            string drive = null;

            string legacy = Path.Combine(WorkDir, "office-" + prod.Id + ".img");
            if (!File.Exists(img) && File.Exists(legacy) && new FileInfo(legacy).Length > 3L * 1024 * 1024 * 1024)
            {
                try
                {
                    File.Move(legacy, img);
                    File.WriteAllText(meta, new FileInfo(img).Length.ToString(), new UTF8Encoding(false));
                    Log("发现旧命名镜像，已按当前选择（" + lang + "/x" + arch + "）复用；若语言不符请删除后重下", Color.FromArgb(190, 120, 0));
                }
                catch { }
            }

            bool needDl = true;
            if (File.Exists(img))
            {
                long sz = new FileInfo(img).Length;
                long want = 0;
                try { if (File.Exists(meta)) long.TryParse(File.ReadAllText(meta).Trim(), out want); } catch { }
                double szGB = sz / 1024.0 / 1024 / 1024;

                if (want > 0 && sz == want)
                {
                    Log("镜像已完整: " + szGB.ToString("F2") + " GB，跳过下载");
                    needDl = false;
                }
                else if (want > 0 && sz < want)
                {
                    Log("镜像未下载完: " + szGB.ToString("F2") + " / " + (want / 1024.0 / 1024 / 1024).ToString("F2") + " GB，将断点续传");
                }
                else if (want > 0 && sz > want)
                {
                    Log("镜像大小异常（比预期大），删除后重新下载", Color.FromArgb(190, 120, 0));
                    try { File.Delete(img); File.Delete(meta); } catch { }
                }
                else if (sz > 3L * 1024 * 1024 * 1024)
                {
                    Log("镜像已存在: " + szGB.ToString("F2") + " GB（无大小记录），按完整处理");
                    needDl = false;
                }
                else
                {
                    Log("发现未完成的镜像 (" + (sz / 1024.0 / 1024).ToString("F1") + " MB)，将断点续传");
                }
            }

            if (needDl)
            {
                Report(8, "准备下载镜像...");
                Log("开始下载镜像（约 4.6 GB，支持断点续传）");
                string u = url;
                string m = meta;
                await Task.Run(delegate { Download(u, img, m, ct); });

                long finalSize = new FileInfo(img).Length;
                long totalSize = 0;
                try { long.TryParse(File.ReadAllText(meta).Trim(), out totalSize); } catch { }
                if (totalSize > 0 && finalSize != totalSize)
                    throw new Exception("镜像下载不完整（" + (finalSize / 1024.0 / 1024 / 1024).ToString("F2") + " / " + (totalSize / 1024.0 / 1024 / 1024).ToString("F2") + " GB），重新运行可继续下载。");
                if (finalSize < 3L * 1024 * 1024 * 1024)
                    throw new Exception("镜像下载不完整，可重新运行继续下载。");
                Log("下载完成: " + (finalSize / 1024.0 / 1024 / 1024).ToString("F2") + " GB", Color.FromArgb(0, 130, 60));
            }

            Exception pending = null;
            try
            {
                Report(90, "挂载镜像...");
                Log("正在挂载镜像");
                drive = await Task.Run(delegate { return MountImage(img); });
                if (drive == null)
                {
                    long szNow = new FileInfo(img).Length;
                    long wantNow = 0;
                    try { long.TryParse(File.ReadAllText(meta).Trim(), out wantNow); } catch { }
                    if (wantNow > 0 && szNow != wantNow)
                        throw new CorruptImageException("镜像大小与预期不符（" + szNow + " / " + wantNow + " 字节）");
                    throw new Exception("挂载镜像失败，未获得盘符。请确认镜像文件完整、且不在网络盘上。");
                }
                mounted = true;
                Log("镜像已挂载为 " + drive + ": 盘");

                string setupExe = drive + ":\\setup.exe";
                if (!File.Exists(setupExe)) throw new CorruptImageException("镜像内未找到 setup.exe");

                string dataCab = drive + ":\\Office\\Data\\v" + arch + ".cab";
                if (!File.Exists(dataCab))
                    throw new CorruptImageException("镜像内缺少 Office\\Data\\v" + arch + ".cab");

                Report(93, "生成安装配置...");
                string xml = BuildConfigXml(prod.Id, lang, arch, drive + ":\\", exclude, LogDir, prod.Channel);
                File.WriteAllText(cfgPath, xml, new UTF8Encoding(false));
                Log("已生成配置，排除组件: " + (exclude.Count == 0 ? "无" : string.Join(", ", exclude.ToArray())));

                Report(95, "正在安装...");
                Log("调用 setup.exe 安装，安装窗口中的进度即为实际进度");
                int rc = await RunInstallWithRetry(setupExe, "/configure \"" + cfgPath + "\"", drive + ":\\", ct);
                if (rc != 0) throw new Exception("安装未成功，返回码 " + rc + "。请查看日志目录。常见原因：与旧版 Office 冲突、安全软件拦截、空间不足。");
            }
            catch (Exception ex)
            {
                pending = ex;
            }

            if (mounted)
            {
                Report(97, "卸载镜像...");
                await Task.Run(delegate { DismountImage(img); });
                mounted = false;
                await Task.Delay(800);
            }
            if (pending != null) throw pending;

            Report(99, "校验安装结果...");
            string ver = GetInstalledVersion();
            if (!string.IsNullOrEmpty(ver)) Log("安装成功，当前版本 " + ver, Color.FromArgb(0, 130, 60));
            else Log("安装进程成功，但未读出版本号，请手动打开 Word 确认。", Color.FromArgb(190, 120, 0));

            if (!keepSource)
            {
                try
                {
                    File.Delete(img);
                    Log("已删除镜像，释放约 4.6 GB");
                }
                catch
                {
                    Log("镜像删除失败，可手动删除: " + img, Color.FromArgb(190, 120, 0));
                }
            }
            else
            {
                Log("镜像已保留: " + img);
            }

            Report(100, "安装完成");
            MessageBox.Show(prod.Name + " 安装完成。" + Environment.NewLine + "首次打开 Word / Excel 时按提示登录账号即可激活。",
                "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        async Task InstallFromOdt(ProductDef prod, string lang, string arch, List<string> exclude, bool keepSource, CancellationToken ct)
        {
            string odtDir = Path.Combine(WorkDir, "odt");
            string srcDir = Path.Combine(WorkDir, "source");

            Report(8, "准备部署工具...");
            Log("正在获取微软官方 Office 部署工具 (ODT)");
            Action<string> odtLog = new Action<string>(Log);
            string setupExe = await Task.Run(delegate { return OdtTools.Ensure(odtDir, odtLog, ct); });
            Log("部署工具就绪: " + setupExe, Color.FromArgb(0, 130, 60));

            Report(15, "下载安装文件...");
            Log("开始从微软 CDN 下载安装文件到: " + srcDir);
            string dlXml = BuildConfigXml(prod.Id, lang, arch, srcDir, exclude, LogDir, prod.Channel);
            string dlPath = Path.Combine(WorkDir, "download.xml");
            File.WriteAllText(dlPath, dlXml, new UTF8Encoding(false));

            Stopwatch swDl = Stopwatch.StartNew();
            int drc = await Task.Run(delegate { return RunOdtDownload(setupExe, dlPath, srcDir, swDl, ct); });
            if (drc != 0)
            {
                if (drc == 1603) throw new Exception("下载返回 1603：产品 ID 或通道不被支持，请检查所选产品。");
                throw new Exception("下载安装文件失败，返回码 " + drc + "。请查看日志目录。");
            }
            Log("安装文件下载完成，共 " + DirSizeGB(srcDir).ToString("F2") + " GB", Color.FromArgb(0, 130, 60));

            Report(88, "生成安装配置...");
            string instXml = BuildConfigXml(prod.Id, lang, arch, srcDir, exclude, LogDir, prod.Channel);
            string instPath = Path.Combine(WorkDir, "install.xml");
            File.WriteAllText(instPath, instXml, new UTF8Encoding(false));
            Log("已生成配置，排除组件: " + (exclude.Count == 0 ? "无" : string.Join(", ", exclude.ToArray())));

            Report(92, "正在安装...");
            Log("调用 setup.exe 安装，安装窗口中的进度即为实际进度");
            int rc = await RunInstallWithRetry(setupExe, "/configure \"" + instPath + "\"", odtDir, ct);
            if (rc != 0) throw new Exception("安装未成功，返回码 " + rc + "。请查看日志目录。");

            Report(99, "校验安装结果...");
            string ver = GetInstalledVersion();
            if (!string.IsNullOrEmpty(ver)) Log("安装成功，当前版本 " + ver, Color.FromArgb(0, 130, 60));
            else Log("安装进程成功，但未读出版本号，请手动打开 Word 确认。", Color.FromArgb(190, 120, 0));

            if (!keepSource)
            {
                try
                {
                    Directory.Delete(srcDir, true);
                    Log("已删除安装源目录，释放磁盘空间");
                }
                catch
                {
                    Log("安装源删除失败，可手动删除: " + srcDir, Color.FromArgb(190, 120, 0));
                }
            }
            else
            {
                Log("安装源已保留: " + srcDir);
            }

            Report(100, "安装完成");
            MessageBox.Show(prod.Name + " 安装完成。" + Environment.NewLine + "首次打开 Word / Excel 时按提示登录账号即可激活。",
                "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        void WaitForC2RIdle(int timeoutSec, CancellationToken ct)
        {
            DateTime deadline = DateTime.Now.AddSeconds(timeoutSec);
            int calm = 0;
            while (DateTime.Now < deadline)
            {
                if (ct.IsCancellationRequested) ct.ThrowIfCancellationRequested();
                bool busy = false;
                foreach (Process p in Process.GetProcesses())
                {
                    string n = p.ProcessName.ToLower();
                    if (n == "officec2rclient" || n == "setup" || n == "integratedoffice")
                    {
                        busy = true;
                        break;
                    }
                }
                if (busy)
                {
                    calm = 0;
                    Thread.Sleep(2000);
                    continue;
                }
                calm++;
                if (calm >= 2) return;
                Thread.Sleep(2000);
            }
            Log("等待 Office 客户端空闲超时，继续尝试", Color.FromArgb(190, 120, 0));
        }

        async Task<int> RunInstallWithRetry(string setupExe, string args, string workDir, CancellationToken ct)
        {
            int rc = 0;
            for (int attempt = 1; attempt <= 2; attempt++)
            {
                await Task.Run(delegate { WaitForC2RIdle(300, ct); });
                ReportText("正在安装... （第 " + attempt + " 次尝试）");
                Stopwatch sw = Stopwatch.StartNew();
                rc = await Task.Run(delegate { return RunSetup(setupExe, args, workDir, sw, ct); });
                Log("安装进程返回码: " + rc + "（第 " + attempt + " 次）");

                string nowVer = GetInstalledVersion();
                if (!string.IsNullOrEmpty(nowVer))
                {
                    if (rc != 0) Log("进程返回 " + rc + "，但检测到 Office 已安装成功（" + nowVer + "），视为成功", Color.FromArgb(0, 130, 60));
                    return 0;
                }

                if (rc != 17004 || attempt == 2) return rc;
                Log("17004 = 与另一个 Office 操作冲突，等待 30 秒后重试", Color.FromArgb(190, 120, 0));
                await Task.Delay(30000, ct);
            }
            return rc;
        }

        int RunOdtDownload(string setupExe, string cfgPath, string srcDir, Stopwatch sw, CancellationToken ct)
        {
            Directory.CreateDirectory(srcDir);
            ProcessStartInfo psi = new ProcessStartInfo(setupExe);
            psi.Arguments = "/download \"" + cfgPath + "\"";
            psi.UseShellExecute = false;
            psi.CreateNoWindow = true;
            psi.WorkingDirectory = Path.GetDirectoryName(setupExe);

            using (Process p = Process.Start(psi))
            {
                DateTime last = DateTime.Now;
                while (!p.HasExited)
                {
                    if (ct.IsCancellationRequested)
                    {
                        try { p.Kill(); } catch { }
                        ct.ThrowIfCancellationRequested();
                    }
                    Thread.Sleep(800);
                    if ((DateTime.Now - last).TotalSeconds < 5) continue;
                    last = DateTime.Now;
                    double gb = DirSizeGB(srcDir);
                    int pct = 15 + (int)Math.Min(70.0, gb / 5.0 * 70.0);
                    ReportText("下载安装文件 " + gb.ToString("F2") + " GB   已用 " + sw.Elapsed.ToString(@"mm\:ss"));
                    Report(pct, lblStatus.Text);
                }
                return p.ExitCode;
            }
        }

        static double DirSizeGB(string dir)
        {
            try
            {
                if (!Directory.Exists(dir)) return 0;
                long total = 0;
                foreach (string f in Directory.GetFiles(dir, "*", SearchOption.AllDirectories))
                {
                    try { total += new FileInfo(f).Length; } catch { }
                }
                return total / 1024.0 / 1024 / 1024;
            }
            catch { return 0; }
        }

        void Download(string url, string dest, string metaPath, CancellationToken ct)
        {
            long existing = File.Exists(dest) ? new FileInfo(dest).Length : 0;

            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            req.Timeout = 60000;
            req.ReadWriteTimeout = 60000;
            req.UserAgent = "Mozilla/5.0";
            if (existing > 0) req.AddRange(existing);

            HttpWebResponse resp = (HttpWebResponse)req.GetResponse();
            bool append = resp.StatusCode == HttpStatusCode.PartialContent;
            if (!append) existing = 0;
            long total = existing + resp.ContentLength;
            Log("服务器返回 " + (int)resp.StatusCode + "，总大小 " + (total / 1024.0 / 1024 / 1024).ToString("F2") + " GB");

            if (total > 0 && !string.IsNullOrEmpty(metaPath))
            {
                try { File.WriteAllText(metaPath, total.ToString(), new UTF8Encoding(false)); } catch { }
            }

            FileMode mode = append ? FileMode.Append : FileMode.Create;
            using (FileStream fs = new FileStream(dest, mode, FileAccess.Write))
            using (Stream st = resp.GetResponseStream())
            {
                if (append) fs.Seek(existing, SeekOrigin.Begin);
                byte[] buf = new byte[262144];
                long done = existing;
                DateTime last = DateTime.Now;
                while (true)
                {
                    if (ct.IsCancellationRequested) { resp.Close(); ct.ThrowIfCancellationRequested(); }
                    int n = st.Read(buf, 0, buf.Length);
                    if (n <= 0) break;
                    fs.Write(buf, 0, n);
                    done += n;
                    if ((DateTime.Now - last).TotalMilliseconds > 300)
                    {
                        last = DateTime.Now;
                        int pct = 8;
                        if (total > 0) pct = 8 + (int)(80L * done / total);
                        Report(pct, "下载中 " + (done / 1024.0 / 1024 / 1024).ToString("F2") + " / " + (total / 1024.0 / 1024 / 1024).ToString("F2") + " GB");
                    }
                }
            }
            resp.Close();
        }

        int RunSetup(string setupExe, string args, string workDir, Stopwatch sw, CancellationToken ct)
        {
            ProcessStartInfo psi = new ProcessStartInfo(setupExe);
            psi.Arguments = args;
            psi.UseShellExecute = false;
            if (!string.IsNullOrEmpty(workDir)) psi.WorkingDirectory = workDir;
            Process p = Process.Start(psi);
            while (!p.HasExited)
            {
                if (ct.IsCancellationRequested)
                {
                    try { p.Kill(); } catch { }
                    ct.ThrowIfCancellationRequested();
                }
                Thread.Sleep(500);
                ReportText("处理中... 已用时 " + sw.Elapsed.ToString(@"mm\:ss"));
            }
            return p.ExitCode;
        }

        string MountImage(string img)
        {
            string cmd = "$p=" + PsQuote(img) + ";Mount-DiskImage -ImagePath $p -Access ReadOnly -ErrorAction Stop | Out-Null;"
                + "$i=0;while(-not (Get-DiskImage -ImagePath $p | Get-Volume).DriveLetter -and $i -lt 20){Start-Sleep -Milliseconds 500;$i++};"
                + "(Get-DiskImage -ImagePath $p | Get-Volume).DriveLetter";
            string outp = RunPs(cmd, 60000);
            if (outp == null) return null;
            string d = outp.Trim();
            if (d.Length != 1) return null;
            return d;
        }

        void DismountImage(string img)
        {
            try
            {
                RunPs("Dismount-DiskImage -ImagePath " + PsQuote(img) + " -ErrorAction SilentlyContinue | Out-Null", 60000);
            }
            catch { }
        }

        static string RunPs(string command, int timeoutMs)
        {
            ProcessStartInfo psi = new ProcessStartInfo("powershell.exe");
            psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -Command \"" + command + "\"";
            psi.UseShellExecute = false;
            psi.RedirectStandardOutput = true;
            psi.RedirectStandardError = true;
            psi.CreateNoWindow = true;
            using (Process p = Process.Start(psi))
            {
                string outp = p.StandardOutput.ReadToEnd();
                p.StandardError.ReadToEnd();
                if (!p.WaitForExit(timeoutMs))
                {
                    try { p.Kill(); } catch { }
                    return null;
                }
                return outp;
            }
        }

        internal static List<string> DetectOffice()
        {
            List<string> found = new List<string>();
            string[] paths = new string[]
            {
                @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
                @"SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
            };
            foreach (string p in paths)
            {
                try
                {
                    using (RegistryKey k = Registry.LocalMachine.OpenSubKey(p))
                    {
                        if (k == null) continue;
                        object ids = k.GetValue("ProductReleaseIds");
                        object ver = k.GetValue("VersionToReport");
                        if (ids != null)
                            found.Add("即点即用版: " + ids + "  " + (ver == null ? "" : ver.ToString()));
                    }
                }
                catch { }
            }
            string[] vers = new string[] { "16.0", "15.0", "14.0", "12.0" };
            string[] pre = new string[] { "", @"WOW6432Node\" };
            foreach (string v in vers)
            {
                foreach (string pp in pre)
                {
                    try
                    {
                        using (RegistryKey k = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\" + pp + @"Microsoft\Office\" + v + @"\Word\InstallRoot"))
                        {
                            if (k != null) found.Add("MSI 旧版: Office " + v);
                        }
                    }
                    catch { }
                }
            }
            return found;
        }

        internal static string GetInstalledVersion()
        {
            try
            {
                using (RegistryKey k = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"))
                {
                    if (k == null) return "";
                    object v = k.GetValue("VersionToReport");
                    return v == null ? "" : v.ToString();
                }
            }
            catch { return ""; }
        }

        internal static string BuildConfigXml(string product, string lang, string arch, string source, List<string> exclude, string logPath, string channel)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            sb.AppendLine("<Configuration ID=\"o365-gui\">");
            sb.AppendLine("  <Add OfficeClientEdition=\"" + arch + "\" Channel=\"" + channel + "\" SourcePath=\"" + source + "\">");
            sb.AppendLine("    <Product ID=\"" + product + "\">");
            sb.AppendLine("      <Language ID=\"" + lang + "\" />");
            foreach (string e in exclude) sb.AppendLine("      <ExcludeApp ID=\"" + e + "\" />");
            sb.AppendLine("    </Product>");
            sb.AppendLine("  </Add>");
            sb.AppendLine("  <Property Name=\"AUTOACTIVATE\" Value=\"1\" />");
            sb.AppendLine("  <Property Name=\"FORCEAPPSHUTDOWN\" Value=\"TRUE\" />");
            sb.AppendLine("  <Property Name=\"SharedComputerLicensing\" Value=\"0\" />");
            sb.AppendLine("  <Property Name=\"DeviceBasedLicensing\" Value=\"0\" />");
            sb.AppendLine("  <Property Name=\"PinIconsToTaskbar\" Value=\"FALSE\" />");
            sb.AppendLine("  <Updates Enabled=\"TRUE\" Channel=\"" + channel + "\" />");
            sb.AppendLine("  <Display Level=\"Full\" AcceptEULA=\"TRUE\" />");
            sb.AppendLine("  <Logging Level=\"Standard\" Path=\"" + logPath + "\" />");
            sb.AppendLine("</Configuration>");
            return sb.ToString();
        }
    }
}
