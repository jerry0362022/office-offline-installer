using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

namespace O365Installer
{
    class LicEntry
    {
        public string Name = "";
        public string Status = "";
        public string Grace = "";
        public string Key = "";
    }

    static class OfficeLicense
    {
        static string cachedOspp;

        public static string FindOsppVbs()
        {
            if (!string.IsNullOrEmpty(cachedOspp) && File.Exists(cachedOspp)) return cachedOspp;

            string[] candidates = new string[]
            {
                @"C:\Program Files\Microsoft Office\Office16\OSPP.VBS",
                @"C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS",
                @"C:\Program Files\Microsoft Office\root\Office16\OSPP.VBS",
                @"C:\Program Files (x86)\Microsoft Office\root\Office16\OSPP.VBS",
                @"C:\Program Files\Microsoft Office\Office15\OSPP.VBS",
                @"C:\Program Files (x86)\Microsoft Office\Office15\OSPP.VBS"
            };
            foreach (string c in candidates)
            {
                if (File.Exists(c)) { cachedOspp = c; return c; }
            }

            string[] roots = new string[]
            {
                @"C:\Program Files\Microsoft Office",
                @"C:\Program Files (x86)\Microsoft Office"
            };
            foreach (string r in roots)
            {
                try
                {
                    if (!Directory.Exists(r)) continue;
                    string[] hits = Directory.GetFiles(r, "OSPP.VBS", SearchOption.AllDirectories);
                    if (hits.Length > 0) { cachedOspp = hits[0]; return hits[0]; }
                }
                catch { }
            }
            return null;
        }

        public static string FindWordExe()
        {
            string[] candidates = new string[]
            {
                @"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
                @"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
                @"C:\Program Files\Microsoft Office\Office16\WINWORD.EXE",
                @"C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE"
            };
            foreach (string c in candidates) if (File.Exists(c)) return c;
            try
            {
                using (RegistryKey k = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"))
                {
                    if (k != null)
                    {
                        string ip = k.GetValue("InstallPath") as string;
                        if (!string.IsNullOrEmpty(ip))
                        {
                            string p = Path.Combine(ip, "WINWORD.EXE");
                            if (File.Exists(p)) return p;
                        }
                    }
                }
            }
            catch { }
            return null;
        }

        public static string Run(string args)
        {
            string vbs = FindOsppVbs();
            if (vbs == null) throw new Exception(L.T("当前 Office 安装目录中没有 OSPP.VBS（该脚本只随部分安装包提供），无法通过它管理授权。可打开 Word 用账号登录激活。"));

            ProcessStartInfo psi = new ProcessStartInfo("cscript.exe");
            psi.Arguments = "//nologo \"" + vbs + "\" " + args;
            psi.UseShellExecute = false;
            psi.RedirectStandardOutput = true;
            psi.RedirectStandardError = true;
            psi.CreateNoWindow = true;
            try { psi.StandardOutputEncoding = Encoding.GetEncoding(936); } catch { }

            using (Process p = Process.Start(psi))
            {
                string o = p.StandardOutput.ReadToEnd();
                string e = p.StandardError.ReadToEnd();
                if (!p.WaitForExit(120000))
                {
                    try { p.Kill(); } catch { }
                    throw new Exception(L.T("执行超时。"));
                }
                return o + e;
            }
        }

        public static string ReadReg(string path, string name)
        {
            try
            {
                using (RegistryKey k = Registry.LocalMachine.OpenSubKey(path))
                {
                    if (k == null) return null;
                    object v = k.GetValue(name);
                    return v == null ? null : v.ToString();
                }
            }
            catch { return null; }
        }

        public static string FindKmsHost()
        {
            string[] paths = new string[]
            {
                @"SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform",
                @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
            };
            foreach (string p in paths)
            {
                string h = ReadReg(p, "KeyManagementServiceName");
                if (!string.IsNullOrEmpty(h) && h.Trim().Length > 0) return h.Trim();
            }
            return null;
        }

        public static bool CanConnect(string host, int port, int ms)
        {
            try
            {
                using (System.Net.Sockets.TcpClient c = new System.Net.Sockets.TcpClient())
                {
                    IAsyncResult ar = c.BeginConnect(host, port, null, null);
                    if (!ar.AsyncWaitHandle.WaitOne(ms)) return false;
                    c.EndConnect(ar);
                    return true;
                }
            }
            catch { return false; }
        }

        public static string OfficeInfo(string key)
        {
            string v = ReadReg(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", key);
            return string.IsNullOrEmpty(v) ? L.T("（未读取到）") : v;
        }

        public static List<LicEntry> Parse(string outp)
        {
            List<LicEntry> list = new List<LicEntry>();
            if (string.IsNullOrEmpty(outp)) return list;

            string[] lines = outp.Replace("\r\n", "\n").Split('\n');
            LicEntry cur = null;
            foreach (string raw in lines)
            {
                string line = raw.Trim();
                if (line.StartsWith("LICENSE NAME:", StringComparison.OrdinalIgnoreCase))
                {
                    if (cur != null) list.Add(cur);
                    cur = new LicEntry();
                    cur.Name = line.Substring("LICENSE NAME:".Length).Trim();
                }
                else if (cur == null) continue;
                else if (line.StartsWith("LICENSE STATUS:", StringComparison.OrdinalIgnoreCase))
                    cur.Status = line.Substring("LICENSE STATUS:".Length).Trim();
                else if (line.StartsWith("REMAINING GRACE:", StringComparison.OrdinalIgnoreCase))
                    cur.Grace = line.Substring("REMAINING GRACE:".Length).Trim();
                else if (line.StartsWith("Last 5 characters", StringComparison.OrdinalIgnoreCase))
                {
                    int i = line.LastIndexOf(':');
                    if (i > 0) cur.Key = line.Substring(i + 1).Trim();
                }
            }
            if (cur != null) list.Add(cur);
            return list;
        }

        public static string FriendlyName(string raw)
        {
            string s = raw;
            int comma = s.IndexOf(',');
            if (comma > 0) s = s.Substring(comma + 1).Trim();
            if (s.EndsWith(" edition", StringComparison.OrdinalIgnoreCase)) s = s.Substring(0, s.Length - 8).Trim();

            string k = s.ToLower();
            if (k.StartsWith("office16")) k = k.Substring(8);
            else if (k.StartsWith("office15")) k = k.Substring(8);

            if (k == "o365proplusr_subscription2") return L.T("Microsoft 365 企业应用版");
            if (k == "o365businessr_subscription2") return L.T("Microsoft 365 商业版");
            if (k == "o365homepremr_subscription2") return L.T("Microsoft 365 家庭版");
            if (k == "proplusvl_kms_client") return L.T("Office 专业增强版（KMS）");
            if (k == "standardvl_kms_client") return L.T("Office 标准版（KMS）");
            if (k == "proplusr_grace") return L.T("Office 专业增强版（零售）");
            if (k == "visioprovl_kms_client") return L.T("Visio 专业版（KMS）");
            if (k == "visiostdvl_kms_client") return L.T("Visio 标准版（KMS）");
            if (k == "projectprovol_kms_client") return L.T("Project 专业版（KMS）");
            if (k == "projectstdvol_kms_client") return L.T("Project 标准版（KMS）");
            return s;
        }

        public static string GraceText(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return "-";
            string s = raw;
            int p = s.IndexOf('(');
            if (p > 0) s = s.Substring(0, p);
            s = s.Trim();
            s = s.Replace("days", L.T("天")).Replace("day", L.T("天")).Replace("hours", L.T("小时")).Replace("hour", L.T("小时")).Replace("minutes", L.T("分钟"));
            return s.Trim();
        }

        public static string StatusText(string raw)
        {
            string u = (raw ?? "").ToUpper();
            if (u.Contains("---LICENSED---")) return L.T("已激活");
            if (u.Contains("---OOB_GRACE---") || u.Contains("---OOT_GRACE---")) return L.T("宽限期");
            if (u.Contains("---NOTIFICATIONS---")) return L.T("通知模式");
            if (u.Contains("---UNLICENSED---")) return L.T("未激活");
            if (u.Length == 0) return L.T("未知");
            return raw.Trim('-');
        }

        public static Color StatusColor(string raw)
        {
            string u = (raw ?? "").ToUpper();
            if (u.Contains("---LICENSED---")) return Color.FromArgb(0, 130, 60);
            if (u.Contains("---OOB_GRACE---") || u.Contains("---OOT_GRACE---")) return Color.FromArgb(190, 120, 0);
            if (u.Contains("---UNLICENSED---") || u.Contains("---NOTIFICATIONS---")) return Color.FromArgb(190, 40, 40);
            return Color.FromArgb(80, 85, 95);
        }
    }

    class LicenseForm : Form
    {
        TextBox txtOut, txtKey, txtKms;
        Label lblSummary, lblKms;
        ListView lv;
        Button btnRefresh, btnKey, btnKms, btnClose, btnWord, btnIid, btnUnp, btnCkms, btnExport;
        bool busy;

        public LicenseForm()
        {
            BuildUi();
            Shown += async delegate
            {
                await RefreshStatus();
                if (UiTestShot) ShotAndClose();
            };
        }

        internal static bool UiTestShot;

        void ShotAndClose()
        {
            System.Windows.Forms.Timer t = new System.Windows.Forms.Timer();
            t.Interval = 500;
            t.Tick += delegate
            {
                t.Stop();
                try
                {
                    Rectangle r = new Rectangle(0, 0, ClientSize.Width, ClientSize.Height);
                    Bitmap b = new Bitmap(r.Width, r.Height);
                    DrawToBitmap(b, r);
                    b.Save(Path.Combine(Path.GetTempPath(), "o365lic_ui.png"));
                    b.Dispose();
                }
                catch { }
                Close();
            };
            t.Start();
        }

        void BuildUi()
        {
            Text = L.T("授权管理（调用微软 OSPP）");
            ClientSize = new Size(640, 600);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            Font = new Font("Microsoft YaHei UI", 9F);
            BackColor = Theme.Bg;

            lblSummary = new Label();
            lblSummary.Text = L.T("正在读取授权状态...");
            lblSummary.Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold);
            lblSummary.Location = new Point(16, 12);
            lblSummary.Size = new Size(608, 24);
            Controls.Add(lblSummary);

            lblKms = new Label();
            lblKms.Text = "";
            lblKms.Font = new Font("Microsoft YaHei UI", 8F);
            lblKms.ForeColor = Color.FromArgb(120, 125, 135);
            lblKms.Location = new Point(16, 36);
            lblKms.Size = new Size(608, 18);
            lblKms.AutoEllipsis = true;
            Controls.Add(lblKms);

            lv = new ListView();
            lv.View = View.Details;
            lv.FullRowSelect = true;
            lv.GridLines = true;
            lv.MultiSelect = false;
            lv.HeaderStyle = ColumnHeaderStyle.Nonclickable;
            lv.Location = new Point(16, 58);
            lv.Size = new Size(608, 126);
            lv.Columns.Add(L.T("产品"), 250);
            lv.Columns.Add(L.T("状态"), 104);
            lv.Columns.Add(L.T("剩余宽限期"), 108);
            lv.Columns.Add(L.T("密钥尾号"), 92);
            Controls.Add(lv);

            Label lvCap = new Label();
            lvCap.Text = L.T("OSPP 原始输出");
            lvCap.ForeColor = Color.FromArgb(120, 125, 135);
            lvCap.Font = new Font("Microsoft YaHei UI", 8F);
            lvCap.Location = new Point(16, 190);
            lvCap.Size = new Size(200, 18);
            Controls.Add(lvCap);

            txtOut = new TextBox();
            txtOut.Multiline = true;
            txtOut.ReadOnly = true;
            txtOut.ScrollBars = ScrollBars.Vertical;
            txtOut.Font = new Font("Consolas", 8.5F);
            txtOut.Location = new Point(16, 210);
            txtOut.Size = new Size(608, 132);
            txtOut.BackColor = Color.FromArgb(252, 252, 253);
            Controls.Add(txtOut);

            GroupBox g1 = new GroupBox();
            g1.Text = L.T("产品密钥");
            g1.Location = new Point(16, 350);
            g1.Size = new Size(608, 70);
            Controls.Add(g1);

            txtKey = new TextBox();
            txtKey.Location = new Point(14, 28);
            txtKey.Size = new Size(440, 24);
            txtKey.CharacterCasing = CharacterCasing.Upper;
            g1.Controls.Add(txtKey);

            btnKey = new Button();
            btnKey.Text = L.T("安装密钥");
            btnKey.Location = new Point(466, 27);
            btnKey.Size = new Size(126, 26);
            btnKey.Click += async delegate { await InstallKey(); };
            g1.Controls.Add(btnKey);

            GroupBox g2 = new GroupBox();
            g2.Text = L.T("KMS 主机（仅限使用你自己的授权服务器）");
            g2.Location = new Point(16, 428);
            g2.Size = new Size(608, 70);
            Controls.Add(g2);

            txtKms = new TextBox();
            txtKms.Location = new Point(14, 28);
            txtKms.Size = new Size(440, 24);
            txtKms.Text = "";
            g2.Controls.Add(txtKms);

            btnKms = new Button();
            btnKms.Text = L.T("设置并激活");
            btnKms.Location = new Point(466, 27);
            btnKms.Size = new Size(126, 26);
            btnKms.Click += async delegate { await SetKms(); };
            g2.Controls.Add(btnKms);

            btnRefresh = new Button();
            btnRefresh.Text = L.T("刷新状态");
            btnRefresh.Location = new Point(16, 506);
            btnRefresh.Size = new Size(96, 32);
            btnRefresh.Click += async delegate { await RefreshStatus(); };
            Controls.Add(btnRefresh);

            btnWord = new Button();
            btnWord.Text = L.T("打开 Word 登录");
            btnWord.Location = new Point(120, 506);
            btnWord.Size = new Size(116, 32);
            btnWord.Click += delegate
            {
                string exe = OfficeLicense.FindWordExe();
                if (exe == null)
                {
                    MessageBox.Show(L.T("未找到 WINWORD.EXE。"), L.T("提示"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                try { Process.Start(exe); }
                catch (Exception ex) { MessageBox.Show(ex.Message, L.T("启动失败"), MessageBoxButtons.OK, MessageBoxIcon.Error); }
            };
            Controls.Add(btnWord);

            btnIid = new Button();
            btnIid.Text = L.T("安装 ID");
            btnIid.Location = new Point(244, 506);
            btnIid.Size = new Size(104, 32);
            btnIid.Click += async delegate { await ShowInstallId(); };
            Controls.Add(btnIid);

            btnUnp = new Button();
            btnUnp.Text = L.T("卸载密钥");
            btnUnp.Location = new Point(356, 506);
            btnUnp.Size = new Size(96, 32);
            btnUnp.Click += async delegate { await RemoveKey(); };
            Controls.Add(btnUnp);

            btnCkms = new Button();
            btnCkms.Text = L.T("清除 KMS 设置");
            btnCkms.Location = new Point(16, 546);
            btnCkms.Size = new Size(120, 32);
            btnCkms.Click += async delegate { await ClearKms(); };
            Controls.Add(btnCkms);

            btnExport = new Button();
            btnExport.Text = L.T("导出诊断报告");
            btnExport.Location = new Point(144, 546);
            btnExport.Size = new Size(170, 32);
            btnExport.Click += delegate { ExportReport(); };
            Controls.Add(btnExport);

            btnClose = new Button();
            btnClose.Text = L.T("关闭");
            btnClose.Location = new Point(534, 546);
            btnClose.Size = new Size(90, 32);
            btnClose.DialogResult = DialogResult.OK;
            Controls.Add(btnClose);
            Theme.Apply(this);
        }

        void SetBusy(bool b)
        {
            busy = b;
            btnRefresh.Enabled = !b;
            btnKey.Enabled = !b;
            btnKms.Enabled = !b;
            btnClose.Enabled = !b;
            btnIid.Enabled = !b;
            btnUnp.Enabled = !b;
            btnCkms.Enabled = !b;
            btnExport.Enabled = !b;
        }

        static string PromptInput(string title, string label, string def)
        {
            Form d = new Form();
            d.Text = title;
            d.ClientSize = new Size(420, 150);
            d.StartPosition = FormStartPosition.CenterParent;
            d.FormBorderStyle = FormBorderStyle.FixedDialog;
            d.MaximizeBox = false;
            d.MinimizeBox = false;
            d.ShowInTaskbar = false;
            d.Font = new Font("Microsoft YaHei UI", 9F);

            Label t = new Label();
            t.Text = label;
            t.Location = new Point(16, 16);
            t.Size = new Size(388, 20);
            d.Controls.Add(t);

            TextBox tb = new TextBox();
            tb.Location = new Point(16, 42);
            tb.Size = new Size(388, 24);
            tb.CharacterCasing = CharacterCasing.Upper;
            tb.Text = def;
            d.Controls.Add(tb);

            Button ok = new Button();
            ok.Text = L.T("确定");
            ok.Location = new Point(226, 96);
            ok.Size = new Size(84, 32);
            ok.DialogResult = DialogResult.OK;
            d.Controls.Add(ok);

            Button no = new Button();
            no.Text = L.T("取消");
            no.Location = new Point(320, 96);
            no.Size = new Size(84, 32);
            no.DialogResult = DialogResult.Cancel;
            d.Controls.Add(no);

            d.AcceptButton = ok;
            d.CancelButton = no;

            string result = null;
            if (d.ShowDialog() == DialogResult.OK) result = tb.Text.Trim();
            d.Dispose();
            return result;
        }

        async Task ShowInstallId()
        {
            SetBusy(true);
            try
            {
                string outp = await Task.Run(delegate { return OfficeLicense.Run("/dinstid"); });
                txtOut.Text = outp;

                Match m = Regex.Match(outp, @"(\d[\d\s\-]{18,}\d)");
                string iid = m.Success ? m.Groups[1].Value.Replace(" ", "").Replace("-", "").Trim() : null;
                if (!string.IsNullOrEmpty(iid))
                {
                    try { Clipboard.SetText(iid); } catch { }
                    MessageBox.Show(L.T("安装 ID 已复制到剪贴板：\r\n\r\n") + iid + L.T("\r\n\r\n用于电话激活时向微软客服提供。"),
                        L.T("安装 ID"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(L.T("未从输出中解析出安装 ID，请查看窗口内的原始输出。"),
                        L.T("安装 ID"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, L.T("获取安装 ID 失败"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally { SetBusy(false); }
        }

        async Task RemoveKey()
        {
            string tail = PromptInput(L.T("卸载密钥"), L.T("输入要卸载的产品密钥后 5 位（ospp /unpkey）"), "");
            if (tail == null) return;
            tail = tail.Trim().ToUpper();
            if (tail.Length != 5)
            {
                MessageBox.Show(L.T("需要 5 位字符，通常是密钥的最后一段。"), L.T("提示"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            SetBusy(true);
            try
            {
                string outp = await Task.Run(delegate { return OfficeLicense.Run("/unpkey:" + tail); });
                txtOut.Text = outp;
                await RefreshStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, L.T("卸载密钥失败"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally { SetBusy(false); }
        }

        async Task ClearKms()
        {
            DialogResult r = MessageBox.Show(L.T("将清除本机配置的 KMS 主机地址（ospp /ckms），产品密钥不受影响。\r\n\r\n继续？"),
                L.T("清除 KMS 设置"), MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (r != DialogResult.Yes) return;

            SetBusy(true);
            try
            {
                string outp = await Task.Run(delegate { return OfficeLicense.Run("/ckms"); });
                txtOut.Text = outp;
                await RefreshStatus();
                MessageBox.Show(L.T("已清除 KMS 设置。若产品本来靠 KMS 激活，之后需要重新指定可用主机。"),
                    L.T("完成"), MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, L.T("清除失败"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally { SetBusy(false); }
        }

        void ExportReport()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(L.T("Office 离线安装器 · 诊断报告"));
            sb.AppendLine(L.T("生成时间: ") + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            sb.AppendLine(L.T("程序版本: ") + Assembly.GetExecutingAssembly().GetName().Version.ToString());
            sb.AppendLine(L.T("系统: ") + Environment.OSVersion.VersionString + (Environment.Is64BitOperatingSystem ? L.T("（64 位）") : L.T("（32 位）")));
            sb.AppendLine();

            sb.AppendLine(L.T("---- Office 安装信息 ----"));
            sb.AppendLine(L.T("版本(VersionToReport): ") + OfficeLicense.OfficeInfo("VersionToReport"));
            sb.AppendLine(L.T("产品(ProductReleaseIds): ") + OfficeLicense.OfficeInfo("ProductReleaseIds"));
            sb.AppendLine(L.T("位数(Platform): ") + OfficeLicense.OfficeInfo("Platform"));
            sb.AppendLine(L.T("通道(CDNBaseUrl): ") + OfficeLicense.OfficeInfo("CDNBaseUrl"));
            sb.AppendLine(L.T("安装路径: ") + OfficeLicense.OfficeInfo("InstallPath"));
            sb.AppendLine();

            string ospp = OfficeLicense.FindOsppVbs();
            sb.AppendLine(L.T("---- 授权环境 ----"));
            sb.AppendLine("OSPP.VBS: " + (ospp == null ? L.T("未找到") : ospp));
            string kms = OfficeLicense.FindKmsHost();
            sb.AppendLine(L.T("KMS 主机: ") + (kms == null ? L.T("未配置") : kms));
            if (kms != null)
                sb.AppendLine(L.T("KMS 1688 连通性: ") + (OfficeLicense.CanConnect(kms, 1688, 1500) ? L.T("可连接") : L.T("连接失败")));
            sb.AppendLine();

            sb.AppendLine(L.T("---- 授权状态 ----"));
            foreach (LicEntry e in OfficeLicense.Parse(txtOut.Text))
                sb.AppendLine(OfficeLicense.FriendlyName(e.Name) + " | " + OfficeLicense.StatusText(e.Status) + " | " + (e.Grace.Length == 0 ? "-" : e.Grace) + " | " + (e.Key.Length == 0 ? "-" : e.Key));
            sb.AppendLine();

            sb.AppendLine(L.T("---- OSPP 原始输出 ----"));
            sb.AppendLine(txtOut.Text);

            SaveFileDialog sf = new SaveFileDialog();
            sf.Filter = L.T("文本文件 (*.txt)|*.txt");
            sf.FileName = L.T("Office诊断_") + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".txt";
            try { sf.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory); } catch { }
            if (sf.ShowDialog(this) != DialogResult.OK) return;

            try
            {
                File.WriteAllText(sf.FileName, sb.ToString(), new UTF8Encoding(true));
                MessageBox.Show(L.T("已保存到:\r\n") + sf.FileName, L.T("导出成功"), MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, L.T("保存失败"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        async Task RefreshStatus()
        {
            if (OfficeLicense.FindOsppVbs() == null)
            {
                lblSummary.Text = L.T("无法查询：当前 Office 未包含 OSPP.VBS");
                lblSummary.ForeColor = Color.FromArgb(190, 120, 0);
                lblKms.Text = "";
                lv.Items.Clear();
                txtOut.Text =
                    L.T("当前安装的 Office 目录下没有 OSPP.VBS，因此无法通过它查询或管理授权。\r\n\r\n") +
                    L.T("原因：OSPP.VBS 只随部分 Office 安装包提供（批量/LTSC 版通常带有），") +
                    L.T("订阅版或新版打包的镜像里可能不包含。\r\n\r\n") +
                    L.T("你可以这样做：\r\n") +
                    L.T("· 订阅版（Microsoft 365）：打开 Word 或 Excel，用已购买订阅的账号登录即可激活\r\n") +
                    L.T("· 零售永久版：打开任一 Office 应用，按提示输入 25 位产品密钥\r\n") +
                    L.T("· 企业批量授权：联系管理员获取 KMS 地址或 MAK 密钥\r\n\r\n") +
                    L.T("点下方「打开 Word」可直接启动应用去登录。");
                txtKey.Enabled = false;
                txtKms.Enabled = false;
                btnKey.Enabled = false;
                btnKms.Enabled = false;
                btnIid.Enabled = false;
                btnUnp.Enabled = false;
                btnCkms.Enabled = false;
                btnWord.Enabled = OfficeLicense.FindWordExe() != null;
                return;
            }

            txtKey.Enabled = true;
            txtKms.Enabled = true;
            btnKey.Enabled = !busy;
            btnKms.Enabled = !busy;
            btnIid.Enabled = !busy;
            btnUnp.Enabled = !busy;
            btnCkms.Enabled = !busy;
            btnWord.Enabled = true;

            string outp;
            try { outp = await Task.Run(delegate { return OfficeLicense.Run("/dstatus"); }); }
            catch (Exception ex)
            {
                lblSummary.Text = L.T("读取失败");
                lblSummary.ForeColor = Color.FromArgb(190, 40, 40);
                txtOut.Text = ex.Message;
                return;
            }

            txtOut.Text = outp;

            List<LicEntry> entries = OfficeLicense.Parse(outp);
            lv.BeginUpdate();
            lv.Items.Clear();
            int licensed = 0;
            foreach (LicEntry e in entries)
            {
                ListViewItem it = new ListViewItem(OfficeLicense.FriendlyName(e.Name));
                it.SubItems.Add(OfficeLicense.StatusText(e.Status));
                it.SubItems.Add(OfficeLicense.GraceText(e.Grace));
                it.SubItems.Add(e.Key.Length == 0 ? "-" : e.Key);
                it.ForeColor = OfficeLicense.StatusColor(e.Status);
                if (OfficeLicense.StatusText(e.Status) == L.T("已激活")) licensed++;
                lv.Items.Add(it);
            }
            lv.EndUpdate();

            if (entries.Count == 0)
            {
                string upper = outp.ToUpper();
                if (upper.Contains("---LICENSED---")) { lblSummary.Text = L.T("状态：已激活"); lblSummary.ForeColor = Color.FromArgb(0, 130, 60); }
                else if (upper.Contains("---OOB_GRACE---") || upper.Contains("---OOT_GRACE---")) { lblSummary.Text = L.T("状态：未激活（处于宽限期）"); lblSummary.ForeColor = Color.FromArgb(190, 120, 0); }
                else if (upper.Contains("---UNLICENSED---") || upper.Contains("---NOTIFICATIONS---")) { lblSummary.Text = L.T("状态：未激活"); lblSummary.ForeColor = Color.FromArgb(190, 40, 40); }
                else if (outp.IndexOf("No installed product keys", StringComparison.OrdinalIgnoreCase) >= 0) { lblSummary.Text = L.T("状态：未安装任何密钥"); lblSummary.ForeColor = Color.FromArgb(190, 40, 40); }
                else { lblSummary.Text = L.T("状态：未识别（见下方原文）"); lblSummary.ForeColor = Color.FromArgb(80, 85, 95); }
            }
            else if (licensed == entries.Count)
            {
                lblSummary.Text = L.T("状态：已激活（") + licensed + L.T(" 个产品）");
                lblSummary.ForeColor = Color.FromArgb(0, 130, 60);
            }
            else if (licensed == 0)
            {
                lblSummary.Text = L.T("状态：未激活（0 / ") + entries.Count + L.T(" 个产品已激活）");
                lblSummary.ForeColor = Color.FromArgb(190, 40, 40);
            }
            else
            {
                lblSummary.Text = L.T("状态：部分激活（") + licensed + " / " + entries.Count + L.T(" 个产品）");
                lblSummary.ForeColor = Color.FromArgb(190, 120, 0);
            }

            UpdateKmsLine();

            if (outp.IndexOf("0x4004F00C", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                lblSummary.Text = L.T("状态：无法判定 — Office 应用正在运行，请关闭后刷新");
                lblSummary.ForeColor = Color.FromArgb(190, 120, 0);
            }
        }

        void UpdateKmsLine()
        {
            string host = OfficeLicense.FindKmsHost();
            if (host == null)
            {
                lblKms.Text = L.T("KMS 主机：未配置（零售版或订阅版属正常状态）");
                lblKms.ForeColor = Color.FromArgb(120, 125, 135);
                return;
            }

            bool ok = OfficeLicense.CanConnect(host, 1688, 1500);
            if (ok)
            {
                lblKms.Text = L.T("KMS 主机：") + host + L.T("（1688 端口可连接）");
                lblKms.ForeColor = Color.FromArgb(0, 130, 60);
            }
            else
            {
                lblKms.Text = L.T("KMS 主机：") + host + L.T("（1688 端口连接失败，可能是已失效的残留配置）");
                lblKms.ForeColor = Color.FromArgb(190, 120, 0);
                if (txtKms.Text.Length == 0) txtKms.Text = host;
            }
        }

        async Task InstallKey()
        {
            string key = txtKey.Text.Trim().Replace(" ", "");
            if (key.Length < 20)
            {
                MessageBox.Show(L.T("请输入完整的 25 位产品密钥。"), L.T("提示"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            SetBusy(true);
            try
            {
                string r = await Task.Run(delegate { return OfficeLicense.Run("/inpkey:" + key); });
                txtOut.Text = r;
                string act = await Task.Run(delegate { return OfficeLicense.Run("/act"); });
                txtOut.AppendText(Environment.NewLine + act);
                txtKey.Text = "";
                await RefreshStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, L.T("安装密钥失败"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally { SetBusy(false); }
        }

        async Task SetKms()
        {
            string host = txtKms.Text.Trim();
            if (host.Length == 0)
            {
                MessageBox.Show(L.T("请填写 KMS 主机地址（例如 kms.contoso.com）。"), L.T("提示"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            SetBusy(true);
            try
            {
                string r = await Task.Run(delegate { return OfficeLicense.Run("/sethst:" + host); });
                txtOut.Text = r;
                string act = await Task.Run(delegate { return OfficeLicense.Run("/act"); });
                txtOut.AppendText(Environment.NewLine + act);
                await RefreshStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, L.T("设置 KMS 失败"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally { SetBusy(false); }
        }
    }
}
