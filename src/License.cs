using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace O365Installer
{

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
            if (vbs == null) throw new Exception("当前 Office 安装目录中没有 OSPP.VBS（该脚本只随部分安装包提供），无法通过它管理授权。可打开 Word 用账号登录激活。");

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
                    throw new Exception("执行超时。");
                }
                return o + e;
            }
        }
    }

    class LicenseForm : Form
    {
        TextBox txtOut, txtKey, txtKms;
        Label lblSummary;
        Button btnRefresh, btnKey, btnKms, btnClose, btnWord, btnIid, btnUnp;
        bool busy;

        public LicenseForm()
        {
            BuildUi();
            RefreshStatus();
        }

        void BuildUi()
        {
            Text = "授权管理（调用微软 OSPP）";
            ClientSize = new Size(600, 460);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            Font = new Font("Microsoft YaHei UI", 9F);
            BackColor = Color.FromArgb(248, 249, 251);

            lblSummary = new Label();
            lblSummary.Text = "正在读取授权状态...";
            lblSummary.Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold);
            lblSummary.Location = new Point(16, 14);
            lblSummary.Size = new Size(568, 24);
            Controls.Add(lblSummary);

            txtOut = new TextBox();
            txtOut.Multiline = true;
            txtOut.ReadOnly = true;
            txtOut.ScrollBars = ScrollBars.Vertical;
            txtOut.Font = new Font("Consolas", 8.5F);
            txtOut.Location = new Point(16, 44);
            txtOut.Size = new Size(568, 190);
            txtOut.BackColor = Color.FromArgb(252, 252, 253);
            Controls.Add(txtOut);

            GroupBox g1 = new GroupBox();
            g1.Text = "产品密钥";
            g1.Location = new Point(16, 246);
            g1.Size = new Size(568, 70);
            Controls.Add(g1);

            txtKey = new TextBox();
            txtKey.Location = new Point(14, 28);
            txtKey.Size = new Size(400, 24);
            txtKey.CharacterCasing = CharacterCasing.Upper;
            g1.Controls.Add(txtKey);

            btnKey = new Button();
            btnKey.Text = "安装密钥";
            btnKey.Location = new Point(426, 27);
            btnKey.Size = new Size(126, 26);
            btnKey.Click += async delegate { await InstallKey(); };
            g1.Controls.Add(btnKey);

            GroupBox g2 = new GroupBox();
            g2.Text = "KMS 主机（仅限使用你自己的授权服务器）";
            g2.Location = new Point(16, 326);
            g2.Size = new Size(568, 70);
            Controls.Add(g2);

            txtKms = new TextBox();
            txtKms.Location = new Point(14, 28);
            txtKms.Size = new Size(400, 24);
            txtKms.Text = "";
            g2.Controls.Add(txtKms);

            btnKms = new Button();
            btnKms.Text = "设置并激活";
            btnKms.Location = new Point(426, 27);
            btnKms.Size = new Size(126, 26);
            btnKms.Click += async delegate { await SetKms(); };
            g2.Controls.Add(btnKms);

            btnRefresh = new Button();
            btnRefresh.Text = "刷新状态";
            btnRefresh.Location = new Point(16, 410);
            btnRefresh.Size = new Size(100, 32);
            btnRefresh.Click += delegate { RefreshStatus(); };
            Controls.Add(btnRefresh);

            btnWord = new Button();
            btnWord.Text = "打开 Word 登录";
            btnWord.Location = new Point(112, 410);
            btnWord.Size = new Size(110, 32);
            btnWord.Click += delegate
            {
                string exe = OfficeLicense.FindWordExe();
                if (exe == null)
                {
                    MessageBox.Show("未找到 WINWORD.EXE。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                try { Process.Start(exe); }
                catch (Exception ex) { MessageBox.Show(ex.Message, "启动失败", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            };
            Controls.Add(btnWord);

            btnIid = new Button();
            btnIid.Text = "安装 ID";
            btnIid.Location = new Point(228, 410);
            btnIid.Size = new Size(100, 32);
            btnIid.Click += async delegate { await ShowInstallId(); };
            Controls.Add(btnIid);

            btnUnp = new Button();
            btnUnp.Text = "卸载密钥";
            btnUnp.Location = new Point(334, 410);
            btnUnp.Size = new Size(100, 32);
            btnUnp.Click += async delegate { await RemoveKey(); };
            Controls.Add(btnUnp);

            btnClose = new Button();
            btnClose.Text = "关闭";
            btnClose.Location = new Point(494, 410);
            btnClose.Size = new Size(90, 32);
            btnClose.DialogResult = DialogResult.OK;
            Controls.Add(btnClose);
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
            ok.Text = "确定";
            ok.Location = new Point(226, 96);
            ok.Size = new Size(84, 32);
            ok.DialogResult = DialogResult.OK;
            d.Controls.Add(ok);

            Button no = new Button();
            no.Text = "取消";
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
                    MessageBox.Show("安装 ID 已复制到剪贴板：\r\n\r\n" + iid + "\r\n\r\n用于电话激活时向微软客服提供。",
                        "安装 ID", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("未从输出中解析出安装 ID，请查看窗口内的原始输出。",
                        "安装 ID", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "获取安装 ID 失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally { SetBusy(false); }
        }

        async Task RemoveKey()
        {
            string tail = PromptInput("卸载密钥", "输入要卸载的产品密钥后 5 位（ospp /unpkey）", "");
            if (tail == null) return;
            tail = tail.Trim().ToUpper();
            if (tail.Length != 5)
            {
                MessageBox.Show("需要 5 位字符，通常是密钥的最后一段。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            SetBusy(true);
            try
            {
                string outp = await Task.Run(delegate { return OfficeLicense.Run("/unpkey:" + tail); });
                txtOut.Text = outp;
                RefreshStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "卸载密钥失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally { SetBusy(false); }
        }

        void RefreshStatus()
        {
            if (OfficeLicense.FindOsppVbs() == null)
            {
                lblSummary.Text = "无法查询：当前 Office 未包含 OSPP.VBS";
                lblSummary.ForeColor = Color.FromArgb(190, 120, 0);
                txtOut.Text =
                    "当前安装的 Office 目录下没有 OSPP.VBS，因此无法通过它查询或管理授权。\r\n\r\n" +
                    "原因：OSPP.VBS 只随部分 Office 安装包提供（批量/LTSC 版通常带有），" +
                    "订阅版或新版打包的镜像里可能不包含。\r\n\r\n" +
                    "你可以这样做：\r\n" +
                    "· 订阅版（Microsoft 365）：打开 Word 或 Excel，用已购买订阅的账号登录即可激活\r\n" +
                    "· 零售永久版：打开任一 Office 应用，按提示输入 25 位产品密钥\r\n" +
                    "· 企业批量授权：联系管理员获取 KMS 地址或 MAK 密钥\r\n\r\n" +
                    "点下方「打开 Word」可直接启动应用去登录。";
                txtKey.Enabled = false;
                txtKms.Enabled = false;
                btnKey.Enabled = false;
                btnKms.Enabled = false;
                btnIid.Enabled = false;
                btnUnp.Enabled = false;
                btnWord.Enabled = OfficeLicense.FindWordExe() != null;
                return;
            }

            txtKey.Enabled = true;
            txtKms.Enabled = true;
            btnKey.Enabled = !busy;
            btnKms.Enabled = !busy;
            btnIid.Enabled = !busy;
            btnUnp.Enabled = !busy;
            btnWord.Enabled = true;

            string outp;
            try { outp = OfficeLicense.Run("/dstatus"); }
            catch (Exception ex)
            {
                lblSummary.Text = "读取失败";
                lblSummary.ForeColor = Color.FromArgb(190, 40, 40);
                txtOut.Text = ex.Message;
                return;
            }

            txtOut.Text = outp;

            string upper = outp.ToUpper();
            if (upper.Contains("---LICENSED---"))
            {
                lblSummary.Text = "状态：已激活";
                lblSummary.ForeColor = Color.FromArgb(0, 130, 60);
            }
            else if (upper.Contains("---OOB_GRACE---") || upper.Contains("---OOT_GRACE---"))
            {
                lblSummary.Text = "状态：未激活（处于宽限期）";
                lblSummary.ForeColor = Color.FromArgb(190, 120, 0);
            }
            else if (upper.Contains("---UNLICENSED---") || upper.Contains("---NOTIFICATIONS---"))
            {
                lblSummary.Text = "状态：未激活";
                lblSummary.ForeColor = Color.FromArgb(190, 40, 40);
            }
            else if (outp.IndexOf("No installed product keys", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                lblSummary.Text = "状态：未安装任何密钥";
                lblSummary.ForeColor = Color.FromArgb(190, 40, 40);
            }
            else
            {
                lblSummary.Text = "状态：未识别（见下方原文）";
                lblSummary.ForeColor = Color.FromArgb(80, 85, 95);
            }
        }

        async Task InstallKey()
        {
            string key = txtKey.Text.Trim().Replace(" ", "");
            if (key.Length < 20)
            {
                MessageBox.Show("请输入完整的 25 位产品密钥。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                RefreshStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "安装密钥失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally { SetBusy(false); }
        }

        async Task SetKms()
        {
            string host = txtKms.Text.Trim();
            if (host.Length == 0)
            {
                MessageBox.Show("请填写 KMS 主机地址（例如 kms.contoso.com）。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            SetBusy(true);
            try
            {
                string r = await Task.Run(delegate { return OfficeLicense.Run("/sethst:" + host); });
                txtOut.Text = r;
                string act = await Task.Run(delegate { return OfficeLicense.Run("/act"); });
                txtOut.AppendText(Environment.NewLine + act);
                RefreshStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "设置 KMS 失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally { SetBusy(false); }
        }
    }
}
