using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

namespace O365Installer
{
    static class OfficeRemover
    {

        public static string InstalledVersion()
        {
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
                        object ver = k.GetValue("VersionToReport");
                        if (ver != null) return ver.ToString();
                    }
                }
                catch { }
            }
            return null;
        }

        public static string InstalledProductId()
        {
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
                        if (ids == null) continue;
                        string v = ids.ToString();
                        int comma = v.IndexOf(',');
                        return comma > 0 ? v.Substring(0, comma).Trim() : v.Trim();
                    }
                }
                catch { }
            }
            return null;
        }

        public static string InstalledCulture()
        {
            try
            {
                using (RegistryKey k = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"))
                {
                    if (k == null) return "zh-cn";
                    object c = k.GetValue("ClientCulture");
                    return c == null ? "zh-cn" : c.ToString();
                }
            }
            catch { return "zh-cn"; }
        }

        public static string ClickToRunExe()
        {
            string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles), @"microsoft shared\ClickToRun");
            string exe = Path.Combine(dir, "OfficeClickToRun.exe");
            return File.Exists(exe) ? exe : null;
        }

        public static string ToolVersion(string exe)
        {
            try { return FileVersionInfo.GetVersionInfo(exe).FileVersion; }
            catch { return null; }
        }

        public static int CompareVersion(string a, string b)
        {
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b)) return 0;
            try { return new Version(a).CompareTo(new Version(b)); }
            catch { return 0; }
        }

        public static string GetC2RInfo()
        {
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
                        if (ids == null) continue;
                        object ver = k.GetValue("VersionToReport");
                        return ids.ToString() + "  " + (ver == null ? "" : ver.ToString());
                    }
                }
                catch { }
            }
            return null;
        }

        public static List<string[]> FindMsiOffice()
        {
            List<string[]> list = new List<string[]>();
            string[] roots = new string[]
            {
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            };
            Regex nameRe = new Regex("Office|Word|Excel|PowerPoint|Outlook|OneNote|Access|Publisher|Visio|Project|Skype for Business", RegexOptions.IgnoreCase);

            foreach (string r in roots)
            {
                try
                {
                    using (RegistryKey k = Registry.LocalMachine.OpenSubKey(r))
                    {
                        if (k == null) continue;
                        foreach (string sub in k.GetSubKeyNames())
                        {
                            if (!sub.StartsWith("{")) continue;
                            try
                            {
                                using (RegistryKey p = k.OpenSubKey(sub))
                                {
                                    if (p == null) continue;
                                    string name = p.GetValue("DisplayName") as string;
                                    string pub = p.GetValue("Publisher") as string;
                                    string us = p.GetValue("UninstallString") as string;
                                    if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(pub)) continue;
                                    if (pub.IndexOf("Microsoft", StringComparison.OrdinalIgnoreCase) < 0) continue;
                                    if (!nameRe.IsMatch(name)) continue;
                                    if (us == null || us.IndexOf("msiexec", StringComparison.OrdinalIgnoreCase) < 0) continue;

                                    bool dup = false;
                                    foreach (string[] x in list) if (x[1] == sub) { dup = true; break; }
                                    if (!dup)
                                    {
                                        object v = p.GetValue("DisplayVersion");
                                        list.Add(new string[] { name, sub, v == null ? "" : v.ToString() });
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                }
                catch { }
            }
            return list;
        }

        public static int RemoveMsi(string productCode, CancellationToken ct)
        {
            ProcessStartInfo psi = new ProcessStartInfo("msiexec.exe");
            psi.Arguments = "/x " + productCode + " /qn /norestart";
            psi.UseShellExecute = false;
            psi.CreateNoWindow = true;
            using (Process p = Process.Start(psi))
            {
                while (!p.HasExited)
                {
                    if (ct.IsCancellationRequested)
                    {
                        try { p.Kill(); } catch { }
                        ct.ThrowIfCancellationRequested();
                    }
                    Thread.Sleep(400);
                }
                return p.ExitCode;
            }
        }

        public static string BuildRemoveC2RXml(string logPath)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            sb.AppendLine("<Configuration ID=\"o365-remove\">");
            sb.AppendLine("  <Remove All=\"TRUE\" />");
            sb.AppendLine("  <Property Name=\"FORCEAPPSHUTDOWN\" Value=\"TRUE\" />");
            sb.AppendLine("  <Display Level=\"None\" AcceptEULA=\"TRUE\" />");
            sb.AppendLine("  <Logging Level=\"Standard\" Path=\"" + logPath + "\" />");
            sb.AppendLine("</Configuration>");
            return sb.ToString();
        }
    }

    class UninstallForm : Form
    {
        CheckedListBox lst;
        TextBox txtLog;
        Label lblC2R, lblTip;
        Button btnC2R, btnMsi, btnToggle, btnRefresh, btnClose;
        CancellationTokenSource cts = new CancellationTokenSource();
        bool busy;
        List<string[]> msiItems = new List<string[]>();
        string c2rInfo;

        public UninstallForm()
        {
            BuildUi();
            Theme.Apply(this);
            RefreshList();
        }

        void BuildUi()
        {
            Text = L.T("卸载旧版 Office");
            ClientSize = new Size(600, 470);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            Font = new Font("Microsoft YaHei UI", 9F);
            BackColor = Theme.Bg;

            lblC2R = new Label();
            lblC2R.Text = L.T("即点即用版：检测中...");
            lblC2R.Location = new Point(16, 14);
            lblC2R.AutoSize = true;
            Controls.Add(lblC2R);

            Label t = new Label();
            t.Text = L.T("MSI 版条目（勾选后才会卸载，可多选）：");
            t.Location = new Point(16, 40);
            t.AutoSize = true;
            Controls.Add(t);

            lst = new CheckedListBox();
            lst.Location = new Point(16, 64);
            lst.Size = new Size(568, 120);
            lst.BorderStyle = BorderStyle.FixedSingle;
            lst.CheckOnClick = true;
            lst.HorizontalScrollbar = true;
            Controls.Add(lst);

            lblTip = new Label();
            lblTip.Text = "";
            lblTip.ForeColor = Color.FromArgb(190, 40, 40);
            lblTip.Location = new Point(16, 190);
            lblTip.Size = new Size(568, 32);
            lblTip.Font = new Font("Microsoft YaHei UI", 8F);
            Controls.Add(lblTip);

            btnC2R = new Button();
            btnC2R.Text = L.T("卸载即点即用版");
            btnC2R.Location = new Point(16, 226);
            btnC2R.Size = new Size(150, 34);
            btnC2R.Click += async delegate { await RemoveC2R(); };
            Controls.Add(btnC2R);

            btnMsi = new Button();
            btnMsi.Text = L.T("卸载勾选项");
            btnMsi.Location = new Point(174, 226);
            btnMsi.Size = new Size(140, 34);
            btnMsi.Click += async delegate { await RemoveMsiSelected(); };
            Controls.Add(btnMsi);

            btnToggle = new Button();
            btnToggle.Text = L.T("全选");
            btnToggle.Location = new Point(322, 226);
            btnToggle.Size = new Size(80, 34);
            btnToggle.Click += delegate { ToggleAll(); };
            Controls.Add(btnToggle);

            btnRefresh = new Button();
            btnRefresh.Text = L.T("刷新");
            btnRefresh.Location = new Point(410, 226);
            btnRefresh.Size = new Size(80, 34);
            btnRefresh.Click += delegate { RefreshList(); };
            Controls.Add(btnRefresh);

            btnClose = new Button();
            btnClose.Text = L.T("关闭");
            btnClose.Location = new Point(498, 226);
            btnClose.Size = new Size(86, 34);
            btnClose.DialogResult = DialogResult.OK;
            Controls.Add(btnClose);

            txtLog = new TextBox();
            txtLog.Multiline = true;
            txtLog.ReadOnly = true;
            txtLog.ScrollBars = ScrollBars.Vertical;
            txtLog.Font = new Font("Consolas", 8.5F);
            txtLog.Location = new Point(16, 272);
            txtLog.Size = new Size(568, 182);
            txtLog.BackColor = Color.FromArgb(252, 252, 253);
            Controls.Add(txtLog);
        }

        void Log(string msg)
        {
            if (InvokeRequired) { BeginInvoke(new Action<string>(Log), msg); return; }
            txtLog.AppendText("[" + DateTime.Now.ToString("HH:mm:ss") + "] " + msg + Environment.NewLine);
        }

        void SetBusy(bool b)
        {
            busy = b;
            btnC2R.Enabled = !b;
            btnMsi.Enabled = !b;
            btnToggle.Enabled = !b;
            btnRefresh.Enabled = !b;
            btnClose.Enabled = !b;
            if (b) Cursor = Cursors.WaitCursor; else Cursor = Cursors.Default;
        }

        void RefreshList()
        {
            lst.Items.Clear();
            c2rInfo = OfficeRemover.GetC2RInfo();
            lblC2R.Text = c2rInfo == null
                ? L.T("即点即用版：未检测到")
                : L.T("即点即用版：") + c2rInfo;

            msiItems = OfficeRemover.FindMsiOffice();
            bool sharedRisk = false;
            foreach (string[] it in msiItems)
            {
                string ver = it[2];
                bool shared = c2rInfo != null && ver.StartsWith("16.");
                if (shared) sharedRisk = true;
                lst.Items.Add((shared ? "★ " : "") + it[0] + (string.IsNullOrEmpty(ver) ? "" : "  (" + ver + ")"), false);
            }

            if (msiItems.Count == 0) lst.Items.Add(L.T("（未检测到 MSI 版 Office）"), false);

            lblTip.Text = sharedRisk
                ? L.T("★ 标记的条目与当前即点即用版 Office 共用组件目录，卸载会导致 OSPP.VBS 等功能文件缺失，建议不要卸载。")
                : "";

            btnC2R.Enabled = c2rInfo != null && !busy;
            btnMsi.Enabled = msiItems.Count > 0 && !busy;
            btnToggle.Enabled = msiItems.Count > 0 && !busy;
        }

        void ToggleAll()
        {
            bool all = true;
            for (int i = 0; i < msiItems.Count; i++) if (!lst.GetItemChecked(i)) { all = false; break; }
            for (int i = 0; i < msiItems.Count; i++) lst.SetItemChecked(i, !all);
            btnToggle.Text = all ? L.T("全选") : L.T("清空");
        }

        async Task RemoveC2R()
        {
            if (c2rInfo == null) return;
            DialogResult r = MessageBox.Show(
                L.T("将卸载本机所有即点即用版 Office：\r\n\r\n") + c2rInfo + L.T("\r\n\r\n此操作不可撤销，确定继续？"),
                L.T("确认卸载"), MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (r != DialogResult.Yes) return;

            SetBusy(true);
            try
            {
                string workDir = Path.Combine(Path.GetTempPath(), "O365_Setup");
                Directory.CreateDirectory(workDir);
                Log(L.T("准备部署工具..."));
                string wd = workDir;
                Action<string> act = new Action<string>(Log);
                string setup = await Task.Run(delegate { return OdtTools.Ensure(wd, act, cts.Token); });
                string instVer = OfficeRemover.InstalledVersion();
                string toolVer = OfficeRemover.ToolVersion(setup);
                if (instVer != null && toolVer != null && OfficeRemover.CompareVersion(toolVer, instVer) < 0)
                    Log(L.T("部署工具版本低于已安装 Office（") + toolVer + " < " + instVer + L.T("），卸载可能被拒绝"));

                string cfg = Path.Combine(wd, "remove.xml");
                File.WriteAllText(cfg, OfficeRemover.BuildRemoveC2RXml(wd), new UTF8Encoding(false));

                Log(L.T("开始卸载即点即用版 Office..."));
                string sd = Path.GetDirectoryName(setup);
                int rc = await Task.Run(delegate
                {
                    return OdtTools.RunSetup(setup, "/configure \"" + cfg + "\"", sd, null, cts.Token);
                });
                Log(L.T("卸载进程返回码: ") + rc);
                if (rc == 0)
                {
                    await Task.Delay(1500);
                    c2rInfo = OfficeRemover.GetC2RInfo();
                    if (c2rInfo == null) Log(L.T("已卸载完成"));
                    else
                    {
                        Log(L.T("卸载进程返回 0，但系统中仍检测到 Office"));
                        MessageBox.Show(L.T("卸载进程返回成功，但系统里仍能检测到 Office。请重启后再执行一次；若仍不生效，请查看日志目录里的 setup 日志。"),
                            L.T("提示"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    string fallback = OfficeRemover.ClickToRunExe();
                    string pid = OfficeRemover.InstalledProductId();
                    if (fallback != null && pid != null)
                    {
                        Log(L.T("改用系统内置卸载接口重试..."));
                        string culture = OfficeRemover.InstalledCulture();
                        string args = "scenario=install scenariosubtype=ARP sourcetype=None productstoremove=" + pid + ".16_" + culture + "_x-none culture=" + culture + " version=16.0.0.0";
                        try
                        {
                            using (Process p = Process.Start(fallback, args))
                            {
                                p.WaitForExit(600000);
                                Log(L.T("系统卸载接口返回码: ") + p.ExitCode);
                            }
                            await Task.Delay(3000);
                            c2rInfo = OfficeRemover.GetC2RInfo();
                            if (c2rInfo == null)
                            {
                                Log(L.T("已卸载完成"));
                                RefreshList();
                                SetBusy(false);
                                return;
                            }
                        }
                        catch (Exception ex2) { Log(L.T("失败: ") + ex2.Message); }
                    }

                    rc = 17002;
                    string hint = L.T("部署工具版本低于已安装的 Office，卸载被拒绝。可改用「设置 - 应用」卸载，或保留安装源后用它卸载");
                    Log(L.T("卸载未成功，可查看 %TEMP% 下的 Office 日志") + (hint.Length > 0 ? "；" + hint : ""));
                    MessageBox.Show(L.T("卸载未成功，返回码 ") + rc + (hint.Length > 0 ? "\r\n\r\n" + hint : ""),
                        L.T("卸载未成功"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                RefreshList();
            }
            catch (OperationCanceledException)
            {
                Log(L.T("已取消"));
            }
            catch (Exception ex)
            {
                Log(L.T("失败: ") + ex.Message);
            }
            finally { SetBusy(false); }
        }

        async Task RemoveMsiSelected()
        {
            List<string[]> picked = new List<string[]>();
            for (int i = 0; i < msiItems.Count; i++)
            {
                if (lst.GetItemChecked(i)) picked.Add(msiItems[i]);
            }
            if (picked.Count == 0)
            {
                MessageBox.Show(L.T("请先在列表中勾选要卸载的条目。"), L.T("提示"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            bool hasShared = false;
            foreach (string[] it in picked) if (c2rInfo != null && it[2].StartsWith("16.")) hasShared = true;

            StringBuilder sb = new StringBuilder();
            sb.AppendLine(L.T("将卸载以下 ") + picked.Count + L.T(" 项："));
            sb.AppendLine();
            foreach (string[] it in picked) sb.AppendLine("· " + it[0] + (string.IsNullOrEmpty(it[2]) ? "" : " (" + it[2] + ")"));
            sb.AppendLine();
            if (hasShared)
            {
                sb.AppendLine(L.T("注意：其中包含与当前 Office 共用组件的条目，"));
                sb.AppendLine(L.T("卸载后 OSPP.VBS 等文件会消失（不影响 Word/Excel 使用）。"));
                sb.AppendLine();
            }
            sb.AppendLine(L.T("此操作不可撤销，确定继续？"));
            if (MessageBox.Show(sb.ToString(), L.T("确认卸载"), MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;

            SetBusy(true);
            try
            {
                int ok = 0, fail = 0;
                foreach (string[] it in picked)
                {
                    Log(L.T("正在卸载: ") + it[0]);
                    string code = it[1];
                    int rc = await Task.Run(delegate { return OfficeRemover.RemoveMsi(code, cts.Token); });
                    if (rc == 0 || rc == 3010) { ok++; Log(L.T("  完成 (返回码 ") + rc + ")"); }
                    else { fail++; Log(L.T("  失败 (返回码 ") + rc + ")"); }
                }
                Log(L.T("结束：成功 ") + ok + L.T(" 项，失败 ") + fail + L.T(" 项"));
                if (fail > 0) Log(L.T("失败项可能需要手动到「设置 - 应用」中卸载"));
                RefreshList();
            }
            catch (OperationCanceledException)
            {
                Log(L.T("已取消"));
            }
            catch (Exception ex)
            {
                Log(L.T("失败: ") + ex.Message);
            }
            finally { SetBusy(false); }
        }
    }
}
