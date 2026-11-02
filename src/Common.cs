using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace O365Installer
{
    static class Net
    {
        public static string Get(string url)
        {
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            req.Timeout = 60000;
            req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)";
            req.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
            using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse())
            using (StreamReader sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8))
            {
                return sr.ReadToEnd();
            }
        }

        public static void Download(string url, string dest, CancellationToken ct)
        {
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            req.Timeout = 60000;
            req.ReadWriteTimeout = 60000;
            req.UserAgent = "Mozilla/5.0";
            using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse())
            using (Stream st = resp.GetResponseStream())
            using (FileStream fs = new FileStream(dest, FileMode.Create, FileAccess.Write))
            {
                byte[] buf = new byte[65536];
                while (true)
                {
                    if (ct.IsCancellationRequested) ct.ThrowIfCancellationRequested();
                    int n = st.Read(buf, 0, buf.Length);
                    if (n <= 0) break;
                    fs.Write(buf, 0, n);
                }
            }
        }
    }

    static class OdtTools
    {
        public const string FallbackUrl = "https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_20430-20146.exe";

        public static string Ensure(string workDir, Action<string> log, CancellationToken ct)
        {
            string dir = Path.Combine(workDir, "odt");
            string setup = Path.Combine(dir, "setup.exe");
            if (File.Exists(setup)) return setup;

            Directory.CreateDirectory(dir);
            string odtExe = Path.Combine(workDir, "odt.exe");

            if (!File.Exists(odtExe) || new FileInfo(odtExe).Length < 500000)
            {
                string url = null;
                try
                {
                    string page = Net.Get("https://www.microsoft.com/zh-cn/download/details.aspx?id=49117");
                    Match m = Regex.Match(page, @"https://download\.microsoft\.com/download/[0-9a-fA-F\-]+/officedeploymenttool_[0-9\-]+\.exe");
                    if (m.Success) url = m.Value;
                }
                catch { }
                if (string.IsNullOrEmpty(url)) url = FallbackUrl;
                if (log != null) log(L.T("ODT 下载地址: ") + url);
                Net.Download(url, odtExe, ct);
                if (log != null) log(L.T("ODT 下载完成: ") + (new FileInfo(odtExe).Length / 1024) + " KB");
            }

            ProcessStartInfo psi = new ProcessStartInfo(odtExe);
            psi.Arguments = "/extract:\"" + dir + "\" /quiet";
            psi.UseShellExecute = false;
            psi.CreateNoWindow = true;
            using (Process p = Process.Start(psi)) p.WaitForExit(180000);

            for (int i = 0; i < 20 && !File.Exists(setup); i++) Thread.Sleep(500);
            if (!File.Exists(setup)) throw new Exception(L.T("ODT 解压失败，可能是安全软件拦截了部署工具。"));
            return setup;
        }

        public static int RunSetup(string setupExe, string args, string workDir, Action<string> tick, CancellationToken ct)
        {
            ProcessStartInfo psi = new ProcessStartInfo(setupExe);
            psi.Arguments = args;
            psi.UseShellExecute = false;
            if (!string.IsNullOrEmpty(workDir)) psi.WorkingDirectory = workDir;
            using (Process p = Process.Start(psi))
            {
                DateTime start = DateTime.Now;
                while (!p.HasExited)
                {
                    if (ct.IsCancellationRequested)
                    {
                        try { p.Kill(); } catch { }
                        ct.ThrowIfCancellationRequested();
                    }
                    Thread.Sleep(500);
                    if (tick != null) tick(L.T("已用时 ") + (DateTime.Now - start).ToString(@"mm\:ss"));
                }
                return p.ExitCode;
            }
        }
    }
}
