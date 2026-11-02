using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace O365Installer
{
    static partial class L
    {
        public static readonly string[] Codes = new string[] { "zh-CN", "en-US", "ja-JP", "ko-KR" };
        public static readonly string[] NativeNames = new string[] { "简体中文", "English", "日本語", "한국어" };

        static int idx;
        static readonly Dictionary<string, string[]> map = new Dictionary<string, string[]>();
        static Dictionary<string, string> revAll;

        static L()
        {
            FillEn();
            FillJa();
            FillKo();
            Load();
        }

        static void Set(string zh, int lang, string text)
        {
            string[] v;
            if (!map.TryGetValue(zh, out v))
            {
                v = new string[4];
                v[0] = zh;
                map[zh] = v;
            }
            v[lang] = text;
        }

        public static void En(string zh, string s) { Set(zh, 1, s); }
        public static void Ja(string zh, string s) { Set(zh, 2, s); }
        public static void Ko(string zh, string s) { Set(zh, 3, s); }

        static partial void FillEn();
        static partial void FillJa();
        static partial void FillKo();

        public static int Index { get { return idx; } }

        public static string T(string zh)
        {
            if (string.IsNullOrEmpty(zh) || idx == 0) return zh;
            string[] v;
            if (!map.TryGetValue(zh, out v)) return zh;
            string s = v[idx];
            if (!string.IsNullOrEmpty(s)) return s;
            if (!string.IsNullOrEmpty(v[1])) return v[1];
            return zh;
        }

        public static string T(string zh, params object[] args)
        {
            string s = T(zh);
            try { return string.Format(s, args); }
            catch { return s; }
        }

        static string PrefPath()
        {
            string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeOfflineInstaller");
            try { Directory.CreateDirectory(dir); } catch { }
            return Path.Combine(dir, "ui-lang.txt");
        }

        static void Load()
        {
            try
            {
                string p = PrefPath();
                if (!File.Exists(p)) return;
                int v;
                if (int.TryParse(File.ReadAllText(p).Trim(), out v) && v >= 0 && v <= 3) idx = v;
            }
            catch { }
        }

        static void Save()
        {
            try { File.WriteAllText(PrefPath(), idx.ToString()); }
            catch { }
        }

        public static void SetLang(int i)
        {
            if (i < 0 || i > 3 || i == idx) return;
            idx = i;
            Save();
        }

        static void BuildRev()
        {
            revAll = new Dictionary<string, string>();
            foreach (KeyValuePair<string, string[]> kv in map)
            {
                for (int i = 0; i < 4; i++)
                {
                    string s = kv.Value[i];
                    if (!string.IsNullOrEmpty(s) && !revAll.ContainsKey(s)) revAll[s] = kv.Key;
                }
            }
        }

        public static void Apply(Control root)
        {
            if (revAll == null) BuildRev();
            ApplyText(root);
            foreach (Control c in root.Controls)
            {
                ApplyText(c);
                if (c.HasChildren) Apply(c);
            }
        }

        static void ApplyText(Control c)
        {
            string t = c.Text;
            if (string.IsNullOrEmpty(t)) return;
            string k;
            if (!revAll.TryGetValue(t, out k)) return;
            string nt = T(k);
            if (nt != t) c.Text = nt;
        }
    }
}
