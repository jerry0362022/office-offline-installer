using System;
using System.Collections.Generic;

namespace O365Installer
{
    class ProductDef
    {
        public string Id;
        public string Name;
        public string Family;
        public string Channel;
        public bool Img;        // 是否有微软官方 .img 离线镜像（仅 Microsoft 365 零售）
        public bool SingleApp;  // 单应用产品，不走组件排除

        public ProductDef(string id, string name, string family, string channel, bool img, bool singleApp)
        {
            Id = id;
            Name = name;
            Family = family;
            Channel = channel;
            Img = img;
            SingleApp = singleApp;
        }
    }

    static class Catalog
    {
        static void A(List<ProductDef> l, string id, string name, string family, string channel)
        {
            l.Add(new ProductDef(id, name, family, channel, false, false));
        }

        static void Img(List<ProductDef> l, string id, string name, string family, string channel)
        {
            l.Add(new ProductDef(id, name, family, channel, true, false));
        }

        static void App(List<ProductDef> l, string id, string name)
        {
            l.Add(new ProductDef(id, name, "单独应用", "Current", false, true));
        }

        // 产品清单参照微软官方 Office 部署工具支持的产品 ID
        // 零售版使用 Current 通道；批量/LTSC 使用 PerpetualVL<年份> 通道
        public static List<ProductDef> All()
        {
            List<ProductDef> l = new List<ProductDef>();

            // ---------- Microsoft 365 ----------
            Img(l, "O365ProPlusRetail", "Microsoft 365 企业应用版", "Microsoft 365", "Current");
            Img(l, "O365BusinessRetail", "Microsoft 365 商业版", "Microsoft 365", "Current");
            Img(l, "O365HomePremRetail", "Microsoft 365 家庭版", "Microsoft 365", "Current");
            A(l, "O365ProPlusEEANoTeamsRetail", "Microsoft 365 企业应用版（不含 Teams）", "Microsoft 365", "Current");
            A(l, "O365BusinessEEANoTeamsRetail", "Microsoft 365 商业版（不含 Teams）", "Microsoft 365", "Current");

            // ---------- Office 2019 / 2021 / 2024 ----------
            string[] years = new string[] { "2024", "2021", "2019" };
            foreach (string y in years)
            {
                string fam = "Office " + y;
                string vl = "PerpetualVL" + y;
                A(l, "ProPlus" + y + "Retail", "Office 专业增强版 " + y, fam, "Current");
                A(l, "ProPlus" + y + "Volume", "Office LTSC 专业增强版 " + y + "（批量授权）", fam, vl);
                A(l, "Standard" + y + "Volume", "Office LTSC 标准版 " + y + "（批量授权）", fam, vl);

                if (y == "2024")
                {
                    A(l, "Home2024Retail", "Office 家庭版 2024", fam, "Current");
                    A(l, "HomeBusiness2024Retail", "Office 家庭商业版 2024", fam, "Current");
                    A(l, "Professional2024Retail", "Office 小型企业版 2024", fam, "Current");
                }
                if (y == "2021")
                {
                    A(l, "HomeStudent2021Retail", "Office 家庭学生版 2021", fam, "Current");
                    A(l, "HomeBusiness2021Retail", "Office 家庭商业版 2021", fam, "Current");
                    A(l, "Personal2021Retail", "Office 个人版 2021", fam, "Current");
                    A(l, "Professional2021Retail", "Office 小型企业版 2021", fam, "Current");
                }
                if (y == "2019")
                {
                    A(l, "HomeStudent2019Retail", "Office 家庭学生版 2019", fam, "Current");
                    A(l, "HomeBusiness2019Retail", "Office 家庭商业版 2019", fam, "Current");
                }
            }

            // ---------- 单独应用 ----------
            string[] apps = new string[] { "Access", "Excel", "OneNote", "Outlook", "PowerPoint", "Publisher", "Word" };
            string[] appCn = new string[] { "Access 数据库", "Excel", "OneNote", "Outlook", "PowerPoint", "Publisher", "Word" };
            for (int i = 0; i < apps.Length; i++)
            {
                App(l, apps[i] + "2024Retail", appCn[i] + " 2024");
                App(l, apps[i] + "2021Retail", appCn[i] + " 2021");
                App(l, apps[i] + "2019Retail", appCn[i] + " 2019");
            }
            App(l, "OneNoteFreeRetail", "OneNote（免费版）");

            // ---------- Visio ----------
            foreach (string y in years)
            {
                string fam = "Visio " + y;
                string vl = "PerpetualVL" + y;
                A(l, "VisioPro" + y + "Retail", "Visio 专业版 " + y, fam, "Current");
                A(l, "VisioPro" + y + "Volume", "Visio LTSC 专业版 " + y + "（批量授权）", fam, vl);
                A(l, "VisioStd" + y + "Retail", "Visio 标准版 " + y, fam, "Current");
                A(l, "VisioStd" + y + "Volume", "Visio LTSC 标准版 " + y + "（批量授权）", fam, vl);
            }

            // ---------- Project ----------
            foreach (string y in years)
            {
                string fam = "Project " + y;
                string vl = "PerpetualVL" + y;
                A(l, "ProjectPro" + y + "Retail", "Project 专业版 " + y, fam, "Current");
                A(l, "ProjectPro" + y + "Volume", "Project 专业版 " + y + "（批量授权）", fam, vl);
                A(l, "ProjectStd" + y + "Retail", "Project 标准版 " + y, fam, "Current");
                A(l, "ProjectStd" + y + "Volume", "Project 标准版 " + y + "（批量授权）", fam, vl);
            }

            return l;
        }

        public static ProductDef Find(List<ProductDef> list, string id)
        {
            foreach (ProductDef p in list) if (p.Id == id) return p;
            return null;
        }
    }
}
