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
        public bool Img;
        public bool SingleApp;

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
            l.Add(new ProductDef(id, name, L.T("单独应用"), "Current", false, true));
        }

        public static List<ProductDef> All()
        {
            List<ProductDef> l = new List<ProductDef>();

            Img(l, "O365ProPlusRetail", L.T("Microsoft 365 企业应用版"), "Microsoft 365", "Current");
            Img(l, "O365BusinessRetail", L.T("Microsoft 365 商业版"), "Microsoft 365", "Current");
            Img(l, "O365HomePremRetail", L.T("Microsoft 365 家庭版"), "Microsoft 365", "Current");
            A(l, "O365ProPlusEEANoTeamsRetail", L.T("Microsoft 365 企业应用版（不含 Teams）"), "Microsoft 365", "Current");
            A(l, "O365BusinessEEANoTeamsRetail", L.T("Microsoft 365 商业版（不含 Teams）"), "Microsoft 365", "Current");

            A(l, "ProPlusRetail", L.T("Office 专业增强版 2016"), "Office 2016", "Current");
            A(l, "ProPlusVolume", L.T("Office 专业增强版 2016（批量授权）"), "Office 2016", "PerpetualVL2016");
            A(l, "ProfessionalRetail", L.T("Office 小型企业版 2016"), "Office 2016", "Current");

            string[] years = new string[] { "2024", "2021", "2019" };
            foreach (string y in years)
            {
                string fam = "Office " + y;
                string vl = "PerpetualVL" + y;
                A(l, "ProPlus" + y + "Retail", L.T("Office 专业增强版 ") + y, fam, "Current");
                A(l, "ProPlus" + y + "Volume", L.T("Office LTSC 专业增强版 ") + y + L.T("（批量授权）"), fam, vl);
                A(l, "Standard" + y + "Volume", L.T("Office LTSC 标准版 ") + y + L.T("（批量授权）"), fam, vl);

                if (y == "2024")
                {
                    A(l, "Home2024Retail", L.T("Office 家庭版 2024"), fam, "Current");
                    A(l, "HomeBusiness2024Retail", L.T("Office 家庭商业版 2024"), fam, "Current");
                    A(l, "Professional2024Retail", L.T("Office 小型企业版 2024"), fam, "Current");
                }
                if (y == "2021")
                {
                    A(l, "HomeStudent2021Retail", L.T("Office 家庭学生版 2021"), fam, "Current");
                    A(l, "HomeBusiness2021Retail", L.T("Office 家庭商业版 2021"), fam, "Current");
                    A(l, "Personal2021Retail", L.T("Office 个人版 2021"), fam, "Current");
                    A(l, "Professional2021Retail", L.T("Office 小型企业版 2021"), fam, "Current");
                }
                if (y == "2019")
                {
                    A(l, "HomeStudent2019Retail", L.T("Office 家庭学生版 2019"), fam, "Current");
                    A(l, "HomeBusiness2019Retail", L.T("Office 家庭商业版 2019"), fam, "Current");
                }
            }

            string[] apps = new string[] { "Access", "Excel", "OneNote", "Outlook", "PowerPoint", "Publisher", "Word" };
            string[] appCn = new string[] { L.T("Access 数据库"), "Excel", "OneNote", "Outlook", "PowerPoint", "Publisher", "Word" };
            for (int i = 0; i < apps.Length; i++)
            {
                App(l, apps[i] + "2024Retail", appCn[i] + " 2024");
                App(l, apps[i] + "2021Retail", appCn[i] + " 2021");
                App(l, apps[i] + "2019Retail", appCn[i] + " 2019");
            }
            App(l, "OneNoteFreeRetail", L.T("OneNote（免费版）"));

            foreach (string y in years)
            {
                string fam = "Visio " + y;
                string vl = "PerpetualVL" + y;
                A(l, "VisioPro" + y + "Retail", L.T("Visio 专业版 ") + y, fam, "Current");
                A(l, "VisioPro" + y + "Volume", L.T("Visio LTSC 专业版 ") + y + L.T("（批量授权）"), fam, vl);
                A(l, "VisioStd" + y + "Retail", L.T("Visio 标准版 ") + y, fam, "Current");
                A(l, "VisioStd" + y + "Volume", L.T("Visio LTSC 标准版 ") + y + L.T("（批量授权）"), fam, vl);
            }

            foreach (string y in years)
            {
                string fam = "Project " + y;
                string vl = "PerpetualVL" + y;
                A(l, "ProjectPro" + y + "Retail", L.T("Project 专业版 ") + y, fam, "Current");
                A(l, "ProjectPro" + y + "Volume", L.T("Project 专业版 ") + y + L.T("（批量授权）"), fam, vl);
                A(l, "ProjectStd" + y + "Retail", L.T("Project 标准版 ") + y, fam, "Current");
                A(l, "ProjectStd" + y + "Volume", L.T("Project 标准版 ") + y + L.T("（批量授权）"), fam, vl);
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
