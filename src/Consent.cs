using System;
using System.Drawing;
using System.Windows.Forms;

namespace O365Installer
{

    class ConsentForm : Form
    {
        public static string Body { get { return L.T(BodyZh); } }

        public const string BodyZh =
            "使用本程序前请阅读以下条款：" + "\r\n\r\n" +
            "1. 本程序是微软 Office 部署工具（ODT / setup.exe）的图形化前端，仅调用微软官方公开的安装与部署机制；" +
            "安装源来自微软官方 CDN（officecdn.microsoft.com）。" + "\r\n\r\n" +
            "2. 本程序不提供、不内置、不协助任何破解、激活绕过或规避技术保护措施的功能，也不会修改系统的授权状态。" + "\r\n\r\n" +
            "3. 本程序不附带任何 Microsoft Office 授权。使用 Office 需自行通过合法途径取得许可，包括但不限于：" +
            "Microsoft 365 订阅、零售版产品密钥、学校或所在组织的批量授权。" + "\r\n\r\n" +
            "4. 「授权管理」功能仅调用微软自带的 OSPP.VBS，用于查询授权状态、安装使用者自行提供的产品密钥、" +
            "配置使用者自行提供的 KMS 服务器地址。使用者应确保其对相关密钥与授权服务器拥有合法权利。" + "\r\n\r\n" +
            "5. 使用者应自行承担使用本程序产生的一切后果，包括但不限于授权合规性与数据安全。作者不承担任何责任。" + "\r\n\r\n" +
            "6. 本程序按「现状」提供，不提供任何明示或暗示的担保。" + "\r\n\r\n" +
            "7. Microsoft、Office、Microsoft 365 等为微软公司商标，本程序与微软公司无任何隶属或背书关系。";

        public ConsentForm(bool consent)
        {
            Text = consent ? L.T("免责声明 - 请阅读并确认") : L.T("免责声明");
            ClientSize = new Size(600, 452);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = consent;
            Font = new Font("Microsoft YaHei UI", 9F);
            BackColor = Color.FromArgb(248, 249, 251);

            TextBox tb = new TextBox();
            tb.Multiline = true;
            tb.ReadOnly = true;
            tb.ScrollBars = ScrollBars.Vertical;
            tb.Location = new Point(16, 16);
            tb.Size = new Size(568, 370);
            tb.BackColor = Color.FromArgb(252, 252, 253);
            tb.Text = Body.Replace("\r\n", Environment.NewLine);
            Controls.Add(tb);

            if (consent)
            {
                Label t = new Label();
                t.Text = L.T("同意后方可使用本程序；不同意将直接退出");
                t.ForeColor = Color.FromArgb(120, 125, 135);
                t.Font = new Font("Microsoft YaHei UI", 8F);
                t.Location = new Point(16, 394);
                t.AutoSize = true;
                Controls.Add(t);

                Button ok = new Button();
                ok.Text = L.T("同意并继续");
                ok.Location = new Point(352, 396);
                ok.Size = new Size(120, 36);
                ok.BackColor = Color.FromArgb(0, 120, 212);
                ok.ForeColor = Color.White;
                ok.FlatStyle = FlatStyle.Flat;
                ok.FlatAppearance.BorderSize = 0;
                ok.DialogResult = DialogResult.OK;
                Controls.Add(ok);

                Button no = new Button();
                no.Text = L.T("不同意，退出");
                no.Location = new Point(482, 396);
                no.Size = new Size(104, 36);
                no.DialogResult = DialogResult.Cancel;
                Controls.Add(no);

                AcceptButton = ok;
                CancelButton = no;
            }
            else
            {
                Button close = new Button();
                close.Text = L.T("关闭");
                close.Location = new Point(482, 396);
                close.Size = new Size(104, 36);
                close.DialogResult = DialogResult.OK;
                Controls.Add(close);
                AcceptButton = close;
            }
        }

        public static bool Ask()
        {
            using (ConsentForm f = new ConsentForm(true))
            {
                return f.ShowDialog() == DialogResult.OK;
            }
        }
    }
}
