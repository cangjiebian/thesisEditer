using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using System.Net;
using System.IO;
using System.Web;
using System.Threading;

namespace thesisEditer
{
    public partial class BackForm : Form
    {
        public BackForm()
        {
            InitializeComponent();
            textBoxX1.Text = MainForm.DocAuthor;
        }
        private void 复制ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.richTextBoxEx1.SelectedText != "")
            {
                Clipboard.SetDataObject(this.richTextBoxEx1.SelectedText);
            }
        }

        private void 粘贴ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.richTextBoxEx1.SelectedText = Convert.ToString(Clipboard.GetDataObject().GetData(DataFormats.Text));
        }

        private void 全选ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.richTextBoxEx1.SelectAll();
        }

        private void 清空ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.richTextBoxEx1.Clear();
        }

        private void 剪切ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.richTextBoxEx1.SelectedText != "")
            {
                Clipboard.SetDataObject(this.richTextBoxEx1.SelectedText);
                this.richTextBoxEx1.SelectedText = "";
            }
        }
        //证书无效问题
        public bool CheckValidationResult(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
        {
            return true;
        }
        
        private void sendMessage()
        {
            this.circularProgress1.IsRunning = true;
            try
            {
                string s = DocumentSearch.HttpBrowser.GetHttpWebRequest("http://115.159.151.32:8080/thesisediter/suggestBack.jsp?" + "name=" + HttpUtility.UrlEncode(textBoxX1.Text, System.Text.Encoding.GetEncoding("gb2312")) + "&qq=" + HttpUtility.UrlEncode(textBoxX2.Text, System.Text.Encoding.GetEncoding("gb2312")) + "&message=" + HttpUtility.UrlEncode(richTextBoxEx1.Text, System.Text.Encoding.GetEncoding("gb2312")) + "&office_version=" + HttpUtility.UrlEncode(MainForm.officeVersion, System.Text.Encoding.GetEncoding("gb2312")) + "&thesisediter_version=" + HttpUtility.UrlEncode(MainForm.thesisEditerVersion, System.Text.Encoding.GetEncoding("gb2312")));
                if (Boolean.Parse(s))
                {
                    MessageBox.Show("感谢您的反馈，我们会及时查看与修改，您的反馈是我们更新的动力！！！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("由于服务器端错误，反馈失败了，非常抱歉！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("反馈失败了，可能由于以下原因：\n1、网络未连接\n2、服务器未开启", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            this.circularProgress1.IsRunning = false;
            this.Close();
        }
        private void buttonX1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            if (textBoxX1.Text == "")
            {
                MessageBox.Show("姓名不能为空", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (richTextBoxEx1.Text == "")
            {
                MessageBox.Show("反馈意见不能为空", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            Thread newFK = new Thread(sendMessage);
            newFK.Start();

        }
    }
}
