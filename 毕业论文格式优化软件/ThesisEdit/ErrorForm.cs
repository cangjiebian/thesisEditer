using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevComponents.DotNetBar;
using System.Net;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using System.Web;
using System.Threading;

namespace thesisEditer
{
    public partial class ErrorForm : Form
    {
        /// <summary>
        /// 构造函数1
        /// </summary>
        /// <param name="message">信息</param>
        public ErrorForm(string message)
        {
            InitializeComponent();
            richTextBoxEx1.Text = message;
        }
        /// <summary>
        /// 构造函数2
        /// </summary>
        /// <param name="title">标题</param>
        /// <param name="message">信息</param>
        /// <param name="or">是否只读</param>
        public ErrorForm(string title, string message,bool or)
        {
            InitializeComponent();
            this.Text = title;
            richTextBoxEx1.Text = message;
            richTextBoxEx1.ReadOnly = or;
            
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
                string s = DocumentSearch.HttpBrowser.GetHttpWebRequest("http://115.159.151.32:8080/thesisediter/errorBack.jsp?" + "message=" + HttpUtility.UrlEncode(richTextBoxEx1.Text, System.Text.Encoding.GetEncoding("gb2312")) + "&office_version=" + HttpUtility.UrlEncode(MainForm.officeVersion, System.Text.Encoding.GetEncoding("gb2312")) + "&thesisediter_version=" + HttpUtility.UrlEncode(MainForm.thesisEditerVersion, System.Text.Encoding.GetEncoding("gb2312")));
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
            
            Thread newFK = new Thread(sendMessage);
            newFK.Start();
        }
        
        
    }
}