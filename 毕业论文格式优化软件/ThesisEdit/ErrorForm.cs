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
        /// ���캯��1
        /// </summary>
        /// <param name="message">��Ϣ</param>
        public ErrorForm(string message)
        {
            InitializeComponent();
            richTextBoxEx1.Text = message;
        }
        /// <summary>
        /// ���캯��2
        /// </summary>
        /// <param name="title">����</param>
        /// <param name="message">��Ϣ</param>
        /// <param name="or">�Ƿ�ֻ��</param>
        public ErrorForm(string title, string message,bool or)
        {
            InitializeComponent();
            this.Text = title;
            richTextBoxEx1.Text = message;
            richTextBoxEx1.ReadOnly = or;
            
        }
        //֤����Ч����
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
                    MessageBox.Show("��л���ķ��������ǻἰʱ�鿴���޸ģ����ķ��������Ǹ��µĶ���������", "��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("���ڷ������˴��󣬷���ʧ���ˣ��ǳ���Ǹ��", "��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception e) 
            {
                MessageBox.Show("����ʧ���ˣ�������������ԭ��\n1������δ����\n2��������δ����", "��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Error);
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