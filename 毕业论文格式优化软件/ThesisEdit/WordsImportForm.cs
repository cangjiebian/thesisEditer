using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.Web.Script.Serialization;
using System.IO;
using System.Collections;
using System.Threading;
using System.Security.Cryptography;
namespace thesisEditer
{
    public partial class WordsImportForm : Form
    {
        public WordsImportForm()
        {
            InitializeComponent();
        }
        public string chWord = "";
        public string keyWord = "";
        private Random random = new Random();
        private int salt = 0;
        private string appid = "20160302000014222";
        private string key = "eW5EMcYamWIpO94q9Iih";
        private string sign = "";
        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void panelEx4_Click(object sender, EventArgs e)
        {
            if (this.textBoxX1.Text == "")
            {
                MessageBox.Show("请在左侧输入需要翻译的关键字!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            Thread newFY = new Thread(BaiduFY);
            newFY.Start();
            
            
        }
        private void BaiduFY()
        {
            this.circularProgress1.IsRunning = true;
            this.panelEx4.Text = "翻译中";
            this.panelEx4.Enabled = false;
            
            Dictionary<string, object> returnDictionary = new Dictionary<string, object>();
            ArrayList list;
            string msg;
            salt = random.Next(10000);
            MD5 md5 = new MD5CryptoServiceProvider();
            byte[] result = Encoding.UTF8.GetBytes(appid + this.textBoxX1.Text + salt + key);
            byte[] output = md5.ComputeHash(result);
            sign = BitConverter.ToString(output).Replace("-", "").ToLower();
            try
            {
                msg = DocumentSearch.HttpBrowser.GetHttpWebRequest("http://api.fanyi.baidu.com/api/trans/vip/translate?" + "q=" + this.textBoxX1.Text + "&from=zh&to=en&" + "appid=" + appid + "&salt=" + salt + "&sign=" + sign);
            }
            catch
            {
                MessageBox.Show("网络连接失败!", "翻译失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.circularProgress1.IsRunning = false;
                this.panelEx4.Text = "翻译";
                this.panelEx4.Enabled = true;
                return;
            }
            JavaScriptSerializer jss = new JavaScriptSerializer();
            returnDictionary = jss.Deserialize<Dictionary<string, object>>(msg);
            try
            {
                this.textBoxX2.Text = (string)((Dictionary<string, object>)((ArrayList)returnDictionary["trans_result"])[0])["dst"];
                
            }
            catch
            {
                MessageBox.Show("百度词典暂时没有该关键词的翻译哦!", "翻译失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                
            }
            this.circularProgress1.IsRunning = false;
            this.panelEx4.Text = "翻译";
            this.panelEx4.Enabled = true;
        }
        private void buttonX1_Click(object sender, EventArgs e)
        {
            if (this.textBoxX1.Text == "" || this.textBoxX2.Text == "")
            {
                MessageBox.Show("请将中英文关键词填写完整!", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            chWord = this.textBoxX1.Text;
            keyWord = this.textBoxX2.Text;
            this.Close();
        }
        
        

        
    }
}
