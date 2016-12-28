using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.Web.Script.Serialization;
using System.Security.Cryptography;
using System.Threading;

namespace thesisEditer
{
    public partial class FanYiForm : Form
    {
        public string chinese = "";
        public string english = "";
        private Random random = new Random();
        private int salt = 0;
        private string appid = "20160302000014222";
        private string key = "eW5EMcYamWIpO94q9Iih";
        private string sign = "";
        private int type = 0;
        public FanYiForm(string text="",int t = 0)
        {
            InitializeComponent();
            chinese = text;
            richTextBoxEx1.Text = text;
            type = t;
        }
        private void BaiduFY()
        {
            this.circularProgress1.IsRunning = true;
            this.buttonX1.Enabled = false;
            if (this.richTextBoxEx1.Text == "")
            {
                
                this.circularProgress1.IsRunning = false;
                this.buttonX1.Enabled = true;
                return;
            }
            Dictionary<string, object> returnDictionary = new Dictionary<string, object>();
            ArrayList list;
            string msg;
            salt = random.Next(10000);
            MD5 md5 = new MD5CryptoServiceProvider();
            byte[] result = Encoding.UTF8.GetBytes(appid + this.richTextBoxEx1.Text + salt + key);
            byte[] output = md5.ComputeHash(result);
            sign = BitConverter.ToString(output).Replace("-", "").ToLower();
            try
            {
                msg = DocumentSearch.HttpBrowser.GetHttpWebRequest("http://api.fanyi.baidu.com/api/trans/vip/translate?" + "q=" + this.richTextBoxEx1.Text + "&from=zh&to=en&" + "appid=" + appid + "&salt=" + salt + "&sign=" + sign);
            }
            catch
            {
                MessageBox.Show("网络连接失败!", "翻译失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.circularProgress1.IsRunning = false;
                
                this.buttonX1.Enabled = true;
                return;
            }
            JavaScriptSerializer jss = new JavaScriptSerializer();
            returnDictionary = jss.Deserialize<Dictionary<string, object>>(msg);
            try
            {
                chinese = (string)((Dictionary<string, object>)((ArrayList)returnDictionary["trans_result"])[0])["dst"];
                this.richTextBoxEx2.Text = chinese;
                
            }
            catch
            {
                MessageBox.Show("百度词典暂时没有该关键词的翻译哦!", "翻译失败", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            this.circularProgress1.IsRunning = false;
            this.buttonX1.Enabled = true;
        }
        private void beginFy()
        {
            new Thread(BaiduFY).Start();
        }

        private void FanYiForm_Load(object sender, EventArgs e)
        {
            if (type == 1)
            {
                beginFy();
            }
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            if (type == 1)
            {
                if (MessageBox.Show("是否将论文英文名修改为翻译结果？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == System.Windows.Forms.DialogResult.Yes)
                {
                    this.DialogResult = System.Windows.Forms.DialogResult.Yes;
                    return;
                }
            }
            this.DialogResult = System.Windows.Forms.DialogResult.No;
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            beginFy();
        }
        private void 复制ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (richTextBoxEx1.SelectedText != "")
            {
                Clipboard.SetDataObject(richTextBoxEx1.SelectedText);
            }
        }

        private void 粘贴ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBoxEx1.SelectedText = Convert.ToString(Clipboard.GetDataObject().GetData(DataFormats.Text));
        }

        private void 全选ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBoxEx1.SelectAll();
        }

        private void 清空ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBoxEx1.Clear();
        }

        private void 剪切ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (richTextBoxEx1.SelectedText != "")
            {
                Clipboard.SetDataObject(richTextBoxEx1.SelectedText);
                richTextBoxEx1.SelectedText = "";
            }
        }


        private void 复制ToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            if (richTextBoxEx2.SelectedText != "")
            {
                Clipboard.SetDataObject(richTextBoxEx2.SelectedText);
            }
        }

        private void 粘贴ToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            richTextBoxEx2.SelectedText = Convert.ToString(Clipboard.GetDataObject().GetData(DataFormats.Text));
        }

        private void 全选ToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            richTextBoxEx2.SelectAll();
        }

        private void 清空ToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            richTextBoxEx2.Clear();
        }

        private void 剪切ToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            if (richTextBoxEx2.SelectedText != "")
            {
                Clipboard.SetDataObject(richTextBoxEx2.SelectedText);
                richTextBoxEx2.SelectedText = "";
            }
        }
    }
}
