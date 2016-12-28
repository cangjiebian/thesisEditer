using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using MSWord = Microsoft.Office.Interop.Word;

namespace thesisEditer
{
    public partial class FirstPageForm : DevComponents.DotNetBar.Office2007Form
    {
        private string _FilePath;

        public string FilePath
        {
            get { return _FilePath; }
            set
            {
                _FilePath = value;
                if (this.WindowState != FormWindowState.Maximized)
                    this.WindowState = FormWindowState.Maximized;
                readFile();
            }
        }
        public bool Retry = true;
        public bool isPrint = true;
        public MainForm mainForm = null;
        public FirstPageForm(MainForm main)
        {
            InitializeComponent();
            mainForm = main;
        }

        private void FirstPageForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.axFramerControl1.Hide();
            this.axFramerControl1.Close();

            MainForm mdi = this.ParentForm as MainForm;
            mdi.advTree_1.DragDropEnabled = true;
            mdi.advTree_1.CellEdit = true;
            mdi.haveDoc = false;
            mdi.buttonItem5.Enabled = false;
            mdi.buttonItem7.Enabled = false;
            mdi.labelItem2.Text = "无";
        }
        public void readFile()
        {
            if (mainForm == null)
                mainForm = ((MainForm)this.ParentForm);
            mainForm.progressBarItem1.ProgressType = DevComponents.DotNetBar.eProgressItemType.Marquee;

            new Thread(readFileThread).Start();


        }
        private void readFileThread()
        {
            try
            {
                this.axFramerControl1.Hide();
                this.axFramerControl1.Open(FilePath);
                this.axFramerControl1.Show();
            }
            catch
            {
                DialogResult re = MessageBox.Show(this, "效果浏览无法打开,是否重试？", "读取失败", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                if (re == System.Windows.Forms.DialogResult.Retry)
                    readFile();
                else
                {
                    Retry = false;
                    this.Close();
                }
            }
            mainForm.progressBarItem1.ProgressType = DevComponents.DotNetBar.eProgressItemType.Standard;
        }


        

        
        
    }
}
