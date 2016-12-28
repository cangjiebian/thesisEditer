using System;
using System.Windows.Forms;
using MSWord = Microsoft.Office.Interop.Word;
using System.Drawing;
using System.Threading;
namespace thesisEditer
{
    public partial class EditForm : DevComponents.DotNetBar.Office2007Form
    {
        private string _FilePath;

        public string FilePath
        {
            get { return _FilePath; }
            set 
            { 
                _FilePath = value;
                if (this.WindowState!= FormWindowState.Maximized)
                    this.WindowState = FormWindowState.Maximized;
                DocSave();
                readFile();
            }
        }
        public bool Retry = true;
        public MainForm mainForm = null;
        public DevComponents.AdvTree.Node nowSelectNode = new DevComponents.AdvTree.Node();
        public DevComponents.AdvTree.Node prevNode = new DevComponents.AdvTree.Node();
        public EditForm(MainForm main)
        {
            InitializeComponent();
            mainForm = main;
            
        }
        
        object QS = System.Reflection.Missing.Value;//缺省参数
        MSWord.Document axFramerDoc = null;
        MSWord.Application axFramerApp = null;
        public void readFile()
        {
            if(mainForm==null)
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
                axFramerDoc = (MSWord.Document)axFramerControl1.ActiveDocument;
                axFramerApp = (MSWord.Application)(axFramerDoc.Application);
            }
            catch
            {
                DialogResult re = MessageBox.Show(this,"该章节无法打开,是否重试？", "读取失败", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
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
        private void DocSave()
        {
            if (this.axFramerControl1.DocumentFullName != null)
            {
                if (this.axFramerControl1.IsDirty)
                {

                    System.Windows.Forms.DialogResult dlg = MessageBox.Show(this, "检测到该内容已经发生改变,是否保存?", "保存", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dlg == DialogResult.Yes)
                    {
                        this.axFramerControl1.Save();
                    }

                }
            }
        }
        private void EditForm_FormClosing(object sender, FormClosingEventArgs e)
        {

            DocSave();
            
        }
        private void EditForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.axFramerControl1.Hide();
            this.axFramerControl1.Close();

            MainForm mdi = this.ParentForm as MainForm;
            mdi.advTree_1.DragDropEnabled = true;
            mdi.advTree_1.CellEdit = true;
            mdi.haveDoc = false;
            mdi.nowDocName = "";
            mdi.buttonItem5.Enabled = false;
            mdi.buttonItem7.Enabled = false;
            mdi.buttonItem12.Enabled = false;
            mdi.buttonItem13.Enabled = false;
            mdi.buttonItem24.Visible = false;
            mdi.buttonItem25.Visible = false;
            mdi.labelItem2.Text = "无";
            
        }
        
       


    }
    /*
    internal class AxHostCoverter : AxHost
    {
        private AxHostCoverter() : base("") { }


        static public stdole.IPictureDisp ImageToPictureDisp(Image image)
        {
            return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
        }
        static public Image PictureDispToImage(stdole.IPictureDisp pictureDisp)
        {
            return GetPictureFromIPicture(pictureDisp);
        }
    }
    */ 
}
