using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevComponents;
using MSWord = Microsoft.Office.Interop.Word;
using System.IO;
using System.Threading;
namespace thesisEditer
{
    public partial class DocumentEditForm : DevComponents.DotNetBar.Office2007Form
    {

        public DocumentEditForm()
        {
            InitializeComponent();
            
        }

        MSWord.Document readDoc;
        MSWord.Application wordApp;
        public object officePath = System.AppDomain.CurrentDomain.BaseDirectory + @"\office\参考文献.doc";
        object QS = System.Reflection.Missing.Value;//缺省参数
        object format = MSWord.WdSaveFormat.wdFormatDocument;//文件类型为word2003，doc
        public bool isEdit = false;
        MainForm father = null;
        private delegate void UpdateNameDelegate();
        private delegate void deleteNameDelegate();
        private void DocumentForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (isEdit && MessageBox.Show(this, "检测到该内容已经发生改变,是否保存?", "保存", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
            {
                if (this.listView1.Items.Count == 0)
                {
                    File.Delete(officePath.ToString());
                }
                else
                {
                    saveFile();
                }
            }
            MainForm mdi = this.ParentForm as MainForm;
            this.Visible = false;
            
            mdi.advTree_1.DragDropEnabled = true;
            mdi.advTree_1.CellEdit = true;
            mdi.haveDoc = false;
            mdi.nowDocName = "";
            mdi.buttonItem5.Enabled = false;
            mdi.buttonItem7.Enabled = false;
            mdi.labelItem2.Text = "无";
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            DocumentSelectForm select = new DocumentSelectForm();
            select.ShowDialog(this);
            if (select.DocumentText != "")
            {
                this.listView1.Items.Add(select.DocumentText);
                UpdateName();
            }
        }
        private void buttonX2_Click(object sender, EventArgs e)
        {
            string DocumentList;
            DocumentLibraryForm library = new DocumentLibraryForm();

            library.ShowDialog(this);
            if (library.DocumentList != "")
            {
                this.listView1.BeginUpdate();
                DocumentList = library.DocumentList;
                foreach (string now in DocumentList.Split('\n'))
                {
                    if (now != "")
                    {
                        this.listView1.Items.Add(now);
                    }
                }
                this.listView1.EndUpdate();
            }
            UpdateName();
        }
       
        private void readFile()
        {
            father = ((MainForm)this.ParentForm);
            father.progressBarItem1.ProgressType = DevComponents.DotNetBar.eProgressItemType.Marquee;
            
            new Thread(readFileThread).Start();
        }
        private void readFileThread()
        {
            if (File.Exists(officePath.ToString()) == false)
            {
                ((MainForm)this.ParentForm).progressBarItem1.ProgressType = DevComponents.DotNetBar.eProgressItemType.Standard;
                return;
            }
            wordApp = new MSWord.Application();
            readDoc = wordApp.Documents.Open(ref officePath, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS);
            this.listView1.BeginUpdate();
            foreach (MSWord.Paragraph nowPar in readDoc.Paragraphs)
            {
                this.listView1.Items.Add(nowPar.Range.Text);
            }
            this.listView1.EndUpdate();
            readDoc.Close(ref QS, ref QS, ref QS);
            wordApp.Quit(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);
            father.progressBarItem1.ProgressType = DevComponents.DotNetBar.eProgressItemType.Standard;
        }
        private void saveFile()
        {
            saveFileThread();
        }
        private void saveFileThread()
        {
            wordApp = new MSWord.Application();
            readDoc = wordApp.Documents.Add(ref QS, ref QS, ref QS, ref QS);
            foreach (ListViewItem nowItem in this.listView1.Items)
            {
                readDoc.Paragraphs.Last.Range.Text = nowItem.Text + "\n";
            }
            readDoc.Paragraphs.Last.Range.Delete();
            readDoc.SaveAs(ref officePath, ref format, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS);
            readDoc.Close(ref QS, ref QS, ref QS);
            wordApp.Quit(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);
        }
        
        private void UpdateName()
        {
            this.listView1.BeginUpdate();
            foreach (ListViewItem nowItem in this.listView1.Items)
            {
                if (nowItem.Text.Split(']')[0] != nowItem.Text && (nowItem.Text[1] > '0' && nowItem.Text[1] <= '9'))
                {
                    nowItem.Text = nowItem.Text.Replace(nowItem.Text.Split(']')[0], "[" + (nowItem.Index + 1).ToString());
                }
                else
                {
                    nowItem.Text = "[" + (nowItem.Index + 1).ToString() + "]" + nowItem.Text;
                }
                
            }
            this.listView1.EndUpdate();
            isEdit = true;
        }
        
        private void listView1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right)
                return;
            ListViewItem rItem = this.listView1.GetItemAt(e.X, e.Y);
            if (rItem != null)
            {
                rItem.Selected = true;
                ContextMenuStrip cms = new ContextMenuStrip();
                ToolStripMenuItem Edit = new ToolStripMenuItem("编辑");
                ToolStripMenuItem Delete = new ToolStripMenuItem("删除");

                Edit.Click += new EventHandler(Edit_Click);
                Delete.Click += new EventHandler(Delete_Click);

                cms.Items.Add(Edit);
                cms.Items.Add(Delete);
                
                cms.Show(this.listView1, e.X, e.Y);
            }
            
        }

        private void Edit_Click(object sender, EventArgs e)
        {
            editList();
        }
        private void deleteList()
        {
            if (listView1.SelectedItems.Count != 0)
            {
                this.listView1.SelectedItems[0].Remove();
                UpdateName();
            }
        }
        private void editList()
        {
            if (listView1.SelectedItems.Count != 0)
            {
                this.listView1.SelectedItems[0].BeginEdit();
            }
        }
        private void Delete_Click(object sender, EventArgs e)
        {
            deleteList();

        }

        private void listView1_AfterLabelEdit(object sender, LabelEditEventArgs e)
        {
            UpdateNameDelegate upd = new UpdateNameDelegate(UpdateName);
            this.listView1.BeginInvoke(upd);
        }

        private void DocumentForm_Shown(object sender, EventArgs e)
        {
            if (this.WindowState != FormWindowState.Maximized)
                this.WindowState = FormWindowState.Maximized;
            readFile();
        }

        private void listView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                deleteList();
            }
        }

        private void buttonItem2_Click(object sender, EventArgs e)
        {
            string DocumentList;
            DocumentSearch.SearchMainForm searchForm = new DocumentSearch.SearchMainForm();

            searchForm.ShowDialog(this);
            if (searchForm.returnBookStr != "")
            {
                this.listView1.BeginUpdate();
                DocumentList = searchForm.returnBookStr;
                foreach (string now in DocumentList.Split('\n'))
                {
                    if (now != "")
                    {
                        this.listView1.Items.Add(now);
                    }
                }
                this.listView1.EndUpdate();
            }
            UpdateName();
        }

        private void buttonItem1_Click(object sender, EventArgs e)
        {
            buttonX2_Click(sender, e);
        }

        

        

        

        
        

        

        

        

       
    }
}
