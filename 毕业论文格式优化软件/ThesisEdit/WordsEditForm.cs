using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MSWord = Microsoft.Office.Interop.Word;
using System.IO;
using System.Threading;
namespace thesisEditer
{
    public partial class WordsEditForm : DevComponents.DotNetBar.Office2007Form
    {
        public WordsEditForm()
        {
            InitializeComponent();
        }
        MSWord.Document readDoc;
        MSWord.Application wordApp;
        MainForm father = null;
        public object chWordsPath = System.AppDomain.CurrentDomain.BaseDirectory + @"\office\关键词.doc";
        public object keyWordsPath = System.AppDomain.CurrentDomain.BaseDirectory + @"\office\Key words.doc";
        object QS = System.Reflection.Missing.Value;//缺省参数
        object format = MSWord.WdSaveFormat.wdFormatDocument;//文件类型为word2003，doc
        public bool isEdit = false;
        private void WordsEditForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (isEdit && MessageBox.Show(this, "检测到该内容已经发生改变,是否保存?", "保存", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
            {
                if (this.listViewEx1.Items.Count == 0)
                {
                    try
                    {
                        File.Delete(chWordsPath.ToString());
                    }
                    catch
                    {
                    }
                    try
                    {
                        File.Delete(keyWordsPath.ToString());
                    }
                    catch
                    {
                    }
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
            addList();
        }
        private void readFile()
        {
            father = ((MainForm)this.ParentForm);
            father.progressBarItem1.ProgressType = DevComponents.DotNetBar.eProgressItemType.Marquee;
            new Thread(readFileThread).Start();
            
        }
        private void readFileThread()
        {
            if (File.Exists(chWordsPath.ToString()) == false || File.Exists(keyWordsPath.ToString()) == false)
            {
                father.progressBarItem1.ProgressType = DevComponents.DotNetBar.eProgressItemType.Standard;
                return;
            }
            
            string readText1, readText2;
            wordApp = new MSWord.Application();
            readDoc = wordApp.Documents.Open(ref chWordsPath, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS);
            readText1 = readDoc.Paragraphs[1].Range.Text;
            readText1 = readText1.TrimEnd('\n').TrimEnd('\r').TrimEnd(';');
            readDoc.Close(ref QS, ref QS, ref QS);
            readDoc = wordApp.Documents.Open(ref keyWordsPath, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS);
            readText2 = readDoc.Paragraphs[1].Range.Text;
            readText2 = readText2.TrimEnd('\n').TrimEnd('\r').TrimEnd(';');
            this.listViewEx1.BeginUpdate();
            for (int i = 0; i < readText1.Split(';').Length; i++)
            {
                this.listViewEx1.Items.Add(new ListViewItem(new string[] { readText1.Split(';')[i], readText2.Split(';')[i] }));
            }
            this.listViewEx1.EndUpdate();

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
            string text = "";
            wordApp = new MSWord.Application();
            readDoc = wordApp.Documents.Add(ref QS, ref QS, ref QS, ref QS);
            foreach (ListViewItem nowItem in this.listViewEx1.Items)
            {
                text += nowItem.SubItems[0].Text + ";";
            }
            text = text.TrimEnd(';');
            readDoc.Paragraphs[1].Range.Text = text;
            readDoc.SaveAs(ref chWordsPath, ref format, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS);
            readDoc.Close(ref QS, ref QS, ref QS);
            text = "";
            readDoc = wordApp.Documents.Add(ref QS, ref QS, ref QS, ref QS);
            foreach (ListViewItem nowItem in this.listViewEx1.Items)
            {
                text += nowItem.SubItems[1].Text + ";";
            }
            text = text.TrimEnd(';');
            readDoc.Paragraphs[1].Range.Text = text;
            readDoc.SaveAs(ref keyWordsPath, ref format, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS);
            readDoc.Close(ref QS, ref QS, ref QS);
            wordApp.Quit(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);
        }
        private void listViewEx1_DoubleClick(object sender, EventArgs e)
        {
            editList();
        }

        private void listViewEx1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right)
                return;
            ListViewItem rItem = this.listViewEx1.GetItemAt(e.X, e.Y);
            if (rItem != null)
            {
                rItem.Selected = true;
                ContextMenuStrip cms = new ContextMenuStrip();
                ToolStripMenuItem Edit = new ToolStripMenuItem("编辑");
                ToolStripMenuItem Delete = new ToolStripMenuItem("删除");

                Edit.Click += new EventHandler(listViewEx1_DoubleClick);
                Delete.Click += new EventHandler(Delete_Click);

                cms.Items.Add(Edit);
                cms.Items.Add(Delete);

                cms.Show(this.listViewEx1, e.X, e.Y);
            }
        }
        private void addList()
        {
            WordsImportForm import = new WordsImportForm();
            import.ShowDialog(this);
            if (import.chWord != "" && import.keyWord != "")
            {
                this.listViewEx1.Items.Add(new ListViewItem(new string[] { import.chWord, import.keyWord }));
                isEdit = true;
            }
        }
        private void editList()
        {
            if (this.listViewEx1.SelectedItems.Count != 0)
            {
                if (listViewEx1.SelectedItems.Count != 0)
                {
                    WordsImportForm import = new WordsImportForm();
                    import.Text = "编辑关键词";
                    import.panelEx1.Text = "编辑关键词";
                    import.textBoxX1.Text = this.listViewEx1.SelectedItems[0].SubItems[0].Text;
                    import.textBoxX2.Text = this.listViewEx1.SelectedItems[0].SubItems[1].Text;
                    import.buttonX1.Text = "确定";
                    import.ShowDialog(this);
                    if (import.chWord != "" && import.keyWord != "")
                    {
                        this.listViewEx1.SelectedItems[0].SubItems[0].Text = import.chWord;
                        this.listViewEx1.SelectedItems[0].SubItems[1].Text = import.keyWord;
                        isEdit = true;
                    }
                }
            }
        }
        private void deleteList()
        {
            if (listViewEx1.SelectedItems.Count != 0)
            {
                this.listViewEx1.SelectedItems[0].Remove();
                
                isEdit = true;
            }
        }
        private void Delete_Click(object sender, EventArgs e)
        {

            deleteList();
        }

        private void WordsEditForm_Shown(object sender, EventArgs e)
        {
            if (this.WindowState != FormWindowState.Maximized)
                this.WindowState = FormWindowState.Maximized;
            readFile();
        }

        private void listViewEx1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                deleteList();
            }
        }
    }
}
