using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Globalization;

namespace thesisEditer
{
    public partial class DocumentLibraryForm : Form
    {
        public DocumentLibraryForm()
        {
            InitializeComponent();
        }
        //System.AppDomain.CurrentDomain.BaseDirectory
        public object DocumentPath = System.AppDomain.CurrentDomain.BaseDirectory + @"\hhtcBooks\";
        public string[] oldList;
        public string DocumentList = "";
        private DevComponents.AdvTree.Node[] searchNodes = null;
        private string searchText = "";
        private int searchIndex = 0;
        private int nodeCount = 0;

        private void readDocument()
        {
            int index = 0;
            nodeCount = 0;
            DirectoryInfo dir = new DirectoryInfo(DocumentPath.ToString());
            FileInfo[] files;
            DirectoryInfo[] dirs = dir.GetDirectories();
            DevComponents.AdvTree.Node node;
            foreach (DirectoryInfo nowDir in dirs)
            {
                node = new DevComponents.AdvTree.Node();
                node.Text = nowDir.Name;
                index = this.advTree1.Nodes.Add(node);
                DirectoryInfo dir1 = new DirectoryInfo(nowDir.FullName);
                files = dir1.GetFiles();
                foreach (FileInfo nowFile in files)
                {
                    node = new DevComponents.AdvTree.Node();
                    node.Text = nowFile.Name.Replace(".txt", "");
                    this.advTree1.Nodes[index].Nodes.Add(node);
                    nodeCount++;
                }
                nodeCount++;
            }
        }

        private void advTree1_NodeDoubleClick(object sender, DevComponents.AdvTree.TreeNodeMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
                return;
            if (e.Node.Nodes.Count != 0)
                return;
            string readText;
            string [] readList;
            StreamReader nowDocument = new StreamReader(DocumentPath + e.Node.Parent.Text + @"\" + e.Node.Text + ".txt", Encoding.UTF8);
            readText = nowDocument.ReadToEnd();
            nowDocument.Close();
            readList = readText.Split('\n');
            this.listViewEx1.BeginUpdate();
            this.listViewEx1.Items.Clear();
            foreach (string text in readList)
            {
                if (text == "")
                    break;
                this.listViewEx1.Items.Add(new ListViewItem(new string[] { text.Split('|')[0].Replace(text.Split('|')[0].Split('.')[0] + ".", ""), text.Split('|')[1], text.Replace(text.Split('.')[0] + ".", "") }));
            }
            this.listViewEx1.EndUpdate();
            oldList = readList;
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            if (this.listViewEx1.CheckedItems.Count == 0)
            {
                this.Close();
                return;
            }
            string text;
            string page = "1-10";
            int st = (int)DateTime.Now.Ticks, end;
            ListView.CheckedListViewItemCollection list =  this.listViewEx1.CheckedItems;
            foreach (ListViewItem nowItem in list)
            {
                text = nowItem.SubItems[2].Text;
                if (this.checkBoxX1.CheckState == CheckState.Checked)
                {
                    st = getRandomNum(200,st);
                    end = st + getRandomNum(50,st);
                    page = st.ToString() + "-" + end.ToString();
                }
                DocumentList = DocumentList + text.Split('|')[1] + "." + text.Split('|')[0] + "[M]." + text.Split('|')[2] + "," + text.Split('|')[3] + ":" + page + "." + "\n";
                
            }
            DocumentList.TrimEnd('\n');
            this.Close();
        }
        private int getRandomNum(int len,int seed)
        {
            Random rd = new Random(unchecked(seed));
            return rd.Next(1, len);
        }
        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void buttonX3_Click(object sender, EventArgs e)
        {
            this.textBoxX1.Clear();
        }
        private void textBoxX1_TextChanged(object sender, EventArgs e)
        {
            if (oldList == null)
                return;
            CompareInfo comp = CultureInfo.InvariantCulture.CompareInfo;
            this.listViewEx1.BeginUpdate();
            this.listViewEx1.Items.Clear();
            foreach (string text in oldList)
            {
                if (text == "")
                    break;
                this.listViewEx1.Items.Add(new ListViewItem(new string[] { text.Split('|')[0].Replace(text.Split('|')[0].Split('.')[0] + ".", ""), text.Split('|')[1], text }));
            }
            this.listViewEx1.EndUpdate();
            if (this.textBoxX1.Text != "")
            {
                ListView.ListViewItemCollection list = this.listViewEx1.Items;
                this.listViewEx1.BeginUpdate();
                foreach (ListViewItem nowItem in list)
                {
                    if (comp.IndexOf(nowItem.SubItems[0].Text, this.textBoxX1.Text,CompareOptions.IgnoreCase) == -1)
                    {
                        nowItem.Remove();
                    }
                }
                this.listViewEx1.EndUpdate();
            }
        }

        private void checkBoxX1_CheckedChanged(object sender, EventArgs e)
        {
            if (this.checkBoxX1.CheckState == CheckState.Unchecked)
            {
                MessageBox.Show("建议使用程序的随机页码,效果非常好!\n(不使用随机页码后，添加上的参考文献会统一使用默认页码\"1-10\"，请记得去参考文献界面修改)", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void DocumentLibraryForm_Load(object sender, EventArgs e)
        {
            readDocument();
        }

        private void listSearch()
        {
            if (searchNodes == null)
            {

                searchText = this.textBoxX2.Text;
                searchNodes = new DevComponents.AdvTree.Node[nodeCount];
                listSerachRecurse();
                searchIndex = 0;
            }
            else
            {
                searchIndex++;
            }
            if (searchNodes.Length != 0)
            {
                if (searchIndex >= searchNodes.Length)
                {
                    searchIndex = 0;
                }
                if (searchNodes[searchIndex] == null)
                {
                    searchIndex = 0;
                }
                this.advTree1.SelectedNode = searchNodes[searchIndex];

            }
        }
        private void listSerachRecurse(DevComponents.AdvTree.Node node = null)
        {
            if (node == null)
            {
                searchIndex = 0;
                foreach (DevComponents.AdvTree.Node k in advTree1.Nodes)
                {
                    if (k.Text.IndexOf(searchText, StringComparison.CurrentCultureIgnoreCase) != -1)
                    {
                        searchNodes[searchIndex] = k;
                        searchIndex++;
                    }
                    if (k.Nodes.Count != 0)
                    {
                        listSerachRecurse(k);
                    }

                }
            }
            else
            {
                foreach (DevComponents.AdvTree.Node k in node.Nodes)
                {
                    if (k.Text.IndexOf(searchText, StringComparison.CurrentCultureIgnoreCase) != -1)
                    {
                        searchNodes[searchIndex] = k;
                        searchIndex++;
                    }
                    if (k.Nodes.Count != 0)
                    {
                        listSerachRecurse(k);
                    }

                }
            }
        }


        private void textBoxX2_TextChanged(object sender, EventArgs e)
        {
            searchNodes = null;
            if (textBoxX2.Text == "")
            {
                buttonX4.Enabled = false;
            }
            else
            {
                buttonX4.Enabled = true;
            }
        }

        private void buttonX4_Click(object sender, EventArgs e)
        {
            listSearch();
        }

    }
}
