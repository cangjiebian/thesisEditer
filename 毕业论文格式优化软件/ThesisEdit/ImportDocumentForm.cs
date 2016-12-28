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
using System.Xml;

namespace thesisEditer
{
    public partial class ImportDocumentForm : Form
    {
        MSWord.Application wordApp;
        MSWord.Application tempApp;
        MSWord.Document wordDoc;
        MSWord.Document tempDoc;
        object format = MSWord.WdSaveFormat.wdFormatDocument;
        string docName = @"E:\c#学习\毕业论文编辑器\毕业论文编辑器\bin\Debug\我的毕业论文.doc";
        string officetemp = System.AppDomain.CurrentDomain.BaseDirectory + "officetemp\\";
        string office = System.AppDomain.CurrentDomain.BaseDirectory + "office\\";
        string mulu_font_style = "黑体", list_font_style = "宋体";
        float mulu_font_size = 15f, list_font_size = 12f;
        object QS = System.Reflection.Missing.Value;//缺省参数
        public ImportDocumentForm()
        {
            InitializeComponent();
            if (Directory.Exists("officetemp") == false)//如果不存在就创建officetemp文件夹
            {
                Directory.CreateDirectory("officetemp");
            }
            initWord();
            xmlDoc = new XmlDocument();
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBoxX1.Text = openFileDialog1.FileName;
                docName = openFileDialog1.FileName;
                readList();

            }
        }
        private void initWord()
        {
            wordApp = new MSWord.Application();
            tempApp = new MSWord.Application();

        }
        private void quitWord()
        {
            try
            {
                tempDoc.Close(false, ref QS, ref QS);
            }
            catch { }
            try
            {
                wordDoc.Close(false, ref QS, ref QS);
            }
            catch { }
            try
            {
                wordApp.Quit(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);
            }
            catch { }
            try
            {
                tempApp.Quit(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);
            }
            catch { }
        }
        private void readListThread()
        {

            wordDoc = wordApp.Documents.Open((object)docName, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS);
            wordApp.ActiveWindow.View.ShowFieldCodes = false;
            wordApp.Visible = false;
            advTree1.Nodes.Clear();
            try
            {
                if (wordDoc.TablesOfContents.Count == 0)
                {
                    if (MessageBox.Show(this, "检测到该文档中不包含目录，是否手动创建？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.Yes)
                    {
                        createList();

                    }
                    else
                    {
                        wordDoc.Close(false, ref QS, ref QS);
                        progressBarX1.ProgressType = DevComponents.DotNetBar.eProgressItemType.Standard;
                        return;
                    }


                }
                MSWord.Paragraphs pgs = wordDoc.TablesOfContents[1].Range.Paragraphs;
                if (pgs.Count == 1 && pgs.First.Range.Text.IndexOf("未找到目录项") != -1)
                {
                    MessageBox.Show(this, "创建失败，该文档没有目录！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    progressBarX1.ProgressType = DevComponents.DotNetBar.eProgressItemType.Standard;
                    return;
                }
                int num = 0;
                foreach (MSWord.Paragraph pa in pgs)
                {
                    if (pa.Range.Text.Split('\t')[0].Trim() != "")
                        num++;
                    if (pa.Range.Text.Trim() != "" && pa.Range.Text.Trim()[0] == '摘' && num == 1)
                    {
                        createNode("摘  要");
                    }
                    else if (pa.Range.Text.Trim() != "" && pa.Range.Text.Trim()[0] == '关' && num == 2)
                    {
                        createNode("关键词");
                    }
                    else if (pa.Range.Text.Trim() != "" && (pa.Range.Text.Trim()[0] == 'A' || pa.Range.Text.Trim()[0] == 'a') && num == 3)
                    {
                        createNode("Abstract");
                    }
                    else if (pa.Range.Text.Trim() != "" && (pa.Range.Text.Trim()[0] == 'K' || pa.Range.Text.Trim()[0] == 'k') && num == 4)
                    {
                        createNode("Key word");
                    }
                    else if (pa.Range.Text.Trim() != "" && pa.Range.Text.Trim()[0] == '参' && num == pgs.Count - 2)
                    {
                        createNode("参考文献");
                    }
                    else if (pa.Range.Text.Trim() != "" && pa.Range.Text.Trim()[0] == '致' && num == pgs.Count - 1)
                    {
                        createNode("致  谢");
                    }
                    else
                    {
                        createNode(titleHandle(pa.Range.Text.Split('\t')[0].Trim()));
                    }
                    Application.DoEvents();
                }
                progressBarX1.Maximum = num;
            }
            catch
            {

                advTree1.Nodes.Clear();
                MSWord.Paragraphs pgs = wordDoc.TablesOfContents[1].Range.Paragraphs;
                foreach (MSWord.Paragraph pa in pgs)
                {
                    if (pa.Range.Text.Split('\t')[0].Trim() != "")
                        advTree1.Nodes.Add(new DevComponents.AdvTree.Node(pa.Range.Text.Split('\t')[0].Trim()));
                }
                MessageBox.Show(this, "尝试自动排版目录失败，请手动拖拽排版目录！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            progressBarX1.ProgressType = DevComponents.DotNetBar.eProgressItemType.Standard;
        }
        private void readList()
        {
            progressBarX1.ProgressType = DevComponents.DotNetBar.eProgressItemType.Marquee;
            new Thread(readListThread).Start();

        }
        private void readInfo()
        {
            wordDoc = wordApp.Documents.Open((object)docName, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS);
            wordApp.ActiveWindow.View.ShowFieldCodes = false;
            //wordApp.Visible = true;
            MSWord.Paragraphs pgs = wordDoc.TablesOfContents[1].Range.Paragraphs;

            int start = wordDoc.GoTo(MSWord.WdGoToItem.wdGoToPage, MSWord.WdGoToDirection.wdGoToAbsolute, QS, 1).Start;
            int end = wordDoc.GoTo(MSWord.WdGoToItem.wdGoToPage, MSWord.WdGoToDirection.wdGoToAbsolute, QS, 2).Start;
            //richTextBoxEx1.Text = wordDoc.Range(start, end).Text;
            //richTextBoxEx1.AppendText("\n"+pgs[1].Next().Range.Hyperlinks[1]);
        }
        private void readDocument()
        {
            if (MessageBox.Show(this, "导入后你会清空软件中原有的论文信息，你确定要导入吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.OK)
                new Thread(createThesis).Start();
        }
        XmlDocument xmlDoc;
        XmlNode xmlNode;
        XmlElement xmlEle;
        private void ergodicNodeSave(DevComponents.AdvTree.Node Node, XmlNode xml)
        {
            if (Node == null)
            {
                xmlNode = xmlDoc.SelectSingleNode("Root");
                xmlNode = xmlNode.SelectSingleNode("TreeNode");
                xmlEle = xmlDoc.CreateElement("Node--999");
                xmlEle.SetAttribute("Name", "摘  要");
                xmlNode.AppendChild(xmlEle);
                xmlEle = xmlDoc.CreateElement("Node--999");
                xmlEle.SetAttribute("Name", "关键词");
                xmlNode.AppendChild(xmlEle);
                xmlEle = xmlDoc.CreateElement("Node--999");
                xmlEle.SetAttribute("Name", "Abstract");
                xmlNode.AppendChild(xmlEle);
                xmlEle = xmlDoc.CreateElement("Node--999");
                xmlEle.SetAttribute("Name", "Key words");
                xmlNode.AppendChild(xmlEle);
                foreach (DevComponents.AdvTree.Node nowNode in advTree1.Nodes)
                {

                    xmlNode = xmlDoc.SelectSingleNode("Root");
                    xmlNode = xmlNode.SelectSingleNode("TreeNode");
                    xmlEle = xmlDoc.CreateElement("Node-" + (nowNode.Index).ToString());
                    //xmlEle.InnerText = nowNode.Text;
                    xmlEle.SetAttribute("Name", nowNode.Text);
                    xmlNode.AppendChild(xmlEle);
                    if (nowNode.Nodes.Count != 0)
                    {
                        ergodicNodeSave(nowNode, xmlNode);
                    }

                }
                xmlNode = xmlDoc.SelectSingleNode("Root");
                xmlNode = xmlNode.SelectSingleNode("TreeNode");
                xmlEle = xmlDoc.CreateElement("Node--999");
                xmlEle.SetAttribute("Name", "参考文献");
                xmlNode.AppendChild(xmlEle);
                xmlEle = xmlDoc.CreateElement("Node--999");
                xmlEle.SetAttribute("Name", "致  谢");
                xmlNode.AppendChild(xmlEle);
            }
            else
            {
                foreach (DevComponents.AdvTree.Node nowNode in Node.Nodes)
                {

                    xmlNode = xml.SelectSingleNode("Node-" + Node.Index);
                    xmlEle = xmlDoc.CreateElement("Node-" + nowNode.Index);
                    xmlEle.SetAttribute("Name", nowNode.Text);
                    //xmlEle.InnerText = nowNode.Text;
                    xmlNode.AppendChild(xmlEle);
                    if (nowNode.Nodes.Count != 0)
                    {
                        ergodicNodeSave(nowNode, xmlNode);
                    }

                }

            }
        }
        private void listSave()
        {
            xmlDoc.Load(office + "NodeInfo.xml");
            xmlNode = xmlDoc.SelectSingleNode("Root");
            xmlNode.SelectSingleNode("TreeNode").RemoveAll();
            ergodicNodeSave(null, null);
            xmlDoc.Save(officetemp + "NodeInfo.xml");
        }
        private void onlyListSave()
        {
            if (advTree1.Nodes.Count == 0)
            {
                MessageBox.Show(this, "请先选择文档", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (MessageBox.Show(this, "导入后你会清空软件中原有的目录信息，你确定要导入吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.OK)
            {
                xmlDoc.Load(office + "NodeInfo.xml");
                xmlNode = xmlDoc.SelectSingleNode("Root");
                xmlNode.SelectSingleNode("TreeNode").RemoveAll();
                ergodicNodeSave(null, null);
                xmlDoc.Save(office + "NodeInfo.xml");
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
                this.Close();
            }
        }
        private void saveDocument()
        {

            Directory.Delete(office, true);
            Directory.Move(officetemp, office);
        }
        private void createList()
        {
            MSWord.Range range = wordDoc.GoTo(MSWord.WdGoToItem.wdGoToPage, MSWord.WdGoToDirection.wdGoToAbsolute, QS, 1);
            range.InsertAfter("\r\n");
            range = wordDoc.GoTo(MSWord.WdGoToItem.wdGoToPage, MSWord.WdGoToDirection.wdGoToAbsolute, QS, 1);
            wordDoc.TablesOfContents.Add(range, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, true, ref QS, ref QS);
            wordDoc.TablesOfContents[1].Range.Font.Name = list_font_style;
            wordDoc.TablesOfContents[1].Range.Font.Size = list_font_size;
            wordDoc.TablesOfContents[1].Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceExactly;
            wordDoc.TablesOfContents[1].Range.ParagraphFormat.LineSpacing = 22F;
            range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
            range.Text = "\n目  录\n\n";
            range.Font.Name = mulu_font_style;
            range.Font.Size = mulu_font_size;
            range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceExactly;
            range.ParagraphFormat.LineSpacing = 22F;
        }
        private void createThesis()
        {

            buttonX3.Enabled = false;
            int start, end;
            MSWord.Paragraphs pgs = wordDoc.TablesOfContents[1].Range.Paragraphs;
            wordApp.Selection.SetRange(wordDoc.TablesOfContents[1].Range.End, wordDoc.TablesOfContents[1].Range.End);
            foreach (MSWord.Paragraph pa in pgs)
            {

                if (pa.Range.Text.Split('\t')[0].Trim() != "")
                {

                    pa.Range.Hyperlinks[1].Follow();
                    //wordApp.Visible = false;
                    start = wordApp.Selection.Range.End + 1;
                    wordApp.Selection.SetRange(start, start);
                    if (pa.Next() != null)
                    {
                        if (pa.Next().Range.Text.Trim() == "")
                        {
                            end = wordDoc.Content.End;
                        }
                        else
                        {
                            pa.Next().Range.Hyperlinks[1].Follow();
                            //wordApp.Visible = false;
                            end = wordApp.Selection.Range.Start - 1;
                        }
                        wordApp.Selection.SetRange(start, end);
                    }
                    else
                    {
                        break;
                    }

                    try
                    {
                        if (wordApp.Selection.Range.Text.Trim() != "")
                        {
                            wordApp.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                            wordApp.Selection.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(float.Parse("0"));
                            wordApp.Selection.ParagraphFormat.RightIndent = wordApp.CentimetersToPoints(float.Parse("0"));
                            wordApp.Selection.ClearFormatting();
                            wordApp.Selection.Copy();
                            tempDoc = tempApp.Documents.Add();
                            tempDoc.Paragraphs.Last.Range.Paste();
                            tempDoc.SaveAs(officetemp + titleHandle(pa.Range.Text.Split('\t')[0].Trim()) + ".doc", ref format, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS);
                            tempDoc.Close(ref QS, ref QS, ref QS);
                        }
                    }
                    catch { }
                    wordApp.Selection.SetRange(end, end);
                    progressBarX1.Value++;
                }
            }
            listSave();
            saveDocument();
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }
        private bool createNode(string idname)
        {
            if (idname == "")
                return false;
            if (idname.Split(' ').Length == 1)
            {
                DevComponents.AdvTree.Node addnode = new DevComponents.AdvTree.Node();
                addnode.Text = idname;
                advTree1.Nodes.Add(addnode);
            }
            else
            {
                string id = idname.Split(' ')[0];
                DevComponents.AdvTree.NodeCollection nodes = null;
                DevComponents.AdvTree.Node node = null;
                if (id.Split('.').Length == 1)
                {
                    nodes = advTree1.Nodes;
                }
                else
                {
                    int i = 0;
                    foreach (string num in id.Split('.'))
                    {

                        if (i == 0)
                        {
                            node = advTree1.Nodes[advTree1.Nodes.Count - 1];
                        }
                        else if (i == id.Split('.').Length - 1)
                        {
                            nodes = node.Nodes;
                        }
                        else
                        {
                            node = node.Nodes[node.Nodes.Count - 1];
                        }
                        i++;
                    }
                }
                DevComponents.AdvTree.Node addnode = new DevComponents.AdvTree.Node();
                addnode.Text = idname;
                nodes.Add(addnode);
            }
            return false;
        }
        private string titleHandle(string name)
        {
            if (name == "")
                return name;
            StringBuilder sb = new StringBuilder(name);
            if (!((sb[0] > '0' && sb[0] < '9')))
                return name;
            int knum = 0;
            int len = sb.Length;
            for (int i = 0; i < len; i++)
            {

                if (!((sb[i - knum] > '0' && sb[i - knum] < '9') || sb[i - knum] == '.'))
                {
                    if (sb[i - knum] != ' ')
                    {
                        sb.Insert(i - knum, ' ');
                        break;
                    }
                    else
                    {
                        if (i != 0)
                        {

                            sb.Remove(i - knum, 1);
                            knum++;
                        }
                    }


                }
            }
            return sb.ToString();
        }



        private void buttonX3_Click(object sender, EventArgs e)
        {
            readDocument();
        }

        private void ImportDocumentForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (Directory.Exists("officetemp"))//如果存在就删除officetemp文件夹
            {
                Directory.Delete("officetemp", true);
            }
            quitWord();
        }

        private void buttonX4_Click(object sender, EventArgs e)
        {
            onlyListSave();
        }
    }
}
