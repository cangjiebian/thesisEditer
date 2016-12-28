using System;
using System.Windows.Forms;
using MSWord = Microsoft.Office.Interop.Word;
using System.Threading;
using System.IO;
using System.Xml;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
namespace thesisEditer
{
    public partial class MainForm : Form
    {
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

        ContextMenuStrip cms = new ContextMenuStrip();//右键菜单
        ToolStripMenuItem editThesis = new ToolStripMenuItem("编辑(&E)");
        ToolStripMenuItem editAddproduct = new ToolStripMenuItem("插入目录(&I)");
        ToolStripMenuItem editAddprChildoduct = new ToolStripMenuItem("添加子目录(&A)");
        ToolStripMenuItem editDeleteproduct = new ToolStripMenuItem("删除该目录(&D)");
        ToolStripMenuItem renameproduct = new ToolStripMenuItem("重命名(&M)");
        ToolStripMenuItem clearDoc = new ToolStripMenuItem("清空内容(&C)");

        MSWord.Document readDoc;
        MSWord.Document readTemp;
        MSWord.Application wordApp;
        MSWord.Application readApp;//用于加快控件的反应速度
        EditForm editForm;
        object format = MSWord.WdSaveFormat.wdFormatDocument;//文件类型为word2003，doc
        public ProgressForm progressBarForm;
        /*两个相对路径*/
        public static string rootPath = System.AppDomain.CurrentDomain.BaseDirectory;

        public string officePath = System.AppDomain.CurrentDomain.BaseDirectory + @"\office\";
        public object officeTempName = System.AppDomain.CurrentDomain.BaseDirectory + @"\temp.doc";
        public static string settingsPath = System.AppDomain.CurrentDomain.BaseDirectory + @"\Settings.ini";
        public string nowDocName = "";
        public static string DocAuthor = "无";
        public bool haveDoc = false;
        public DevComponents.AdvTree.Node nowSelectNode = null;//设置当前选择章节
        public string errorText = "";
        object QS = System.Reflection.Missing.Value;//缺省参数
        private DevComponents.AdvTree.Node[] searchNodes = null;
        private string searchText = "";
        private int searchIndex = 0;
        private int nodeCount = 0;
        private int tempNodeCount = 0;//节点数、子节点计数时使用
        private int picIndex = 0, fiPicIndex = 0, tabIndex = 0, fiTabIndex = 0;
        private bool picUpdateIsRuning = false, tabUpdateIsRuning = false;
        public bool CreateThesisIsRuning = false;
        public Thread CreateThesisTh = null;

        public int[] settingsStyleList = new int[11];
        public int[] settingsSizeList = new int[11];
        private string
            mulu_font_style = "黑体",
            list_font_style = "宋体",
            picture_font_style = "宋体",
            subject_font_style = "黑体",
            big_font_style = "黑体",
            small_font_style = "黑体",
            text_font_style = "仿宋",
            page_font_style = "仿宋",
            number_font_style = "Times New Roman",
            reference_font_style = "仿宋",
            intable_font_style = "仿宋",
            table_font_style = "宋体";
        private float
            mulu_font_size = 15f,
            list_font_size = 12f,
            picture_font_size = 10.5f,
            subject_font_size = 16f,
            big_font_size = 12f,
            small_font_size = 12f,
            text_font_size = 12f,
            page_font_size = 10.5f,
            reference_font_size = 10.5f,
            intable_font_size = 10.5f,
            table_font_size = 10.5f;
        private delegate void UpdateAdvNodesNameDelegate(DevComponents.AdvTree.Node node, int qq);
        private delegate void UpdateAdvNodesName_RenameDelegate(DevComponents.AdvTree.Node node);
        private delegate void listSaveDelegate();

        public static string officeVersion;
        public static string thesisEditerVersion = "1.0.0";
        public static int thesisEditerVersionNum = 100;
        public static int serverVersionNum = 0;
        public static string serverVersion = "";
        public static string serverLink = "";
        public static bool hasUpdate = false;
        public MainForm()
        {
            InitializeComponent();
            Init();

        }
        private void Init()
        {
            editThesis.Click += new EventHandler(button3_Click);
            editAddproduct.Click += new EventHandler(button5_Click);
            editAddprChildoduct.Click += new EventHandler(button4_Click);
            editDeleteproduct.Click += new EventHandler(button6_Click);
            renameproduct.Click += new EventHandler(button7_Click);
            clearDoc.Click += new EventHandler(clearDoc_Click);

            cms.Items.Add(editThesis);
            cms.Items.Add(editAddproduct);
            cms.Items.Add(editAddprChildoduct);
            cms.Items.Add(editDeleteproduct);
            cms.Items.Add(renameproduct);
            cms.Items.Add(clearDoc);

            /*杀死word进程*/
            Process[] myProcess = Process.GetProcessesByName("WINWORD");
            updateApp();
            labelItem7.Text += "v" + thesisEditerVersion;
            this.Text += "v" + thesisEditerVersion;
            if (myProcess.Length != 0)
            {
                MessageBox.Show(this, "为了避免程序运行出错，本程序会关闭所有Office Word进程！\n如果你后台打开了word文档，请先保存关闭，避免丢失内容", "关闭提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                foreach (Process p in myProcess)
                {

                    if (p.MainWindowTitle == "")
                    {
                        try
                        {
                            p.Kill();
                        }
                        catch { }

                    }
                }

            }
            readApp = new MSWord.Application();
            readApp.Documents.Add();
            readApp.ActiveWindow.View.ShowFieldCodes = false;
        }
        private void begin()//生成文档
        {

            try
            {
                thesisInfoAndList();
                progressBarForm.Current++;
                thesisBody();//这是生成正文函数
                if (CreateThesisIsRuning)
                {
                    /*以下文档整体处理*/
                    contents();
                    readDoc.Content.WholeStory();
                    readDoc.Content.Font.Name = number_font_style;
                    readDoc.Content.Font.Color = MSWord.WdColor.wdColorBlack;
                    thesisPicture_2();

                    //神秘代码，去除标题前小黑点
                    readDoc.Content.ParagraphFormat.SpaceBeforeAuto = 0;
                    readDoc.Content.ParagraphFormat.SpaceAfterAuto = 0;
                    readDoc.Content.ParagraphFormat.KeepWithNext = 0;
                    readDoc.Content.ParagraphFormat.KeepTogether = 0;
                    //readDoc.Content.ParagraphFormat.WordWrap = 1;
                    progressBarForm.Current++;
                }
                else
                {
                    readDoc.Close(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);
                    wordApp.Quit(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);
                }
            }
            catch (Exception e)
            {
                new ErrorForm(e.ToString()).ShowDialog(this);
            }



        }
        private void start()
        {
            //判断文件是否打开
            foreach (DevComponents.DotNetBar.Office2007Form f in MdiChildren)
            {
                f.Close();
            }

            if (haveDoc)
            {
                MessageBox.Show(this, "请先保存并关闭章节", "警告", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (File.Exists((string)savepath))
            {
                try
                {
                    File.Delete((string)savepath);
                }
                catch
                {
                    MessageBox.Show(this, "请关闭" + savepath + "文件", "覆盖失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            errorText = "";
            listCount += 2;
            progressBarForm = new ProgressForm(this);
            progressBarForm.Max = listCount;
            CreateThesisTh = new Thread(begin);
            CreateThesisTh.Start();
            CreateThesisIsRuning = true;
            if (progressBarForm.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                fileSave(savepath);
            }
            Clipboard.Clear();//清空剪切板

        }
        private bool openThesis()
        {
            foreach (DevComponents.DotNetBar.Office2007Form f in MdiChildren)
            {
                f.Close();
            }
            if (haveDoc == false)
            {
                if (File.Exists((string)savepath))
                {
                    if (editForm == null || editForm.IsDisposed)
                    {
                        editForm = new EditForm(this);
                        editForm.MdiParent = this;
                        editForm.Size = new System.Drawing.Size(0, 0);
                        editForm.WindowState = FormWindowState.Maximized;
                        editForm.Show();

                    }
                    editForm.FilePath = (string)savepath;
                    editForm.Text = savepath.ToString().Split('\\')[savepath.ToString().Split('\\').Length - 1];
                    buttonItem24.Visible = true;
                    buttonItem25.Visible = true;
                    if (editForm.Retry == false)
                    {
                        editForm.Close();
                        return true;
                    }

                    editForm.Update();
                    haveDoc = true;
                    this.labelItem2.Text = "查看与修改";
                    return true;
                }
                else
                {
                    return false;
                }

            }
            return true;
        }
        private void thesisListAdd()//增加目录
        {

            if (haveDoc)
            {
                DialogResult Dr;
                Dr = closeThesis("添加子目录");
                if (Dr == System.Windows.Forms.DialogResult.Cancel)
                {
                    return;
                }
                else
                {
                    foreach (DevComponents.DotNetBar.Office2007Form f in MdiChildren)
                    {
                        f.Close();
                    }
                }
            }
            if (this.advTree_1.SelectedNode != null)
            {


                DevComponents.AdvTree.Node openNode = new DevComponents.AdvTree.Node();
                DevComponents.AdvTree.Node node = new DevComponents.AdvTree.Node();
                node = advTree_1.SelectedNode;

                if (node != null)
                {
                    if (node.Index <= advTree_1.Nodes.Count - 3 && node.Index >= 4 || node.Level != 0)
                    {

                        openNode.Text = "子目录";
                        node.Nodes.Add(openNode);
                        node.Expand();
                        UpdateAdvNodesName(null);
                        listSave();
                        node.Nodes[node.Nodes.Count - 1].BeginEdit();
                    }
                }
            }
        }
        private void thesisListInsert()//插入目录
        {
            DialogResult Dr;
            if (haveDoc)
            {
                Dr = closeThesis("插入目录");
                if (Dr == System.Windows.Forms.DialogResult.Cancel)
                {
                    return;
                }
                else
                {
                    foreach (DevComponents.DotNetBar.Office2007Form f in MdiChildren)
                    {
                        f.Close();
                    }
                }
            }
            if (this.advTree_1.SelectedNode != null)
            {

                DevComponents.AdvTree.Node openNode = new DevComponents.AdvTree.Node();
                DevComponents.AdvTree.Node node = new DevComponents.AdvTree.Node();
                node = advTree_1.SelectedNode;
                if (node != null)
                {
                    if (node.Parent != null)
                    {
                        openNode.Text = "子目录";
                        node.Parent.Nodes.Insert(node.Index + 1, openNode);

                    }
                    else
                    {
                        if (node.Index <= advTree_1.Nodes.Count - 3 && node.Index >= 3)
                        {
                            openNode.Text = "目录";
                            advTree_1.Nodes.Insert(node.Index + 1, openNode);
                        }

                    }
                    UpdateAdvNodesName(null);
                    listSave();
                    node.NextNode.BeginEdit();
                }
                else
                {
                    MessageBox.Show(this, "请先选择要插入的目录位置!", "选择位置", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
        }
        private void thesisListDelete()//删除目录
        {
            DialogResult Dr;
            if (haveDoc)
            {
                Dr = closeThesis("删除目录");
                if (Dr == System.Windows.Forms.DialogResult.Cancel)
                {
                    return;
                }
                else
                {
                    foreach (DevComponents.DotNetBar.Office2007Form f in MdiChildren)
                    {
                        f.Close();
                    }
                }
            }
            if (this.advTree_1.SelectedNode != null)
            {


                DevComponents.AdvTree.Node node = new DevComponents.AdvTree.Node();
                node = advTree_1.SelectedNode;
                if (node != null)
                {
                    if (node.Index <= advTree_1.Nodes.Count - 3 && node.Index >= 4 || node.Level != 0)
                    {
                        if (node.Nodes.Count == 0)
                        {
                            if (File.Exists(officePath + node.Text + ".doc"))
                            {
                                if (MessageBox.Show(this, "该标题内含有正文内容,删除后无法复原!确定删除吗?", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
                                {
                                    try
                                    {
                                        File.Delete(officePath + node.Text + ".doc");
                                        node.Remove();
                                    }
                                    catch
                                    {
                                        MessageBox.Show(this, "删除失败!", "无法删除", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                }
                            }
                            else
                            {
                                node.Remove();
                            }


                        }
                        else
                        {
                            MessageBox.Show(this, "请先删除该目录下的子目录!", "失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                else
                {
                    MessageBox.Show(this, "请先选择删除的目录!", "请选择", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                UpdateAdvNodesName(null);
                listSave();
            }
        }


        /*以下为函数为运算使用*/

        private void fileSave(object savepath)
        {

            try
            {
                readApp.ActiveWindow.View.ShowFieldCodes = false;
                readDoc.SaveAs(ref savepath, ref format, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS);
                readDoc.Close(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);
                wordApp.Quit(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);
                if (errorText == "")
                {
                    MessageBox.Show(this, "文档" + savepath + "已生成!\n接下来请对毕业论文查看与修改，是否合学校规范", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(this, "文档" + savepath + "已生成!\n接下来请对毕业论文查看与修改，是否合学校规范\n异常章节\n" + errorText, "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                openThesis();
            }
            catch
            {
                MessageBox.Show(this, "文档" + savepath + "生成失败！\n请重新生成或重启本软件", "失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /*文章编辑函数*/
        private void thesisEdit()
        {
            DevComponents.AdvTree.Node node = new DevComponents.AdvTree.Node();
            node = advTree_1.SelectedNode;
            if (node != null)
            {
                if (node.Text == nowDocName)//判断该文档已经打开就跳出方法
                    return;
                nowDocName = node.Text;
                this.labelItem2.Text = nowDocName;
                if (node.Level == 0 && node.Index == advTree_1.Nodes.Count - 2)//参考文献界面
                {
                    foreach (DevComponents.DotNetBar.Office2007Form f in MdiChildren)
                    {
                        f.Close();
                    }
                    DocumentEditForm docForm = new DocumentEditForm();
                    docForm.Text = node.Text;
                    docForm.MdiParent = this;
                    docForm.Size = new System.Drawing.Size(0, 0);
                    docForm.WindowState = FormWindowState.Maximized;
                    docForm.Show();

                    docForm.Update();
                }
                else if (node.Level == 0 && (node.Index == 1 || node.Index == 3))//关键词界面
                {
                    foreach (DevComponents.DotNetBar.Office2007Form f in MdiChildren)
                    {
                        //if (!(f is WordsEditForm))
                        f.Close();
                    }
                    WordsEditForm wordsForm = new WordsEditForm();
                    wordsForm.MdiParent = this;
                    wordsForm.Size = new System.Drawing.Size(0, 0);
                    wordsForm.WindowState = FormWindowState.Maximized;
                    wordsForm.Show();
                    wordsForm.Update();
                }
                else//普通文档界面
                {
                    foreach (DevComponents.DotNetBar.Office2007Form f in MdiChildren)
                    {
                        if (!(f is EditForm))
                            f.Close();
                    }
                    if (editForm == null || editForm.IsDisposed)
                    {
                        editForm = new EditForm(this);
                        editForm.MdiParent = this;
                        editForm.Size = new System.Drawing.Size(0, 0);
                        editForm.WindowState = FormWindowState.Maximized;
                        editForm.Show();
                    }
                    if (File.Exists(officePath + node.Text + ".doc"))
                    {


                        editForm.Text = node.Text;
                        editForm.FilePath = officePath + node.Text + ".doc";
                        if (editForm.Retry == false)
                        {
                            editForm.Close();
                            return;
                        }
                        editForm.Update();
                    }
                    else
                    {
                        MSWord.Document newDoc = new MSWord.Document();

                        object path = officePath + node.Text + ".doc";
                        newDoc.SaveAs(ref path, ref format, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS);
                        newDoc.Close(ref QS, ref QS, ref QS);
                        editForm.Text = node.Text;
                        editForm.FilePath = officePath + node.Text + ".doc";
                        editForm.Update();
                    }
                }


                this.advTree_1.DragDropEnabled = false;
                this.advTree_1.CellEdit = false;
                haveDoc = true;
                nowSelectNode = node;
                if (node.Level == 0 && node.Index == 0)
                {
                    this.buttonItem7.Enabled = true;
                }
                else if (node.Level == 0 && node.Index == this.advTree_1.Nodes.Count - 1)
                {
                    this.buttonItem5.Enabled = true;
                }
                else
                {
                    this.buttonItem5.Enabled = true;
                    this.buttonItem7.Enabled = true;
                }
                if (node.Level != 0 || (node.Index > 3 && node.Index < this.advTree_1.Nodes.Count - 2))
                {
                    this.buttonItem12.Enabled = true;
                    this.buttonItem13.Enabled = true;
                }
                buttonItem24.Visible = false;
                buttonItem25.Visible = false;
            }
        }


        /*结尾加空格,用于对齐缩进*/
        private void thesisIndentation()
        {
            MSWord.Paragraphs dl;
            dl = readDoc.Paragraphs;
            foreach (MSWord.Paragraph d in dl)
            {

                readDoc.Range(d.Range.Start, d.Range.End - 1).InsertAfter("                                    ");
            }
        }


        /*删除结尾空行,和项目符号,调用前readTemp对象要有内容*/
        private void nullLine()
        {
            /*结尾空行处理*/
            char[] ch;
            bool del = true;

            while (del)
            {
                if (readTemp.Paragraphs.Count == 1)
                {
                    break;
                }
                ch = readTemp.Paragraphs.Last.Range.Text.ToCharArray();
                foreach (char s in ch)
                {
                    if ((int)s != 13 && (int)s != 10 && (int)s != 9 && (int)s != 32)
                    {
                        del = false;
                        break;
                    }
                }
                if (del)
                {
                    readTemp.Paragraphs.Last.Range.Delete();

                }
                if (readTemp.Paragraphs.Count == 1)//只有一段
                {
                    break;
                }
                if (readTemp.Content.Tables.Count != 0)//有表格
                {
                    break;
                }
            }

            /*项目符号问题*/
            wordApp.Selection.WholeStory();
            wordApp.Selection.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(float.Parse("0"));


        }


        /*处理图片缩进,调用前readTemp对象要有内容*/
        private void thesisPicture()
        {
            /*图片问题,图片居中处理*/
            MSWord.InlineShapes shapes;
            shapes = readTemp.InlineShapes;
            MSWord.Paragraphs dl;
            foreach (MSWord.InlineShape sp in shapes)
            {
                sp.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                sp.Range.ParagraphFormat.FirstLineIndent = wordApp.CentimetersToPoints(0);//删除图片的缩进
                sp.Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                sp.Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceSingle;
                try
                {
                    if (sp.Range.Paragraphs[1].Next().Range.ParagraphFormat.Alignment == MSWord.WdParagraphAlignment.wdAlignParagraphCenter || sp.Range.Paragraphs[1].Next().Range.Text.Length < 25)
                    {
                        //找到图题
                        sp.Range.Paragraphs.First.Next().Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                        sp.Range.Paragraphs.First.Next().Range.Font.Name = picture_font_style;
                        sp.Range.Paragraphs.First.Next().Range.Font.Size = picture_font_size;
                    }
                }
                catch (Exception e)
                {

                }

            }
            /*所有居中段落的首行不缩进*/
            dl = readTemp.Paragraphs;
            thesisIndent(dl);

        }
        /// <summary>
        /// 首行缩进（居中段落除外）
        /// </summary>
        private void thesisIndent(MSWord.Paragraphs dl)
        {
            foreach (MSWord.Paragraph d in dl)
            {
                if (d.Range.ParagraphFormat.Alignment == MSWord.WdParagraphAlignment.wdAlignParagraphCenter)
                {
                    d.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                    d.Range.ParagraphFormat.FirstLineIndent = wordApp.CentimetersToPoints(0);//删除所有居中段落的首行缩进
                }
                else
                {
                    d.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 2;
                    d.Range.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(0);
                    d.Range.ParagraphFormat.RightIndent = wordApp.CentimetersToPoints(0);
                    d.Range.ParagraphFormat.CharacterUnitLeftIndent = 0;
                    d.Range.ParagraphFormat.CharacterUnitRightIndent = 0;
                }
            }
        }
        private void thesisTable()//取消表格缩进
        {

            MSWord.Tables tabs = readTemp.Content.Tables;
            if (tabs.Count != 0)
            {
                foreach (MSWord.Table tab in tabs)
                {

                    tab.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                    tab.Range.ParagraphFormat.FirstLineIndent = wordApp.CentimetersToPoints(0);
                    tab.Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                    tab.Rows.Alignment = MSWord.WdRowAlignment.wdAlignRowCenter;
                    tab.Range.Font.Size = intable_font_size;
                    tab.Range.Font.Name = intable_font_style;

                    if (tab.Range.Paragraphs.First.Previous() != null && tab.Range.Paragraphs.First.Previous().Range.ParagraphFormat.Alignment == MSWord.WdParagraphAlignment.wdAlignParagraphCenter)
                    {
                        //找到表序
                        tab.Range.Paragraphs.First.Previous().Range.Font.Name = table_font_style;
                        tab.Range.Paragraphs.First.Previous().Range.Font.Size = table_font_size;

                    }

                    /*缩进内容为表序处理，表上方无文字则会出现非常大的bug*/
                }
            }

        }
        /*行间距处理*/
        private void verticalSpacing()
        {
            /*行间距22固定值，由于图片必须单倍行间距，图片处理函数会修改行间距为单倍*/
            readTemp.Content.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceExactly;
            readTemp.Content.ParagraphFormat.LineSpacing = 22F;
        }
        /*论文目录函数*/
        private void contents()
        {

            MSWord.Range range = readDoc.GoTo(MSWord.WdGoToItem.wdGoToPage, MSWord.WdGoToDirection.wdGoToAbsolute, QS, 4 + addNum);
            readDoc.TablesOfContents.Add(range, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, true, ref QS, ref QS);
            readDoc.TablesOfContents[1].Range.Font.Name = list_font_style;
            readDoc.TablesOfContents[1].Range.Font.Size = list_font_size;
            readDoc.TablesOfContents[1].Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceExactly;
            readDoc.TablesOfContents[1].Range.ParagraphFormat.LineSpacing = 22F;
            range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
            range.Text = "\n目  录\n\n";
            range.Font.Name = mulu_font_style;
            range.Font.Size = mulu_font_size;
            range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceExactly;
            range.ParagraphFormat.LineSpacing = 22F;
        }
        /*结尾处理图片图号同一页要求，暂时未加入进度条中*/
        private void thesisPicture_2()
        {
            int picturePage, nextPage;
            MSWord.InlineShapes shapes;
            MSWord.Paragraph par;
            shapes = readDoc.InlineShapes;
            foreach (MSWord.InlineShape sp in shapes)
            {
                picturePage = (int)sp.Range.get_Information(MSWord.WdInformation.wdActiveEndPageNumber);
                if (sp.Range.Paragraphs[1].Next().Range.ParagraphFormat.Alignment == MSWord.WdParagraphAlignment.wdAlignParagraphCenter || sp.Range.Paragraphs[1].Next().Range.Text.Length < 25)
                {
                    //找到图题
                    nextPage = (int)sp.Range.Paragraphs.First.Next().Range.get_Information(MSWord.WdInformation.wdActiveEndPageNumber);
                    if (picturePage != nextPage)
                    {
                        sp.Range.Paragraphs.Add(readDoc.Range(sp.Range.Paragraphs.First.Range.Start));
                        sp.Range.Paragraphs.First.Previous().Range.InsertBreak(MSWord.WdBreakType.wdPageBreak);
                    }
                }

            }
        }
        /*页码函数1*/
        private void pageNumber_1()
        {
            object oAlignment = Microsoft.Office.Interop.Word.WdPageNumberAlignment.wdAlignPageNumberCenter;

            object oFirstPage = true;
            wordApp.Selection.GoTo(MSWord.WdGoToItem.wdGoToPage, MSWord.WdGoToDirection.wdGoToAbsolute, QS, 6 + addNum);
            wordApp.Selection.Sections[1].Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.Add(oAlignment, oFirstPage);

            wordApp.Selection.GoTo(MSWord.WdGoToItem.wdGoToPage, MSWord.WdGoToDirection.wdGoToAbsolute, QS, 5 + addNum);//跳转第四页
            /*插入页码*/
            wordApp.Selection.Sections[1].Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
            wordApp.Selection.Sections[1].Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.NumberStyle = MSWord.WdPageNumberStyle.wdPageNumberStyleUppercaseRoman;
            wordApp.Selection.Sections[1].Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.RestartNumberingAtSection = true;
            wordApp.Selection.Sections[1].Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.StartingNumber = 1;
            wordApp.Selection.Sections[1].Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font.Name = page_font_style;
            wordApp.Selection.Sections[1].Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font.Size = page_font_size;
            wordApp.Selection.GoTo(MSWord.WdGoToItem.wdGoToPage, MSWord.WdGoToDirection.wdGoToNext, QS, 1);//跳转第一页
            /*删除前几页页码*/
            wordApp.Selection.Sections[1].Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers[1].Cut();
            /*删除空白页*/
            MSWord.Range range1 = wordApp.Selection.GoTo(MSWord.WdGoToItem.wdGoToPage, MSWord.WdGoToDirection.wdGoToAbsolute, QS, 5 + addNum);
            MSWord.Range range2 = wordApp.Selection.GoTo(MSWord.WdGoToItem.wdGoToPage, MSWord.WdGoToDirection.wdGoToAbsolute, QS, 6 + addNum);
            readDoc.Range(range1.Start, range2.Start).Delete();
        }
        /*页码函数2*/
        private void pageNumber_2()
        {
            readDoc.Paragraphs.Last.Range.Select();
            wordApp.Selection.InsertBreak(MSWord.WdBreakType.wdSectionBreakNextPage);
            wordApp.Selection.Sections[1].Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.NumberStyle = MSWord.WdPageNumberStyle.wdPageNumberStyleArabic;
        }
        /*更新目录标号*/
        //参数qq是表示是否更新目录所指的文档名qq=0则更新,否则不更新.
        int listCount = 0;

        private void UpdateAdvNodesName(DevComponents.AdvTree.Node node, int qq = 0)
        {
            string text;


            if (node == null)
            {
                listCount = 0;
                nodeCount = 0;
                foreach (DevComponents.AdvTree.Node k in advTree_1.Nodes)
                {
                    nodeCount++;
                    if (k.Index >= 4 && k.Index <= advTree_1.Nodes.Count - 3)
                    {
                        text = k.Text;
                        if (text.Split(' ')[0] != text && (text[0] > '0' && text[0] <= '9'))
                            k.Text = k.Text.Remove(0, text.Split(' ')[0].Length + 1);
                        if (k.Name != (k.Index - 3).ToString() + " " + k.Text && qq == 0)
                        {
                            if (File.Exists(officePath + k.Name + ".doc") && File.Exists(officePath + (k.Index - 3).ToString() + " " + k.Text + ".doc") == false)
                            {

                                File.Move(officePath + k.Name + ".doc", officePath + (k.Index - 3).ToString() + " " + k.Text + ".doc");
                            }
                        }
                        k.Name = (k.Index - 3).ToString() + " " + k.Text;
                        k.Text = (k.Index - 3).ToString() + " " + k.Text;
                    }

                    if (k.Nodes.Count != 0)
                    {
                        UpdateAdvNodesName(k, qq);
                    }
                    else
                    {
                        listCount++;
                    }
                }
            }
            else
            {
                foreach (DevComponents.AdvTree.Node k in node.Nodes)
                {
                    nodeCount++;
                    text = k.Text;
                    if (text.Split(' ')[0] != text && (text[0] > '0' && text[0] <= '9'))
                        k.Text = k.Text.Remove(0, text.Split(' ')[0].Length + 1);
                    if (k.Name != k.Parent.Name.Split(' ')[0] + "." + (k.Index + 1).ToString() + " " + k.Text && qq == 0)
                    {
                        if (File.Exists(officePath + k.Name + ".doc") && File.Exists(officePath + k.Parent.Name.Split(' ')[0] + "." + (k.Index + 1).ToString() + " " + k.Text + ".doc") == false)
                        {

                            File.Move(officePath + k.Name + ".doc", officePath + k.Parent.Name.Split(' ')[0] + "." + (k.Index + 1).ToString() + " " + k.Text + ".doc");
                        }
                    }
                    k.Name = k.Parent.Name.Split(' ')[0] + "." + (k.Index + 1).ToString() + " " + k.Text;
                    k.Text = k.Parent.Name.Split(' ')[0] + "." + (k.Index + 1).ToString() + " " + k.Text;
                    if (k.Nodes.Count != 0)
                    {
                        UpdateAdvNodesName(k, qq);
                    }
                    else
                    {
                        listCount++;
                    }
                }
            };
        }
        private void UpdateAdvNodesName_Rename(DevComponents.AdvTree.Node node = null)
        {
            if (node.Parent == null)
            {
                if (File.Exists(officePath + node.Name + ".doc") && File.Exists(officePath + (node.Index - 3).ToString() + " " + node.Text + ".doc") == false)
                {

                    File.Move(officePath + node.Name + ".doc", officePath + (node.Index - 3).ToString() + " " + node.Text + ".doc");
                }
                node.Name = (node.Index - 3).ToString() + " " + node.Text.Trim();
                node.Text = (node.Index - 3).ToString() + " " + node.Text.Trim();
            }
            else
            {
                if (File.Exists(officePath + node.Name + ".doc") && File.Exists(officePath + node.Parent.Name.Split(' ')[0] + "." + (node.Index + 1).ToString() + " " + node.Text + ".doc") == false)
                {

                    File.Move(officePath + node.Name + ".doc", officePath + node.Parent.Name.Split(' ')[0] + "." + (node.Index + 1).ToString() + " " + node.Text + ".doc");
                }
                node.Name = node.Parent.Name.Split(' ')[0] + "." + (node.Index + 1).ToString() + " " + node.Text.Trim();
                node.Text = node.Parent.Name.Split(' ')[0] + "." + (node.Index + 1).ToString() + " " + node.Text.Trim();
            }
        }
        XmlDocument xmlDoc;
        XmlNode xmlNode;
        XmlElement xmlEle;

        private void listSave()
        {
            Control nowControl = this.textBox_xnumber;
            xmlDoc = new XmlDocument();
            XmlDeclaration xmlDec = xmlDoc.CreateXmlDeclaration("1.0", "gb2312", null);
            xmlDoc.AppendChild(xmlDec);
            xmlEle = xmlDoc.CreateElement("Root");
            xmlDoc.AppendChild(xmlEle);
            xmlNode = xmlDoc.SelectSingleNode("Root");
            xmlEle = xmlDoc.CreateElement("TreeNode");
            xmlNode.AppendChild(xmlEle);
            ergodicNodeSave(null, null);
            xmlNode = xmlDoc.SelectSingleNode("Root");
            xmlEle = xmlDoc.CreateElement("Infos");
            xmlNode.AppendChild(xmlEle);
            string[] blList = new string[12];
            blList[0] = this.textBox_xnumber.Text;
            blList[1] = this.textBox_LWname.Text;
            blList[2] = this.textBox_english.Text;
            blList[3] = this.textBox_name.Text;
            blList[4] = this.textBox_number.Text;
            blList[5] = this.textBox_xib.Text;
            blList[6] = this.textBox_zhuany.Text;
            blList[7] = this.textBox_teacher.Text;
            blList[8] = this.comboBox1.Text;
            blList[9] = this.dateTimePicker1.Text;
            blList[10] = this.dateTimePicker2.Text;
            blList[11] = this.dateTimePicker3.Text;
            for (int i = 0; i < 12; i++)
            {

                xmlNode = xmlDoc.SelectSingleNode("Root");
                xmlNode = xmlNode.SelectSingleNode("Infos");
                xmlEle = xmlDoc.CreateElement("Info-" + (i + 1).ToString());
                xmlEle.SetAttribute("Name", blList[i]);
                xmlNode.AppendChild(xmlEle);


            }

            xmlDoc.Save(officePath + "NodeInfo.xml");
        }
        private void listRead()
        {
            advTree_1.Nodes.Clear();
            string[] blList = new string[12];
            int i = 0;
            xmlDoc = new XmlDocument();


            try
            {
                xmlDoc.Load(officePath + "NodeInfo.xml");
            }
            catch
            {

                DevComponents.AdvTree.Node[] nodeList = new DevComponents.AdvTree.Node[6];
                nodeList[0] = new DevComponents.AdvTree.Node();
                nodeList[1] = new DevComponents.AdvTree.Node();
                nodeList[2] = new DevComponents.AdvTree.Node();
                nodeList[3] = new DevComponents.AdvTree.Node();
                nodeList[4] = new DevComponents.AdvTree.Node();
                nodeList[5] = new DevComponents.AdvTree.Node();
                nodeList[0].Text = "摘  要";
                nodeList[1].Text = "关键词";
                nodeList[2].Text = "Abstract";
                nodeList[3].Text = "Key words";
                nodeList[4].Text = "参考文献";
                nodeList[5].Text = "致  谢";
                this.advTree_1.Nodes.AddRange(nodeList);
                listSave();
                return;
            }
            this.advTree_1.Nodes.Clear();

            ergodicNodeRead(null, null);
            xmlNode = xmlDoc.SelectSingleNode("Root");
            xmlNode = xmlNode.SelectSingleNode("Infos");
            foreach (XmlNode xmlnode in xmlNode.ChildNodes)
            {
                blList[i] = xmlnode.Attributes["Name"].Value;
                i++;

            }
            this.textBox_xnumber.Text = blList[0];
            this.textBox_LWname.Text = blList[1];
            this.textBox_english.Text = blList[2];
            this.textBox_name.Text = blList[3];
            this.textBox_number.Text = blList[4];
            this.textBox_xib.Text = blList[5];
            this.textBox_zhuany.Text = blList[6];
            this.textBox_teacher.Text = blList[7];
            this.comboBox1.Text = blList[8];
            this.dateTimePicker1.Text = blList[9];
            this.dateTimePicker2.Text = blList[10];
            this.dateTimePicker3.Text = blList[11];




        }
        private void ergodicNodeSave(DevComponents.AdvTree.Node Node, XmlNode xml)
        {
            if (Node == null)
            {
                foreach (DevComponents.AdvTree.Node nowNode in advTree_1.Nodes)
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
        private void ergodicNodeRead(DevComponents.AdvTree.Node Node, XmlNode xml)
        {


            int index;
            if (xml == null)
            {
                xmlNode = xmlDoc.SelectSingleNode("Root");
                xmlNode = xmlNode.SelectSingleNode("TreeNode");
                foreach (XmlNode xmlnode in xmlNode.ChildNodes)
                {
                    DevComponents.AdvTree.Node node = new DevComponents.AdvTree.Node();
                    node.Text = xmlnode.Attributes["Name"].Value;
                    index = this.advTree_1.Nodes.Add(node);
                    if (xmlnode.HasChildNodes)
                    {
                        ergodicNodeRead(this.advTree_1.Nodes[this.advTree_1.Nodes.Count - 1], xmlnode);
                    }
                }
            }
            else
            {

                foreach (XmlNode xmlnode in xml.ChildNodes)
                {
                    DevComponents.AdvTree.Node node = new DevComponents.AdvTree.Node();
                    node.Text = xmlnode.Attributes["Name"].Value;
                    index = Node.Nodes.Add(node);
                    if (xmlnode.HasChildNodes)
                    {
                        ergodicNodeRead(Node.Nodes[Node.Nodes.Count - 1], xmlnode);
                    }
                }
            }
        }
        /// <summary>
        /// 查找替换函数
        /// </summary>
        /// <param name="find">查找的字符串</param>
        /// <param name="repl">替换的字符串</param>
        private void replace(object find, object repl)
        {
            object QS = System.Reflection.Missing.Value;//缺省参数
            object Replace = MSWord.WdReplace.wdReplaceAll;
            readDoc.Content.Find.ClearFormatting();
            readDoc.Content.Find.Execute(ref find, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref repl, ref Replace, ref QS, ref QS, ref QS, ref QS);


        }


        /*通过书签插入内容函数(带下划线,加粗,缩进)*/
        private void insertText(string book, string text, string suoj, string suoy, int cu)
        {
            //wordApp.Selection.Font.Name = "Times New Roman"
            object QS = System.Reflection.Missing.Value;//缺省参数
            object what = MSWord.WdGoToItem.wdGoToBookmark;
            object obook = book;
            string[] str;
            int i, j;
            readDoc.ActiveWindow.Selection.GoTo(ref what, ref QS, ref QS, ref obook);
            wordApp.Selection.Font.Underline = MSWord.WdUnderline.wdUnderlineSingle;

            if (suoj == "0")
            {
                wordApp.Selection.ParagraphFormat.RightIndent = wordApp.CentimetersToPoints(float.Parse(suoy));
                wordApp.Selection.Font.Bold = cu;
                readDoc.ActiveWindow.Selection.TypeText(text);


            }
            else if (suoj == "5.4")
            {
                str = text.Split('\n');
                //wordApp.Selection.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(float.Parse("0"));
                for (i = 0; i < str.Length; i++)
                {
                    wordApp.Selection.Font.Bold = cu;
                    wordApp.Selection.ParagraphFormat.RightIndent = wordApp.CentimetersToPoints(float.Parse(suoy));
                    readDoc.ActiveWindow.Selection.TypeText(str[i] + "\n");
                    wordApp.Selection.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(float.Parse(suoj));

                }
                if (i < 4)
                {
                    for (j = 0; j < 4 - i; j++)
                    {
                        wordApp.Selection.ParagraphFormat.RightIndent = wordApp.CentimetersToPoints(float.Parse(suoy));
                        wordApp.Selection.Font.Bold = cu;
                        readDoc.ActiveWindow.Selection.TypeText("                                  \n");

                    }

                }
                readDoc.ActiveWindow.Selection.Delete();
            }
            else if (suoj == "2.45")
            {
                str = text.Split('\n');

                for (i = 0; i < str.Length; i++)
                {

                    wordApp.Selection.Font.Bold = cu;
                    wordApp.Selection.ParagraphFormat.RightIndent = wordApp.CentimetersToPoints(float.Parse(suoy));
                    readDoc.ActiveWindow.Selection.TypeText(str[i] + "\n");
                    wordApp.Selection.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(float.Parse(suoj));

                }
                if (i < 2)
                {
                    for (j = 0; j < 2 - i; j++)
                    {
                        wordApp.Selection.ParagraphFormat.RightIndent = wordApp.CentimetersToPoints(float.Parse(suoy));
                        wordApp.Selection.Font.Bold = cu;
                        readDoc.ActiveWindow.Selection.TypeText("                                  \n");

                    }

                }
                readDoc.ActiveWindow.Selection.Delete();

            }


        }

        /*以下是论文编写的函数*/


        /*写论文正文函数*/
        object tempPath;
        private void thesisBody(DevComponents.AdvTree.Node node = null)
        {
            if (CreateThesisIsRuning == false)
                return;
            if (node == null)
            {
                int num = 1;
                int numq = 1;
                foreach (DevComponents.AdvTree.Node k in advTree_1.Nodes)
                {
                    if (CreateThesisIsRuning == false)
                        return;
                    if (k.Index == advTree_1.Nodes.Count - 2)//参考文献正文
                    {
                        readDoc.Paragraphs.Last.Range.InsertBreak(MSWord.WdBreakType.wdPageBreak);
                        readDoc.Paragraphs.Add();
                        tempPath = officePath + k.Text + ".doc";
                        readDoc.Paragraphs.Last.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                        readDoc.Paragraphs.Last.Range.ParagraphFormat.FirstLineIndent = wordApp.CentimetersToPoints(0);
                        readDoc.Paragraphs.Last.set_Style(MSWord.WdBuiltinStyle.wdStyleHeading1);
                        readDoc.Paragraphs.Last.Previous().set_Style(MSWord.WdBuiltinStyle.wdStyleHtmlNormal);
                        readDoc.Paragraphs.Last.Range.Font.Name = big_font_style;
                        readDoc.Paragraphs.Last.Range.Font.Size = big_font_size;//小四
                        readDoc.Paragraphs.Last.Range.Bold = 0;
                        readDoc.Paragraphs.Last.Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                        readDoc.Paragraphs.Last.Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceExactly;
                        readDoc.Paragraphs.Last.Range.ParagraphFormat.LineSpacing = 22F;
                        readDoc.Paragraphs.Last.Range.Text = k.Text + "\n";
                        try
                        {

                            readTemp = wordApp.Documents.Open(ref tempPath, ref QS, true, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS);
                            wordApp.Visible = false;
                            //设置正文格式
                            nullLine();
                            verticalSpacing();
                            readTemp.Content.Font.Size = reference_font_size;
                            readTemp.Content.Font.Name = reference_font_style;
                            readTemp.Content.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                            readTemp.Content.ParagraphFormat.FirstLineIndent = wordApp.CentimetersToPoints(0);
                            readTemp.Content.Font.Bold = 0;
                            thesisPicture();
                            thesisTable();
                            readTemp.Content.Copy();
                            
                            readDoc.Paragraphs.Last.Range.PasteAndFormat(MSWord.WdRecoveryType.wdFormatOriginalFormatting);
                            
                            
                            readTemp.Close(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);
                            readDoc.Paragraphs.Last.set_Style(MSWord.WdBuiltinStyle.wdStyleNormal);
                            readDoc.Paragraphs.Last.Range.Text = "\n";

                        }
                        catch
                        {
                            errorText = errorText + k.Text + "\n";
                            try
                            {
                                readTemp.Close(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);
                            }
                            catch { }
                        }
                        progressBarForm.Current++;
                    }
                    else if (k.Index == advTree_1.Nodes.Count - 1)//致谢正文
                    {
                        readDoc.Paragraphs.Last.Range.InsertBreak(MSWord.WdBreakType.wdPageBreak);
                        readDoc.Paragraphs.Add();
                        tempPath = officePath + k.Text + ".doc";
                        readDoc.Paragraphs.Last.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                        readDoc.Paragraphs.Last.Range.ParagraphFormat.FirstLineIndent = wordApp.CentimetersToPoints(0);
                        readDoc.Paragraphs.Last.set_Style(MSWord.WdBuiltinStyle.wdStyleHeading1);
                        readDoc.Paragraphs.Last.Range.Font.Name = big_font_style;
                        readDoc.Paragraphs.Last.Range.Font.Size = big_font_size;//小四
                        readDoc.Paragraphs.Last.Range.Bold = 0;
                        readDoc.Paragraphs.Last.Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                        readDoc.Paragraphs.Last.Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceExactly;
                        readDoc.Paragraphs.Last.Range.ParagraphFormat.LineSpacing = 22F;
                        readDoc.Paragraphs.Last.Range.Text = k.Text + "\n";
                        try
                        {

                            readTemp = wordApp.Documents.Open(ref tempPath, ref QS, true, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS);
                            wordApp.Visible = false;
                            //设置正文格式
                            nullLine();
                            verticalSpacing();


                            readTemp.Content.Font.Size = text_font_size;
                            readTemp.Content.Font.Name = text_font_style;
                            readTemp.Content.Font.Bold = 0;
                            thesisPicture();
                            thesisTable();
                            readTemp.Content.Copy();
                            
                            readDoc.Paragraphs.Last.Range.PasteAndFormat(MSWord.WdRecoveryType.wdFormatOriginalFormatting);
                            
                            
                            readTemp.Close(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);
                            readDoc.Paragraphs.Last.set_Style(MSWord.WdBuiltinStyle.wdStyleNormal);
                            readDoc.Paragraphs.Last.Range.Text = "\n";

                        }
                        catch
                        {
                            errorText = errorText + k.Text + "\n";
                            try
                            {
                                readTemp.Close(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);
                            }
                            catch { }
                        }
                        progressBarForm.Current++;
                    }
                    else if (k.Index >= 4 && k.Index < advTree_1.Nodes.Count - 2)//中间正文项,第一级标题
                    {
                        tempPath = officePath + k.Text + ".doc";
                        readDoc.Paragraphs.Last.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                        readDoc.Paragraphs.Last.Range.ParagraphFormat.FirstLineIndent = wordApp.CentimetersToPoints(0);
                        readDoc.Paragraphs.Last.set_Style(MSWord.WdBuiltinStyle.wdStyleHeading1);
                        readDoc.Paragraphs.Last.Range.Font.Name = big_font_style;
                        readDoc.Paragraphs.Last.Range.Font.Size = big_font_size;
                        readDoc.Paragraphs.Last.Range.Bold = 0;
                        readDoc.Paragraphs.Last.Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceExactly;
                        readDoc.Paragraphs.Last.Range.ParagraphFormat.LineSpacing = 22F;
                        readDoc.Paragraphs.Last.Range.Text = k.Text + "\n";

                        if (k.Nodes.Count != 0)
                        {
                            try
                            {
                                readTemp = wordApp.Documents.Open(ref tempPath, ref QS, true, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS);
                                wordApp.Visible = false;
                                //设置正文格式
                                nullLine();
                                verticalSpacing();

                                if (readTemp.Content.Text.Trim() != "")
                                {
                                    //readTemp.Content.set_Style(MSWord.WdBuiltinStyle.wdStyleNormal);//设置成默认格式，会消除项目符号
                                    readTemp.Content.Font.Size = text_font_size;
                                    readTemp.Content.Font.Name = text_font_style;
                                    readTemp.Content.Font.Bold = 0;
                                    thesisPicture();
                                    thesisTable();
                                    readTemp.Content.Copy();

                                    
                                    readDoc.Paragraphs.Last.Range.PasteAndFormat(MSWord.WdRecoveryType.wdFormatOriginalFormatting);
                                    
                                    
                                }
                                readTemp.Close(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);
                            }
                            catch
                            {
                                try
                                {
                                    readTemp.Close(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);
                                }
                                catch { }
                            }
                            thesisBody(k);
                        }
                        else
                        {
                            try
                            {
                                readTemp = wordApp.Documents.Open(ref tempPath, ref QS, true, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS);
                                wordApp.Visible = false;
                                //设置正文格式
                                nullLine();
                                verticalSpacing();
                                //readTemp.Content.set_Style(MSWord.WdBuiltinStyle.wdStyleNormal);
                                readTemp.Content.Font.Size = text_font_size;
                                readTemp.Content.Font.Name = text_font_style;
                                readTemp.Content.Font.Bold = 0;
                                thesisPicture();
                                thesisTable();
                                readTemp.Content.Copy();
                                
                                readDoc.Paragraphs.Last.Range.PasteAndFormat(MSWord.WdRecoveryType.wdFormatOriginalFormatting);
                                
                                
                                readTemp.Close(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);


                            }
                            catch
                            {
                                errorText = errorText + k.Text + "\n";
                                try
                                {
                                    readTemp.Close(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);
                                }
                                catch { }
                            }
                            progressBarForm.Current++;
                        }
                        num++;
                    }
                    else if (k.Index < 4)//前四项根标题,要插入论文名字
                    {
                        tempPath = officePath + k.Text + ".doc";
                        if (numq == 1)
                        {

                            readDoc.Paragraphs.Last.Range.set_Style(MSWord.WdBuiltinStyle.wdStyleNormal);
                            readDoc.Paragraphs.Last.Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                            readDoc.Paragraphs.Last.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                            readDoc.Paragraphs.Last.Range.ParagraphFormat.FirstLineIndent = wordApp.CentimetersToPoints(0);
                            readDoc.Paragraphs.Last.Range.Font.Name = subject_font_style;
                            readDoc.Paragraphs.Last.Range.Font.Size = subject_font_size;
                            readDoc.Paragraphs.Last.Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceExactly;
                            readDoc.Paragraphs.Last.Range.ParagraphFormat.LineSpacing = 22F;
                            readDoc.Paragraphs.Last.Range.Text = this.textBox_LWname.Text + "\n\n";
                            pageNumber_1();
                        }
                        else if (numq == 3)
                        {
                            readDoc.Paragraphs.Last.Range.InsertBreak(MSWord.WdBreakType.wdPageBreak);
                            readDoc.Paragraphs.Add();
                            readDoc.Paragraphs.Last.Range.set_Style(MSWord.WdBuiltinStyle.wdStyleNormal);
                            readDoc.Paragraphs.Last.Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                            readDoc.Paragraphs.Last.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                            readDoc.Paragraphs.Last.Range.ParagraphFormat.FirstLineIndent = wordApp.CentimetersToPoints(0);
                            readDoc.Paragraphs.Last.Range.Font.Name = number_font_style;
                            readDoc.Paragraphs.Last.Range.Font.Size = subject_font_size;
                            readDoc.Paragraphs.Last.Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceExactly;
                            readDoc.Paragraphs.Last.Range.ParagraphFormat.LineSpacing = 22F;
                            readDoc.Paragraphs.Last.Range.Text = this.textBox_english.Text + "\n\n";

                        }
                        if (numq == 1 || numq == 2)
                        {
                            try
                            {
                                readDoc.Paragraphs.Last.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                                readDoc.Paragraphs.Last.Range.ParagraphFormat.FirstLineIndent = wordApp.CentimetersToPoints(0);
                                readDoc.Paragraphs.Last.set_Style(MSWord.WdBuiltinStyle.wdStyleHeading1);
                                readDoc.Paragraphs.Last.Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                                readDoc.Paragraphs.Last.Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceExactly;
                                readDoc.Paragraphs.Last.Range.ParagraphFormat.LineSpacing = 22F;
                                readDoc.Paragraphs.Last.Range.Text = k.Text + "\n";
                                readTemp = wordApp.Documents.Open(ref tempPath, ref QS, true, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS);
                                wordApp.Visible = false;
                                //设置正文格式
                                nullLine();
                                verticalSpacing();


                                readTemp.Content.Font.Size = 12;
                                readTemp.Content.Font.Name = "仿宋";
                                readTemp.Content.Font.Bold = 0;
                                thesisPicture();
                                thesisTable();
                                readTemp.Content.Copy();
                                
                                readDoc.Paragraphs.Last.Range.PasteAndFormat(MSWord.WdRecoveryType.wdFormatOriginalFormatting);
                                
                                
                                readTemp.Close(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);
                                readDoc.Paragraphs.Last.set_Style(MSWord.WdBuiltinStyle.wdStyleNormal);


                            }
                            catch
                            {
                                errorText = errorText + k.Text + "\n";
                                try
                                {
                                    readTemp.Close(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);
                                }
                                catch { }
                            }
                            progressBarForm.Current++;
                        }
                        else
                        {
                            try
                            {
                                readDoc.Paragraphs.Last.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                                readDoc.Paragraphs.Last.Range.ParagraphFormat.FirstLineIndent = wordApp.CentimetersToPoints(0);
                                readDoc.Paragraphs.Last.set_Style(MSWord.WdBuiltinStyle.wdStyleHeading1);
                                readDoc.Paragraphs.Last.Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                                readDoc.Paragraphs.Last.Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceExactly;
                                readDoc.Paragraphs.Last.Range.ParagraphFormat.LineSpacing = 22F;
                                readDoc.Paragraphs.Last.Range.Text = k.Text + "\n";
                                readTemp = wordApp.Documents.Open(ref tempPath, ref QS, true, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS);
                                wordApp.Visible = false;
                                //设置正文格式
                                nullLine();
                                verticalSpacing();


                                readTemp.Content.Font.Size = 12;
                                readTemp.Content.Font.Name = "Times New Roman";
                                readTemp.Content.Font.Bold = 0;
                                thesisPicture();
                                thesisTable();
                                readTemp.Content.Copy();
                                
                                readDoc.Paragraphs.Last.Range.PasteAndFormat(MSWord.WdRecoveryType.wdFormatOriginalFormatting);
                                
                                
                                readTemp.Close(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);
                                readDoc.Paragraphs.Last.set_Style(MSWord.WdBuiltinStyle.wdStyleNormal);
                                //readDoc.Paragraphs.Last.Range.Text = "\n";

                            }
                            catch
                            {
                                errorText = errorText + k.Text + "\n";
                                try
                                {
                                    readTemp.Close(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);
                                }
                                catch { }
                            }
                            if (numq == 4)
                            {
                                pageNumber_2();


                            }
                            progressBarForm.Current++;
                        }




                        numq++;
                    }
                    Application.DoEvents();

                }
                //readDoc.Paragraphs.Last.set_Style(MSWord.WdBuiltinStyle.wdStyleNormal);
            }
            else
            {
                int num = 1;
                foreach (DevComponents.AdvTree.Node k in node.Nodes)
                {
                    if (CreateThesisIsRuning == false)
                        return;
                    if (k.Level == 1)
                    {
                        readDoc.Paragraphs.Last.set_Style(MSWord.WdBuiltinStyle.wdStyleHeading2);
                    }
                    else if (k.Level == 2)
                    {
                        readDoc.Paragraphs.Last.set_Style(MSWord.WdBuiltinStyle.wdStyleHeading3);
                    }
                    else if (k.Level == 3)
                    {
                        readDoc.Paragraphs.Last.set_Style(MSWord.WdBuiltinStyle.wdStyleHeading4);
                    }
                    else if (k.Level == 4)
                    {
                        readDoc.Paragraphs.Last.set_Style(MSWord.WdBuiltinStyle.wdStyleHeading5);
                    }
                    tempPath = officePath + k.Text + ".doc";
                    //小标题,用判断来设置字体
                    if (k.Nodes.Count != 0)
                    {

                        readDoc.Paragraphs.Last.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                        readDoc.Paragraphs.Last.Range.ParagraphFormat.FirstLineIndent = wordApp.CentimetersToPoints(0);
                        readDoc.Paragraphs.Last.Range.Font.Name = small_font_style;
                        readDoc.Paragraphs.Last.Range.Font.Size = small_font_size;
                        readDoc.Paragraphs.Last.Range.Bold = 0;
                        readDoc.Paragraphs.Last.Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceExactly;
                        readDoc.Paragraphs.Last.Range.ParagraphFormat.LineSpacing = 22F;
                        readDoc.Paragraphs.Last.Range.Text = k.Text + "\n";
                        try
                        {

                            readTemp = wordApp.Documents.Open(ref tempPath, ref QS, true, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS);
                            wordApp.Visible = false;
                            //设置正文格式
                            nullLine();
                            verticalSpacing();

                            if (readTemp.Content.Text.Trim() != "")
                            {
                                //readTemp.Content.set_Style(MSWord.WdBuiltinStyle.wdStyleNormal);
                                readTemp.Content.Font.Size = text_font_size;
                                readTemp.Content.Font.Name = text_font_style;
                                readTemp.Content.Font.Bold = 0;
                                thesisPicture();
                                thesisTable();
                                readTemp.Content.Copy();
                                
                                readDoc.Paragraphs.Last.Range.PasteAndFormat(MSWord.WdRecoveryType.wdFormatOriginalFormatting);
                                
                                
                            }
                            readTemp.Close(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);
                        }
                        catch (Exception e)
                        {
                            try
                            {
                                readTemp.Close(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);
                            }
                            catch { }
                        }
                        thesisBody(k);
                    }
                    else
                    {
                        readDoc.Paragraphs.Last.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                        readDoc.Paragraphs.Last.Range.ParagraphFormat.FirstLineIndent = wordApp.CentimetersToPoints(0);
                        readDoc.Paragraphs.Last.Range.Font.Name = small_font_style;
                        readDoc.Paragraphs.Last.Range.Font.Size = small_font_size;
                        readDoc.Paragraphs.Last.Range.Bold = 0;
                        readDoc.Paragraphs.Last.Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceExactly;
                        readDoc.Paragraphs.Last.Range.ParagraphFormat.LineSpacing = 22F;
                        readDoc.Paragraphs.Last.Range.Text = k.Text + "\n";
                        try
                        {

                            readTemp = wordApp.Documents.Open(ref tempPath, ref QS, true, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS);
                            wordApp.Visible = false;
                            //设置正文格式
                            nullLine();
                            verticalSpacing();

                            //readTemp.Content.set_Style(MSWord.WdBuiltinStyle.wdStyleNormal);
                            readTemp.Content.Font.Size = text_font_size;
                            readTemp.Content.Font.Name = text_font_style;
                            readTemp.Content.Font.Bold = 0;
                            thesisPicture();
                            thesisTable();
                            readTemp.Content.Copy();
                            
                            readDoc.Paragraphs.Last.Range.PasteAndFormat(MSWord.WdRecoveryType.wdFormatOriginalFormatting);

                            
                            
                            readTemp.Close(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);

                        }
                        catch (Exception e)
                        {
                            errorText = errorText + k.Text + "\n";
                            try
                            {
                                readTemp.Close(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);
                            }
                            catch { }
                        }
                        progressBarForm.Current++;
                    }

                    num++;
                    Application.DoEvents();
                }
            }


        }
        int addNum = 0;
        object savepath, readPath;
        /*生成头信息函数,同时调用了正文函数*/
        private void thesisInfoAndList()
        {
            readPath = rootPath + "templet.doc";
            if (!File.Exists(readPath.ToString()))
            {
                MessageBox.Show(this, "模板文件不存在,请重新安装程序!!!", "失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            UpdateAdvNodesName(null, 1);
            /*输入赋值*/
            string[] blList = new string[9];
            blList[0] = this.textBox_xnumber.Text;
            blList[1] = this.textBox_LWname.Text;
            blList[2] = this.textBox_english.Text;
            blList[3] = this.textBox_name.Text;
            blList[4] = this.textBox_number.Text;
            blList[5] = this.textBox_xib.Text;
            blList[6] = this.textBox_zhuany.Text;
            blList[7] = this.textBox_teacher.Text + "  " + this.comboBox1.Text;
            string[] time = new string[2];
            string[] timenow = new string[3];
            time[0] = this.dateTimePicker1.Text.Split('/')[0] + "." + this.dateTimePicker1.Text.Split('/')[1];
            time[1] = this.dateTimePicker2.Text.Split('/')[0] + "." + this.dateTimePicker2.Text.Split('/')[1];
            blList[8] = time[0] + "-" + time[1];
            timenow = this.dateTimePicker3.Text.Split('/');



            wordApp = new MSWord.Application();
            readDoc = wordApp.Documents.Open(ref readPath, ref QS, true, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS);

            int j, len;
            string bookma, eng = "", engl = "";

            System.Collections.IEnumerator bookEmu = wordApp.ActiveDocument.Bookmarks.GetEnumerator();
            MSWord.Bookmark book;
            int i = 0;

            while (bookEmu.MoveNext())
            {



                blList[i] = blList[i].Trim();
                blList[i] = blList[i].Replace("\n", "");
                len = 0;
                book = (MSWord.Bookmark)bookEmu.Current;
                bookma = book.Name.ToString();
                for (j = 0; j < blList[i].Length; j++)
                {
                    if ((int)blList[i][j] >= 0 && (int)blList[i][j] <= 127)
                    {

                    }
                    else
                    {
                        len++;
                    }
                }

                if (i > 2 && i < 9)//其余的
                {
                    insertText(bookma, " " + (blList[i].PadRight(16 - len, ' ')), "0", "3.75", 0);
                }
                else if (i == 0)//科目号
                {
                    insertText(bookma, (blList[i].PadLeft((10 - blList[i].Length) / 2 + blList[i].Length, ' ')).PadRight(10, ' '), "0", "1.25", 0);
                }
                else if (i == 1)//论文中文名
                {
                    if (len == blList[i].Length)
                    {
                        if (blList[i].Length > 15)
                        {
                            int x, y, lenc;
                            string chs = "";

                            x = 0;
                            y = 15;
                            while (x + y - blList[i].Length < 0)
                            {
                                chs = chs + "    " + blList[i].Substring(x, y) + "\n";

                                x = x + y;

                            }
                            lenc = blList[i].Substring(x).Length;
                            chs = chs + ("    " + blList[i].Substring(x));
                            insertText(bookma, chs, "2.45", "1.5", 0);

                        }
                        else
                        {
                            insertText(bookma, "    " + blList[i], "2.45", "1.5", 0);
                        }
                    }
                    else
                    {
                        if (blList[i].Length > 15)
                        {
                            int k, lenc = 0, kk;
                            string chs = "    ";
                            for (k = 0; k < blList[i].Length; k++)
                            {
                                if (lenc == 15)
                                {
                                    lenc = 0;
                                    chs = chs + "\n    ";
                                }
                                if ((int)blList[i][k] >= 0 && (int)blList[i][k] <= 127)
                                {
                                    kk = k;
                                    while ((int)blList[i][kk] >= 0 && (int)blList[i][kk] <= 127 && (int)blList[i][kk] != 32)
                                    {
                                        kk++;
                                    }
                                    if (kk - k + lenc > 15)
                                    {
                                        lenc = 0;
                                        chs = chs + "\n    ";
                                    }
                                }
                                chs = chs + blList[i][k];
                                lenc++;
                            }
                            insertText(bookma, chs, "2.45", "1.5", 0);
                        }
                        else
                        {
                            insertText(bookma, "    " + blList[i], "2.45", "1.5", 0);
                        }
                    }
                }
                else if (i == 2)//论文英文名
                {

                    if (blList[i].Length + len > 38)
                    {
                        string[] strsplit = blList[i].Split(' ');
                        for (j = 0; j < strsplit.Length; j++)
                        {
                            engl = engl + strsplit[j] + " ";
                            if (j != strsplit.Length - 1)
                            {
                                if (engl.Length + strsplit[j + 1].Length >= 38)
                                {
                                    engl = engl.TrimEnd(' ');
                                    eng = eng + "    " + engl + "\n";
                                    engl = "";
                                    //aaaaa bbbbbb cccccc ddddd eeeee ffffff ggggg hhhhh iiiiii
                                }
                            }
                            else
                            {
                                engl = engl.TrimEnd(' ');
                                eng = eng + "    " + engl + "\n";
                                engl = "";
                            }
                        }
                        if (engl != "")
                        {
                            eng = eng + "    " + engl;
                        }
                        else
                        {
                            eng = eng.TrimEnd('\n');
                        }
                        insertText(bookma, eng, "5.4", "1.5", 0);

                    }
                    else
                    {
                        insertText(bookma, "    " + blList[i], "5.4", "1.5", 0);

                    }
                }

                i++;
            }

            replace("_year_", timenow[0]);
            replace("_month_", timenow[1]);
            replace("_day_", timenow[2]);
            thesisIndentation();
            addNum = readDoc.ComputeStatistics(MSWord.WdStatistic.wdStatisticPages, ref QS) - 5;
            readDoc.Paragraphs.Last.Range.InsertBreak(MSWord.WdBreakType.wdPageBreak);
            readDoc.Paragraphs.Add();



        }


        /*按钮*/

        private void button_Click_go(object sender, EventArgs e)//生成文档按钮
        {
            if (this.dateTimePicker1.Text == "" || this.dateTimePicker2.Text == "" || this.dateTimePicker3.Text == "")
            {
                MessageBox.Show(this, "日期不能为空!!!", "失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            this.saveFileDialog1.FileName = rootPath + "我的毕业论文.doc";
            this.saveFileDialog1.RestoreDirectory = true;
            if (this.saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                savepath = this.saveFileDialog1.FileName;
                start();
            }
        }
        private void lookThesis_Click(object sender, EventArgs e)//查看生成的论文
        {
            if ((string)savepath == null)
            {
                MessageBox.Show(this, "检测不到已生成的毕业论文，请手动选择！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (openFileDialog1.ShowDialog(this) == System.Windows.Forms.DialogResult.OK)
                {
                    savepath = openFileDialog1.FileName;
                }
                else
                {
                    return;
                }
            }
            bool file = openThesis();
            if (file == false)
            {
                MessageBox.Show(this, "文件" + savepath + "不存在！", "打开失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void Exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private DialogResult closeThesis(string cao)
        {

            return MessageBox.Show(this, "进行\"" + cao + "\"操作前需要先关闭编辑窗口,是否继续？", "关闭", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);

        }
        /*右键菜单*/
        private void clearDoc_Click(object sender, EventArgs e)//清空按钮
        {
            clearContent();
        }
        private void button3_Click(object sender, EventArgs e)//编辑按钮
        {

            if (this.advTree_1.SelectedNode != null)
            {
                thesisEdit();
            }
        }
        private void button4_Click(object sender, EventArgs e)//添加按钮
        {
            thesisListAdd();
        }

        private void button5_Click(object sender, EventArgs e)//插入按钮
        {
            thesisListInsert();
        }

        private void button6_Click(object sender, EventArgs e)//删除按钮
        {
            thesisListDelete();
        }

        private void button7_Click(object sender, EventArgs e)//重命名按钮
        {
            DialogResult Dr;
            if (haveDoc)
            {
                Dr = closeThesis("重命名");
                if (Dr == System.Windows.Forms.DialogResult.Cancel)
                {
                    return;
                }
                else
                {
                    foreach (DevComponents.DotNetBar.Office2007Form f in MdiChildren)
                    {
                        f.Close();
                    }
                }
            }
            if (this.advTree_1.SelectedNode != null)
            {
                this.advTree_1.SelectedNode.BeginEdit();
            }
        }

        private void 复制ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.textBox_english.SelectedText != "")
            {
                Clipboard.SetDataObject(this.textBox_english.SelectedText);
            }
        }

        private void 粘贴ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.textBox_english.SelectedText = Convert.ToString(Clipboard.GetDataObject().GetData(DataFormats.Text));
        }

        private void 全选ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.textBox_english.SelectAll();
        }

        private void 清空ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.textBox_english.Clear();
        }

        private void 剪切ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.textBox_english.SelectedText != "")
            {
                Clipboard.SetDataObject(this.textBox_english.SelectedText);
                this.textBox_english.SelectedText = "";
            }
        }

        /*事件*/

        private void advTree_1_NodeDoubleClick(object sender, DevComponents.AdvTree.TreeNodeMouseEventArgs e)
        {
            advTree_1.SelectedNode = e.Node;
            if (e.Node.Nodes.Count == 0)
            {
                thesisEdit();
            }
        }

        private void advTree_1_BeforeCellEdit(object sender, DevComponents.AdvTree.CellEditEventArgs e)//目录名编辑前事件,前后几项无法修改
        {
            string text;
            if (this.advTree_1.SelectedNode.Index > this.advTree_1.Nodes.Count - 3 || this.advTree_1.SelectedNode.Index < 4 && this.advTree_1.SelectedNode.Level == 0)
            {
                e.Cancel = true;

            }
            else
            {
                text = e.Cell.Text;
                if (text.Split(' ')[0] != text && (text[0] > '0' && text[0] <= '9'))
                {
                    e.Cell.Text = text.Remove(0, text.Split(' ')[0].Length + 1);
                }
            }
        }

        private void advTree_1_AfterCellEdit(object sender, DevComponents.AdvTree.CellEditEventArgs e)//编辑结束事件,修改文件名
        {
            UpdateAdvNodesName_RenameDelegate UdtD = new UpdateAdvNodesName_RenameDelegate(UpdateAdvNodesName_Rename);
            this.advTree_1.BeginInvoke(UdtD, e.Cell.Parent);
            listSaveDelegate ls = new listSaveDelegate(listSave);
            this.advTree_1.BeginInvoke(ls);


        }
        private void advTree_1_TextChanged(object sender, EventArgs e)//目录文字发生变化时更新目录并保存
        {
            UpdateAdvNodesName(null);
            listSave();
        }

        private void advTree_1_DragDrop(object sender, DragEventArgs e)//拖动结束更新目录
        {

            UpdateAdvNodesNameDelegate UdtD = new UpdateAdvNodesNameDelegate(UpdateAdvNodesName);
            this.advTree_1.BeginInvoke(UdtD, null, 0);

        }





        private void advTree_1_NodeDragFeedback(object sender, DevComponents.AdvTree.TreeDragFeedbackEventArgs e)//固定项无法拖动
        {
            if ((e.DragNode.Index < 4 || (e.DragNode.Index >= advTree_1.Nodes.Count - 2)) && e.DragNode.Level == 0)
            {
                e.Effect = DragDropEffects.None;

            }

            if (e.ParentNode == null)
            {

                if ((e.InsertPosition < 3 || e.InsertPosition >= advTree_1.Nodes.Count - 2))
                {
                    e.Effect = DragDropEffects.None;

                }
            }
            else
            {
                e.ParentNode.Expand();
                if ((e.ParentNode.Index < 4 || e.ParentNode.Index >= advTree_1.Nodes.Count - 2) && e.ParentNode.Level == 0)
                {
                    e.Effect = DragDropEffects.None;
                }
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            //this.editForm.Close();
            if (MessageBox.Show(this, "确定退出吗？\n退出后所有内容都会保存在程序所在文件夹内,请勿删除该文件夹!!!!", "退出", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Cancel)
            {
                e.Cancel = true;

            }
            listSave();

        }
        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                readApp.Quit(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);
            }
            catch (Exception e1)
            {


            }
            try
            {

                wordApp.Quit(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);
            }
            catch (Exception e1)
            {


            }

            Application.Exit();

        }
        /*输入框输入判断事件,检测*/
        private void textBox_LWname_TextChanged(object sender, EventArgs e)
        {
            int i, len = 0;
            for (i = 0; i < this.textBox_LWname.Text.Length; i++)
            {
                if ((int)this.textBox_LWname.Text[i] >= 0 && this.textBox_LWname.Text[i] < 127)
                {

                }
                else
                {
                    len++;
                }
            }
            if (len != this.textBox_LWname.Text.Length)
            {
                toolTip1.Show("检测到输入的中文名中含有英文字符\n论文排版会被打乱!", this.textBox_LWname, this.textBox_LWname.Width, 0);
            }
            else
            {
                toolTip1.Hide(this.textBox_LWname);
                if (len > 30)
                {
                    toolTip1.Show("检测到输入的中文名超过两行\n论文排版可能会被打乱!", this.textBox_LWname, this.textBox_LWname.Width, 0);
                }
            }
        }


        /*同上*/
        private void textBox_english_TextChanged(object sender, EventArgs e)
        {
            int i, len = 0;
            for (i = 0; i < this.textBox_english.Text.Length; i++)
            {
                if ((int)this.textBox_english.Text[i] >= 0 && this.textBox_english.Text[i] < 127)
                {

                }
                else
                {
                    len++;
                }
            }
            if (len != 0)
            {

                toolTip2.Show("检测到输入的英文名中含有中文字符\n论文排版会被打乱!", this.textBox_english, this.textBox_english.Width, this.textBox_english.Height / 2);
            }
            else
            {
                toolTip2.Hide(this.textBox_english);
                if (textBox_english.Text.Length > 140)
                {
                    toolTip2.Show("检测到输入的英文名字数超过四行\n论文排版可能会被打乱!", this.textBox_english, this.textBox_english.Width, this.textBox_english.Height / 2);
                }

            }
        }

        private void textBox_name_TextChanged(object sender, EventArgs e)
        {
            if (this.textBox_name.Text != "")
            {
                DocAuthor = this.textBox_name.Text;
                this.labelItem4.Text = DocAuthor;
            }
            else
            {
                DocAuthor = "无";
                this.labelItem4.Text = DocAuthor;
            }
        }
        private void advTree_1_NodeMouseDown(object sender, DevComponents.AdvTree.TreeNodeMouseEventArgs e)//右键菜单
        {
            if (e.Button != MouseButtons.Right)
                return;

            DevComponents.AdvTree.Node CurrentNode = e.Node;

            if (CurrentNode != null)
            {
                editThesis.Enabled = false;
                editAddproduct.Enabled = false;
                editAddprChildoduct.Enabled = false;
                editDeleteproduct.Enabled = false;
                renameproduct.Enabled = false;
                clearDoc.Enabled = true;
                if ((CurrentNode.Index < 3 || CurrentNode.Index > advTree_1.Nodes.Count - 3) && CurrentNode.Level == 0)
                {
                    editThesis.Enabled = true;
                }
                else if (CurrentNode.Index == 3 && CurrentNode.Level == 0)
                {
                    editThesis.Enabled = true;
                    editAddproduct.Enabled = true;
                }
                else
                {
                    editThesis.Enabled = true;
                    editAddproduct.Enabled = true;
                    editAddprChildoduct.Enabled = true;
                    editDeleteproduct.Enabled = true;
                    renameproduct.Enabled = true;

                }


                cms.Show(this.advTree_1, e.X, e.Y);

                this.advTree_1.SelectedNode = CurrentNode;

            }
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            this.advTree_1.Nodes.Clear();
            Control.CheckForIllegalCrossThreadCalls = false;
            listRead();
            UpdateAdvNodesName(null, 1);
            readSettings();
            this.tabStrip1.Hide();
            DocAuthor = this.textBox_name.Text;
            this.labelItem4.Text = DocAuthor;
            this.advTree_1.ExpandAll();



        }




        private void bar1_DockTabChange(object sender, DevComponents.DotNetBar.DockTabChangeEventArgs e)
        {
            if (e.NewTab.Text == "论文信息")
            {
                this.bar1.Text = "论文信息";
            }
            else
            {
                this.bar1.Text = "目录";

            }

        }



        private void openAbout()
        {
            ThesisEditerAboutForm aboutform = new ThesisEditerAboutForm();
            aboutform.ShowDialog();
        }
        private void selectNode(DevComponents.AdvTree.Node node)
        {
            if (this.MdiChildren.Length != 0)
            {
                this.advTree_1.SelectedNode = node;
                foreach (DevComponents.DotNetBar.Office2007Form f in MdiChildren)
                {
                    f.Close();
                }
                thesisEdit();
            }
        }
        //上一章节  （从最底层的节点开始打开）
        private void prevDocument(DevComponents.AdvTree.Node node = null)
        {
            if (node == null)//不带参数调用
            {
                if (nowSelectNode.PrevNode != null)//当前选择节点存在上一节点
                {
                    if (nowSelectNode.PrevNode.Nodes.Count == 0)//当前节点的上一节点不存在子节点
                    {
                        this.advTree_1.SelectedNode = nowSelectNode.PrevNode;

                        thesisEdit();
                    }
                    else//当前节点的上一节点存在子节点
                    {
                        prevDocument(nowSelectNode.PrevNode.LastNode);
                    }
                }
                else//当前选择节点不存在上一节点
                {
                    if (nowSelectNode.Parent != null)//当前节点的存在父节点
                    {
                        nowSelectNode = nowSelectNode.Parent;//标记父节点为当前节点
                        //判断对应文件是否存在
                        if (File.Exists(officePath + nowSelectNode.Text + ".doc"))
                        {
                            this.advTree_1.SelectedNode = nowSelectNode;
                            thesisEdit();
                        }
                        else
                        {
                            prevDocument();
                        }
                    }
                }
            }
            else//带参数调用
            {
                if (node.Nodes.Count == 0)//参数节点存在子节点
                {
                    this.advTree_1.SelectedNode = node;

                    thesisEdit();
                }
                else//参数节点不存在子节点
                {
                    prevDocument(node.LastNode);
                }
            }
        }
        //下一章节  （由于调用下一章节时需要从最顶层的节点开始打开，这和上一章节函数的原理不同，所以函数差异很大）
        private void nextDocument(DevComponents.AdvTree.Node node = null)
        {
            if (node == null)//不带参数调用
            {
                if (nowSelectNode != null)//当前选择了节点
                {
                    if (nowSelectNode.Nodes.Count == 0)//当前选择节点没有子节点
                    {
                        if (nowSelectNode.NextNode == null)//当前选择的节点不存在下一个节点
                        {
                            nextDocument(nowSelectNode.Parent);
                        }
                        else//当前选择的节点存在下一个节点
                        {
                            if (nowSelectNode.NextNode.Nodes.Count == 0)//当前选择的节点的下一节点没有子节点
                            {
                                this.advTree_1.SelectedNode = nowSelectNode.NextNode;
                                nowSelectNode = nowSelectNode.NextNode;
                                thesisEdit();
                            }
                            else//当前选择的节点的下一节点存在子节点
                            {
                                if (File.Exists(officePath + nowSelectNode.NextNode.Text + ".doc"))
                                {
                                    this.advTree_1.SelectedNode = nowSelectNode.NextNode;
                                    nowSelectNode = nowSelectNode.NextNode;
                                    thesisEdit();
                                }
                                else
                                {
                                    nowSelectNode = nowSelectNode.NextNode;
                                    nextDocument();
                                }
                            }

                        }
                    }
                    else//当前选择节点存在子节点
                    {

                        nowSelectNode = nowSelectNode.Nodes[0];//将当前节点的第一个子节点标记为当前节点
                        if (nowSelectNode.Nodes.Count == 0)//当前选择的节点没有子节点
                        {
                            this.advTree_1.SelectedNode = nowSelectNode;
                            thesisEdit();

                        }
                        else//当前选择的节点存在子节点
                        {
                            if (File.Exists(officePath + nowSelectNode.Text + ".doc"))
                            {
                                this.advTree_1.SelectedNode = nowSelectNode;
                                thesisEdit();
                            }
                            else
                            {
                                nextDocument();
                            }
                        }
                    }
                }
                else//当前没有选择节点
                {
                    nowSelectNode = advTree_1.Nodes[0];
                }
            }
            else//带参数调用
            {
                if (node.Nodes.Count == 0)//参数节点存在子节点
                {
                    this.advTree_1.SelectedNode = node;
                    nowSelectNode = node;
                    thesisEdit();
                }
                else//参数节点不存在子节点
                {
                    if (node.NextNode == null)//参数节点存在下一节点
                    {
                        nextDocument(node.Parent);
                    }
                    else//参数节点不存在下一节点
                    {
                        if (node.NextNode.Nodes.Count == 0)//参数节点的下一节点不存在子节点
                        {
                            this.advTree_1.SelectedNode = node.NextNode;
                            nowSelectNode = node.NextNode;
                            thesisEdit();
                        }
                        else//参数节点的下一节点存在子节点
                        {
                            if (File.Exists(officePath + node.NextNode.Text + ".doc"))
                            {
                                this.advTree_1.SelectedNode = node.NextNode;
                                nowSelectNode = node.NextNode;
                                thesisEdit();
                            }
                            else
                            {
                                nowSelectNode = node.NextNode;
                                nextDocument();
                            }
                        }
                    }
                }
            }
        }

        private void clearContent()
        {
            if (haveDoc)
            {
                DialogResult Dr;
                Dr = closeThesis("清空内容");
                if (Dr == System.Windows.Forms.DialogResult.Cancel)
                {
                    return;
                }
                else
                {
                    foreach (DevComponents.DotNetBar.Office2007Form f in MdiChildren)
                    {
                        f.Close();
                    }
                }
            }
            DevComponents.AdvTree.Node node = advTree_1.SelectedNode;
            if (File.Exists(officePath + node.Text + ".doc"))
            {
                if (MessageBox.Show(this, "执行【清空内容】操作后，该章节所保存的内容将会清空！是否执行？（子章节不受影响）", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.Yes)
                {
                    if (MessageBox.Show(this, "确认要执行吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.OK)
                    {
                        try
                        {
                            if (node.Level == 0 && (node.Index == 1 || node.Index == 3))
                            {
                                File.Delete(officePath + advTree_1.Nodes[1].Text + ".doc");
                                File.Delete(officePath + advTree_1.Nodes[3].Text + ".doc");
                            }
                            else
                            {
                                File.Delete(officePath + node.Text + ".doc");
                            }
                            MessageBox.Show(this, "清空成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch
                        {
                            MessageBox.Show(this, "清空失败", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
        }
        private void browseFirstPage()
        {
            foreach (DevComponents.DotNetBar.Office2007Form f in MdiChildren)
            {
                f.Close();
            }

            try
            {
                if (this.dateTimePicker1.Text == "" || this.dateTimePicker2.Text == "" || this.dateTimePicker3.Text == "")
                {
                    MessageBox.Show(this, "日期不能为空!!!", "失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                thesisInfoAndList();
            }
            catch (Exception e)
            {

            }
            if (File.Exists((string)officeTempName))
            {
                try
                {
                    File.Delete((string)officeTempName);
                }
                catch
                {
                    MessageBox.Show(this, "请关闭 " + (string)officeTempName + " \"效果浏览\"文件", "覆盖失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            try
            {
                readDoc.SaveAs(ref officeTempName, ref format, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS);
                readDoc.Close(true, ref QS, ref QS);
                wordApp.Quit(false, ref QS, ref QS);
            }
            catch
            {
                MessageBox.Show(this, "浏览失败！", "失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            FirstPageForm browseForm = new FirstPageForm(this);
            browseForm.Text = "效果浏览";
            browseForm.MdiParent = this;
            browseForm.Size = new System.Drawing.Size(0, 0);
            browseForm.WindowState = FormWindowState.Maximized;
            browseForm.Show();
            browseForm.FilePath = (string)officeTempName;

            if (browseForm.Retry == false)
            {
                browseForm.Close();
                return;
            }

            browseForm.Update();

            haveDoc = true;
            nowSelectNode = null;
            this.buttonItem5.Enabled = false;
            this.buttonItem7.Enabled = false;
            this.buttonItem12.Enabled = true;//测试
            this.buttonItem13.Enabled = true;

        }
        private DevComponents.AdvTree.Node rootNode(DevComponents.AdvTree.Node node)
        {
            string num;
            if (node.Level == 0)
                return node;
            num = node.Text.Split('.')[0];
            return advTree_1.Nodes[int.Parse(num) + 3];

        }
        private void indexPictureTitle()
        {
            /*检索图题号*/
            MSWord.InlineShapes shapes;
            shapes = readTemp.InlineShapes;
            foreach (MSWord.InlineShape sp in shapes)
            {

                try
                {
                    picIndex++;
                }
                catch (Exception e)
                {
                    new ErrorForm(e.ToString()).ShowDialog(this);
                }

            }
        }
        private void InsertPictureTitle()
        {
            /*图片题注处理*/
            MSWord.InlineShapes shapes;
            shapes = readTemp.InlineShapes;
            string title = "";
            foreach (MSWord.InlineShape sp in shapes)
            {

                if (sp.Range.Paragraphs.First.Next() == null)
                {
                    sp.Range.InsertAfter("\n");
                }
                picIndex++;
                title = sp.Range.Paragraphs.First.Next().Range.Text.Trim();//得到图片下一行文字
                if (title!="" && title[0] == '图' && title[1] > '0' && title[1] < '9')//如果下一行以‘图+数字’开头，就将其作为图片题注
                {
                    //为没有序号和标题之间不存在空格的题注增加空格
                    StringBuilder sb = new StringBuilder(sp.Range.Paragraphs.First.Next().Range.Text);

                    int knum = 0;
                    int len = sb.Length;
                    for (int i = 0; i < len; i++)
                    {
                        if (!((sb[i - knum] >= '0' && sb[i - knum] <= '9') || sb[i - knum] == '图' || sb[i - knum] == '.'))
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
                    title = sb.ToString();
                    //更新图片题注
                    sp.Range.Paragraphs.First.Next().Range.Text = title.Replace(title.Split(' ')[0], "图" + fiPicIndex + "." + picIndex);


                }
                else//如果下一行不是以‘图’开头
                {
                    //如果下一行是居中，不含嵌入图，不含表格，就当做没有序号的图题处理
                    if (sp.Range.Paragraphs.First.Next().Range.ParagraphFormat.Alignment == MSWord.WdParagraphAlignment.wdAlignParagraphCenter && sp.Range.Paragraphs.First.Next().Range.InlineShapes.Count == 0 && sp.Range.Paragraphs.First.Next().Range.Tables.Count == 0)
                    {
                        sp.Range.Paragraphs.First.Next().Range.Text = "图" + fiPicIndex + "." + picIndex + " " + sp.Range.Paragraphs.First.Next().Range.Text;
                    }
                    //否则当做没有图题
                    else
                    {
                        sp.Range.InsertAfter("\n\r");
                        sp.Range.Paragraphs.First.Next().Range.Text = "图" + fiPicIndex + "." + picIndex + " ";
                        sp.Range.Paragraphs.First.Next().Range.Select();
                    }

                }
                sp.Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                sp.Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceSingle;
                sp.Range.Paragraphs.First.Next().Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                sp.Range.Paragraphs.First.Next().Range.Font.Name = picture_font_style;
                sp.Range.Paragraphs.First.Next().Range.Font.Size = picture_font_size;


            }
        }
        private void InsertTableTitle()
        {
            string title = "";
            MSWord.Tables tabs = readTemp.Content.Tables;
            if (tabs.Count != 0)
            {
                foreach (MSWord.Table tab in tabs)
                {
                    if (tab.Rows.First.Range.Previous() == null)//表格在首行
                    {
                        //MessageBox.Show(this, "表格不能出现在文档首行!", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        tab.Range.Cut();
                        readTemp.Paragraphs.Add(readTemp.Range(0));
                        readTemp.Paragraphs.Add(readTemp.Range(1)).Range.Paste();

                    }
                    else
                    {
                        tabIndex++;
                        title = tab.Range.Paragraphs.First.Previous().Range.Text.Trim();
                        if (title != "" && title[0] == '表' && title[1] > '0' && title[1] < '9')//表上方文字为表+数字
                        {
                            //为没有序号和标题之间不存在空格的题注增加空格
                            StringBuilder sb = new StringBuilder(title);
                            int knum = 0;
                            int len = sb.Length;
                            for (int i = 0; i < len; i++)
                            {
                                if (!((sb[i - knum] >= '0' && sb[i - knum] <= '9') || sb[i - knum] == '表' || sb[i - knum] == '.'))
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

                            title = sb.ToString();
                            //找到表题
                            tab.Range.Paragraphs.First.Previous().Range.Delete();
                            tab.Rows.First.Range.Previous().InsertParagraphBefore();
                            tab.Rows.First.Range.Previous().Previous().Text = title.Replace(title.Split(' ')[0], "表" + fiTabIndex + "." + tabIndex);


                        }
                        else
                        {
                            if (tab.Range.Paragraphs.First.Previous().Range.ParagraphFormat.Alignment == MSWord.WdParagraphAlignment.wdAlignParagraphCenter && tab.Range.Paragraphs.First.Previous().Range.InlineShapes.Count == 0 && tab.Range.Paragraphs.First.Previous().Range.Tables.Count == 0)
                            {//表上方文字居中，并且不含图片
                                //tab.Range.Paragraphs.First.Previous().Range.Delete();
                                //tab.Rows.First.Range.Previous().InsertParagraphBefore();
                                //tab.Rows.First.Range.Previous().Previous().Text = "表" + fiTabIndex + "." + tabIndex + " " + title;
                                if (tab.Rows.First.Range.Previous().Previous() != null)
                                {
                                    tab.Rows.First.Range.Previous().InsertParagraphBefore();
                                    tab.Rows.First.Range.Previous().Previous().Text = "表" + fiTabIndex + "." + tabIndex + " " + title;
                                }
                                else
                                {
                                    tab.Rows.First.Range.Previous().InsertParagraphBefore();
                                    tab.Rows.First.Range.Previous().Previous().Text = "表" + fiTabIndex + "." + tabIndex + " " + title;
                                }
                            }
                            else//表上方文字不居中
                            {
                                tab.Rows.First.Range.Previous().InsertParagraphBefore();
                                tab.Rows.First.Range.Previous().InsertParagraphBefore();
                                tab.Rows.First.Range.Previous().Previous().Text = "表" + fiTabIndex + "." + tabIndex + " ";
                                tab.Rows.First.Range.Previous().Select();
                            }

                        }
                        tab.Range.Font.Size = intable_font_size;
                        tab.Range.Font.Name = intable_font_style;
                        tab.Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                        tab.Rows.Alignment = MSWord.WdRowAlignment.wdAlignRowCenter;
                        tab.Range.Paragraphs.First.Previous().Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                        tab.Range.Paragraphs.First.Previous().Range.Font.Name = table_font_style;
                        tab.Range.Paragraphs.First.Previous().Range.Font.Size = table_font_size;
                    }
                }
            }
        }
        private void updatePictureTitle(DevComponents.AdvTree.Node node)
        {
            tempPath = officePath + node.Text + ".doc";

            if (node.Text == nowDocName)
            {

                readTemp = (MSWord.Document)((EditForm)MdiChildren[0]).axFramerControl1.ActiveDocument;
                InsertPictureTitle();
            }
            else
            {
                if (File.Exists((string)tempPath))
                {
                    readTemp = wordApp.Documents.Open(ref tempPath, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, true, ref QS, ref QS, ref QS, ref QS);

                    InsertPictureTitle();
                    readTemp.Close(MSWord.WdSaveOptions.wdSaveChanges, ref QS, ref QS);
                }
            }
            progressBarItem1.Value++;
            if (node.Nodes.Count != 0)
            {
                foreach (DevComponents.AdvTree.Node k in node.Nodes)
                {
                    updatePictureTitle(k);
                }
            }
        }
        private void updateTableTitle(DevComponents.AdvTree.Node node)
        {
            tempPath = officePath + node.Text + ".doc";

            if (node.Text == nowDocName)
            {

                readTemp = (MSWord.Document)((EditForm)MdiChildren[0]).axFramerControl1.ActiveDocument;
                InsertTableTitle();
            }
            else
            {
                if (File.Exists((string)tempPath))
                {
                    readTemp = wordApp.Documents.Open(ref tempPath, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, ref QS, true, ref QS, ref QS, ref QS, ref QS);

                    InsertTableTitle();
                    readTemp.Close(MSWord.WdSaveOptions.wdSaveChanges, ref QS, ref QS);

                }
            }
            progressBarItem1.Value++;
            if (node.Nodes.Count != 0)
            {
                foreach (DevComponents.AdvTree.Node k in node.Nodes)
                {
                    updateTableTitle(k);
                }
            }
        }
        private void updatePic()
        {
            picUpdateIsRuning = true;
            buttonItem12.Enabled = false;
            tempNodeCount = 0;
            DevComponents.AdvTree.Node root = rootNode(nowSelectNode);
            nodeCounter(root);
            progressBarItem1.Maximum = tempNodeCount;
            progressBarItem1.Value = 0;
            try
            {

                if (root.Index <= advTree_1.Nodes.Count - 3 && root.Index >= 4 || root.Level != 0)
                {
                    wordApp = new MSWord.Application();
                    fiPicIndex = int.Parse(root.Text.Split(' ')[0]);
                    picIndex = 0;
                    updatePictureTitle(root);
                    wordApp.Quit(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);
                }
            }
            catch (Exception e)
            {
                new ErrorForm(e.ToString()).ShowDialog(this);
            }
            progressBarItem1.Value = 0;
            buttonItem12.Enabled = true;
            picUpdateIsRuning = false;
        }
        private void updateTab()
        {
            tabUpdateIsRuning = true;
            buttonItem13.Enabled = false;
            tempNodeCount = 0;
            DevComponents.AdvTree.Node root = rootNode(nowSelectNode);
            nodeCounter(root);
            progressBarItem1.Maximum = tempNodeCount;
            progressBarItem1.Value = 0;
            try
            {
                if (root.Index <= advTree_1.Nodes.Count - 3 && root.Index >= 4 || root.Level != 0)
                {
                    wordApp = new MSWord.Application();
                    fiTabIndex = int.Parse(root.Text.Split(' ')[0]);
                    tabIndex = 0;
                    updateTableTitle(root);
                    wordApp.Quit(MSWord.WdSaveOptions.wdDoNotSaveChanges, ref QS, ref QS);
                }
            }
            catch (Exception e)
            {
                new ErrorForm(e.ToString()).ShowDialog(this);
            }
            progressBarItem1.Value = 0;
            buttonItem13.Enabled = true;
            tabUpdateIsRuning = false;
        }
        private void listSearch()
        {
            if (searchNodes == null)
            {

                searchText = this.textBoxX1.Text;
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
                this.advTree_1.SelectedNode = searchNodes[searchIndex];

            }
        }
        private void listSerachRecurse(DevComponents.AdvTree.Node node = null)
        {
            if (node == null)
            {
                searchIndex = 0;
                foreach (DevComponents.AdvTree.Node k in advTree_1.Nodes)
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
        /// <summary>
        /// 子节点计数
        /// </summary>
        /// <param name="node">父节点</param>
        private void nodeCounter(DevComponents.AdvTree.Node node)
        {
            tempNodeCount++;
            if (node.Nodes.Count != 0)
            {
                foreach (DevComponents.AdvTree.Node k in node.Nodes)
                {
                    nodeCounter(k);
                }
            }
        }
        private void updatePicture()
        {
            if (MdiChildren.Length != 0 && MdiChildren[0] is EditForm)
            {
                if (((MSWord.Document)((EditForm)MdiChildren[0]).axFramerControl1.ActiveDocument).InlineShapes.Count == 0)
                {
                    if (MessageBox.Show(this, "检测到当前打开的章节不存在图片，是否执行“更新图序”操作？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.No)
                    {
                        return;
                    }
                }
                Thread newTh = new Thread(updatePic);
                newTh.Start();
            }
        }
        private void updateTable()
        {
            if (MdiChildren.Length != 0 && MdiChildren[0] is EditForm)
            {
                if (((MSWord.Document)((EditForm)MdiChildren[0]).axFramerControl1.ActiveDocument).Tables.Count == 0)
                {
                    if (MessageBox.Show(this, "检测到当前打开的章节不存在表格，是否执行“更新表序”操作？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.No)
                    {
                        return;
                    }
                }
                Thread newTh = new Thread(updateTab);
                newTh.Start();
            }
        }
        private void restartNum()
        {
            if (MdiChildren.Length != 0 && MdiChildren[0] is EditForm) {
                wordApp = ((MSWord.Document)((EditForm)MdiChildren[0]).axFramerControl1.ActiveDocument).Application;
                if (wordApp.Selection.Range.End - wordApp.Selection.Range.Start != 0)
                {
                    MessageBox.Show(this, "请不要使用光标标记内容，将光标移动到编号的位置！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (wordApp.Selection.Range.ListFormat.ListTemplate == null)
                {
                    MessageBox.Show(this, "当前光标位置没有项目符号或编号，请将光标移动到编号的位置！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                wordApp.Selection.Range.ListFormat.ApplyListTemplateWithLevel(wordApp.Selection.Range.ListFormat.ListTemplate, ContinuePreviousList: false);
            }

        }
        private void continueNum()
        {
            if (MdiChildren.Length != 0 && MdiChildren[0] is EditForm)
            {
                wordApp = ((MSWord.Document)((EditForm)MdiChildren[0]).axFramerControl1.ActiveDocument).Application;
                if (wordApp.Selection.Range.End - wordApp.Selection.Range.Start != 0)
                {
                    MessageBox.Show(this, "请不要使用光标标记内容，将光标移动到编号的位置！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (wordApp.Selection.Range.ListFormat.ListTemplate == null)
                {
                    MessageBox.Show(this, "当前光标位置没有项目符号或编号，请将光标移动到编号的位置！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                wordApp.Selection.Range.ListFormat.ApplyListTemplateWithLevel(wordApp.Selection.Range.ListFormat.ListTemplate, ContinuePreviousList: true);
            }

        }
        private void thesisSettings()
        {
            SettingsForm settingsForm = new SettingsForm(settingsStyleList, settingsSizeList, this);

            settingsForm.ShowDialog();
            if (settingsForm.DialogResult == System.Windows.Forms.DialogResult.OK)
            {
                readSettings();
            }
        }
        public void readSettings()
        {
            settingsStyleList = new int[11];
            settingsSizeList = new int[11];
            StringBuilder temp = new StringBuilder(500);
            string[] font_style_list = { "黑体", "仿宋", "宋体", "Times New Roman" };
            float[] font_size_list = { 22f, 18f, 16f, 15f, 14f, 12f, 10.5f, 9f, 7.5f, 6.5f };
            GetPrivateProfileString("Format", "subject_font_style", "0", temp, 500, settingsPath);
            settingsStyleList[0] = int.Parse(temp.ToString());
            subject_font_style = font_style_list[settingsStyleList[0]];
            GetPrivateProfileString("Format", "big_font_style", "0", temp, 500, settingsPath);
            settingsStyleList[1] = int.Parse(temp.ToString());
            big_font_style = font_style_list[settingsStyleList[1]];
            GetPrivateProfileString("Format", "small_font_style", "0", temp, 500, settingsPath);
            settingsStyleList[2] = int.Parse(temp.ToString());
            small_font_style = font_style_list[settingsStyleList[2]];
            GetPrivateProfileString("Format", "text_font_style", "0", temp, 500, settingsPath);
            settingsStyleList[3] = int.Parse(temp.ToString());
            text_font_style = font_style_list[settingsStyleList[3]];
            GetPrivateProfileString("Format", "page_font_style", "0", temp, 500, settingsPath);
            settingsStyleList[4] = int.Parse(temp.ToString());
            page_font_style = font_style_list[settingsStyleList[4]];
            GetPrivateProfileString("Format", "reference_font_style", "0", temp, 500, settingsPath);
            settingsStyleList[5] = int.Parse(temp.ToString());
            reference_font_style = font_style_list[settingsStyleList[5]];
            GetPrivateProfileString("Format", "mulu_font_style", "0", temp, 500, settingsPath);
            settingsStyleList[6] = int.Parse(temp.ToString());
            mulu_font_style = font_style_list[settingsStyleList[6]];
            GetPrivateProfileString("Format", "list_font_style", "0", temp, 500, settingsPath);
            settingsStyleList[7] = int.Parse(temp.ToString());
            list_font_style = font_style_list[settingsStyleList[7]];
            GetPrivateProfileString("Format", "picture_font_style", "0", temp, 500, settingsPath);
            settingsStyleList[8] = int.Parse(temp.ToString());
            picture_font_style = font_style_list[settingsStyleList[8]];
            GetPrivateProfileString("Format", "table_font_style", "0", temp, 500, settingsPath);
            settingsStyleList[9] = int.Parse(temp.ToString());
            table_font_style = font_style_list[settingsStyleList[9]];
            GetPrivateProfileString("Format", "intable_font_style", "0", temp, 500, settingsPath);
            settingsStyleList[10] = int.Parse(temp.ToString());
            intable_font_style = font_style_list[settingsStyleList[10]];

            GetPrivateProfileString("Format", "subject_font_size", "0", temp, 500, settingsPath);
            settingsSizeList[0] = int.Parse(temp.ToString());
            subject_font_size = font_size_list[settingsSizeList[0]];
            GetPrivateProfileString("Format", "big_font_size", "0", temp, 500, settingsPath);
            settingsSizeList[1] = int.Parse(temp.ToString());
            big_font_size = font_size_list[settingsSizeList[1]];
            GetPrivateProfileString("Format", "small_font_size", "0", temp, 500, settingsPath);
            settingsSizeList[2] = int.Parse(temp.ToString());
            small_font_size = font_size_list[settingsSizeList[2]];
            GetPrivateProfileString("Format", "text_font_size", "0", temp, 500, settingsPath);
            settingsSizeList[3] = int.Parse(temp.ToString());
            text_font_size = font_size_list[settingsSizeList[3]];
            GetPrivateProfileString("Format", "page_font_size", "0", temp, 500, settingsPath);
            settingsSizeList[4] = int.Parse(temp.ToString());
            page_font_size = font_size_list[settingsSizeList[4]];
            GetPrivateProfileString("Format", "reference_font_size", "0", temp, 500, settingsPath);
            settingsSizeList[5] = int.Parse(temp.ToString());
            reference_font_size = font_size_list[settingsSizeList[5]];
            GetPrivateProfileString("Format", "mulu_font_size", "0", temp, 500, settingsPath);
            settingsSizeList[6] = int.Parse(temp.ToString());
            mulu_font_size = font_size_list[settingsSizeList[6]];
            GetPrivateProfileString("Format", "list_font_size", "0", temp, 500, settingsPath);
            settingsSizeList[7] = int.Parse(temp.ToString());
            list_font_size = font_size_list[settingsSizeList[7]];
            GetPrivateProfileString("Format", "picture_font_size", "0", temp, 500, settingsPath);
            settingsSizeList[8] = int.Parse(temp.ToString());
            picture_font_size = font_size_list[settingsSizeList[8]];
            GetPrivateProfileString("Format", "table_font_size", "0", temp, 500, settingsPath);
            settingsSizeList[9] = int.Parse(temp.ToString());
            table_font_size = font_size_list[settingsSizeList[9]];
            GetPrivateProfileString("Format", "intable_font_size", "0", temp, 500, settingsPath);
            settingsSizeList[10] = int.Parse(temp.ToString());
            intable_font_size = font_size_list[settingsSizeList[10]];

        }
        private void updateApp()
        {
            Thread updateapp = new Thread(updateAppThread);
            updateapp.Start();

        }
        private void openUpdate()
        {
            System.Diagnostics.Process.Start(serverLink);
        }
        private void updateAppThread()
        {
            try
            {
                string s = DocumentSearch.HttpBrowser.GetHttpWebRequest("http://115.159.151.32:8080/thesisediter/version.jsp").Trim();
                serverVersion = s.Split(' ')[0];
                serverLink = s.Split(' ')[1];
                serverVersionNum = int.Parse(serverVersion.Replace(".", ""));
                if (serverVersionNum > thesisEditerVersionNum)
                {
                    labelItem7.Text = "毕业论文格式优化软件v" + serverVersion + "已发布,点击下载最新版";
                    hasUpdate = true;
                }
            }
            catch
            {
            }
        }

        private void buttonItem5_Click(object sender, EventArgs e)
        {
            prevDocument();

        }

        private void buttonItem7_Click(object sender, EventArgs e)
        {
            nextDocument();
        }



        private void buttonX1_Click_1(object sender, EventArgs e)
        {
            listSearch();
        }

        private void textBoxX1_TextChanged(object sender, EventArgs e)
        {
            searchNodes = null;
            if (textBoxX1.Text == "")
            {
                buttonX1.Enabled = false;
            }
            else
            {
                buttonX1.Enabled = true;
            }
        }

        private void help_Click(object sender, EventArgs e)
        {
            appHelp();
        }
        private void appHelp()
        {
            foreach (DevComponents.DotNetBar.Office2007Form f in MdiChildren)
            {
                f.Close();
            }
            FirstPageForm helpForm = new FirstPageForm(this);
            helpForm.MdiParent = this;
            helpForm.Size = new System.Drawing.Size(0, 0);
            helpForm.WindowState = FormWindowState.Maximized;
            helpForm.isPrint = false;
            helpForm.Text = "使用帮助";
            helpForm.Show();
            helpForm.FilePath = rootPath + @"\help.doc";

            if (helpForm.Retry == false)
            {
                editForm.Close();
                return;
            }

            helpForm.Update();

            haveDoc = true;
            nowSelectNode = null;
            this.buttonItem5.Enabled = false;
            this.buttonItem7.Enabled = false;
            this.buttonItem12.Enabled = false;
            this.buttonItem13.Enabled = false;
        }



        private void buttonItem10_Click(object sender, EventArgs e)
        {
            updatePicture();
        }

        private void buttonItem11_Click(object sender, EventArgs e)
        {
            updateTable();
        }

        private void buttonItem12_Click(object sender, EventArgs e)
        {

            updatePicture();
        }

        private void buttonItem13_Click(object sender, EventArgs e)
        {
            updateTable();
        }

        private void buttonItem9_Click(object sender, EventArgs e)
        {
            browseFirstPage();
        }

        private void buttonItem14_Click(object sender, EventArgs e)
        {
            

        }

        private void buttonItem15_Click(object sender, EventArgs e)
        {
            thesisSettings();
        }

        private void buttonItem16_Click(object sender, EventArgs e)
        {
            WordErrorForm wef = new WordErrorForm();
            wef.ShowDialog(this);
        }

        private void buttonItem17_Click(object sender, EventArgs e)
        {
            new BackForm().ShowDialog();
        }


        private void labelItem7_MouseLeave(object sender, EventArgs e)
        {
            if (hasUpdate)
            {
                this.labelItem7.ForeColor = System.Drawing.Color.Black;
            }
        }

        private void labelItem7_MouseMove(object sender, MouseEventArgs e)
        {
            if (hasUpdate)
            {
                this.labelItem7.ForeColor = System.Drawing.Color.Red;
            }
        }

        private void labelItem7_Click(object sender, EventArgs e)
        {
            if (hasUpdate)
            {
                openUpdate();
            }
            else
            {
                openAbout();
            }
        }

        private void advTree_1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                thesisListDelete();
            }
            else if (e.KeyCode == Keys.Insert)
            {
                thesisListInsert();
            }
        }



        private void buttonItem18_Click(object sender, EventArgs e)
        {
            if (hasUpdate)
            {
                openUpdate();
            }
            else
            {
                MessageBox.Show(this, "当前软件版本为最新版，无需更新！", "更新软件", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void buttonItem20_Click(object sender, EventArgs e)
        {
            openAbout();
        }

        private void buttonItem21_Click(object sender, EventArgs e)
        {
            new DocumentSearch.SearchAboutForm().ShowDialog();
        }

        private void buttonX3_Click(object sender, EventArgs e)
        {
            XuenumberSelectForm xueform = new XuenumberSelectForm();
            if (xueform.ShowDialog() == System.Windows.Forms.DialogResult.Yes)
            {
                textBox_xnumber.Text = xueform.xueNum;
                textBox_zhuany.Text = xueform.zhuanYe;
                textBox_xib.Text = xueform.xueYuan;
            }

        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            FanYiForm fyform = new FanYiForm(textBox_LWname.Text, 1);
            if (fyform.ShowDialog() == System.Windows.Forms.DialogResult.Yes)
            {
                this.textBox_english.Text = fyform.chinese;
            }

        }

        private void buttonItem22_Click(object sender, EventArgs e)
        {
            FanYiForm fyform = new FanYiForm();
            fyform.buttonX2.Visible = false;
            fyform.buttonX3.Visible = false;
            fyform.ShowDialog();
        }

        private void buttonItem23_Click(object sender, EventArgs e)
        {
            ImportDocumentForm idform = new ImportDocumentForm();
            if (idform.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                listRead();
            }
        }

        private void textBoxX1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                listSearch();
            }
        }

        private void buttonItem24_Click(object sender, EventArgs e)
        {
            restartNum();
        }

        private void buttonItem25_Click(object sender, EventArgs e)
        {
            continueNum();
        }













    }
}
///目前程序状态
///功能：
///1、生成毕业论文
///2、关键词翻译
///3、参考文献库
///4、封面浏览
///界面：
///1、采用第三方控件界面
///2、目录、论文信息使用可隐藏控件
///3、论文章节编辑使用mdi子窗口打开编辑，使用第三方office控件，关键词、参考文献使用独立窗口编辑
///4、生成过程使用进度条
///5、主界面信息条提供上下章节跳转按钮，以及一些论文信息
///未解决问题：
///1、表序图号识别问题,识别表序并处理修改为"表x.x ...",识别图号并处理修改为"图x.x ..."，已解决
///2、插图与图题,插图与图题为一整体，换页，为解决完美
///已解决问题：
///1、论文目录
///2、论文页码
///3、程序目录
///
/// 
/// 
/// 论文只改变字体大小和字号，不改变字体样式（颜色、下划线等）
/// 因为会影响项目符号
///
/// 修改进度条窗体，生成论文时可以关闭进度条  2016-11-20 19:44:09
/// 
/// 下一步  优化错误报告，生成论文时也可以正常报错。
/// 
/// 错误报告优化完成，可以实现报错并且和服务器通信，同时版本检测和意见反馈已正常。  2016-11-21 14:42:03
///
/// 项目全部重命名为  毕业论文格式优化软件  （命名空间没有修改）  2016-11-21 14:42:42
/// 
/// 下一步   优化参考文献查询系统
/// 
/// 优化参考文献查询系统 基本能够正常使用，未来可以直接修改服务器，无需修改客户端软件  2016-11-21 17:04:52
/// 
/// 版本检测优化  可以在服务器修改版本号和软件链接提供给用户下载   2016-11-21 17:07:26
/// 
/// 目前软件功能：
///     能够优化字体大小、字体类型和字体颜色（字体底色暂时不优化，因为会影响项目符号）
///     更新图序和表序(可以一键更新当前大章节内的所有图片或表格的序号，暂时未发现bug）
///     封面浏览（填写好用户信息后，可以进行封面浏览，查看封面效果）
///     翻译（使用百度翻译接口）
///     格式设置（用来应对学校修改格式需求）
///     导入论文（暂时只支持导入目录，前提是你的论文存在有目录，不存在怎么导入啊，这不是难为我吗？）
/// 
/// 下一步   优化 【导入论文】 功能
/// 
/// bug:点击【下一章节】，如果下一章节没有内容会自动创建。
///     拖动章节时，如果章节名一样会有bug，而且内容被清空
///     表格合并单元格时，更新表序异常
///     自动编号不会重置    
///     原因：前面章节有项目符号，当前章节也有项目符号，
///     项目符号就会被当做list，重新编号，真正的编号就没重新编号了。
///     同理，项目符号也会受编号影响
///     解决：
///         1、生成时处理，生成时检测编号处理,优点：用户简单，缺点：我麻烦，用户很难达到效果
///         2、生成后处理，生成后使用工具帮助用户处理该问题，即将编号重新编号。优点：我简单，缺点：用户麻烦
///