using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Web;
using System.Threading;
using System.IO;

namespace DocumentSearch
{
    public partial class SearchMainForm : Form
    {
        public SearchMainForm()
        {
            InitializeComponent();
            toolStripProgressBar1.ProgressBar.Hide();
            myBooks = new List<Book>();
        }
        List<Book> books = null;//搜索结果列表
        List<Book> myBooks = null;//我的书籍列表
        string searchWord;//搜索关键字
        bool isWait = false;//是否在等待操作
        int searchType;//搜索类型
        int searchPage = 1;//请求的页码
        int searchPageSize = 50;//每页显示信息条数
        int maxPage;//最大页码
        int nowPage = 1;//当前页码
        int bookNum;//搜索结果数


        public string returnBookStr = "";
        private delegate void MyDelegate();
        /*功能函数*/

        /// <summary>
        /// 弹出搜索框
        /// </summary>
        private void searchBook()
        {
            if (isWait)
            {
                MessageBox.Show("系统正在执行操作，请稍等片刻再搜索", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            InputSearchForm inputS = new InputSearchForm();
            DialogResult result = inputS.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                if (inputS.word == "")
                {
                    return;
                }
                searchWord = inputS.word;
                searchType = inputS.type;
                searchPage = 1;
                search();
            }
        }
        /// <summary>
        /// 上一页下一页按钮处理
        /// </summary>
        private void buttonEn()
        {
            if (nowPage >= maxPage)
                button2.Enabled = false;
            else
                button2.Enabled = true;
            if (nowPage <= 1)
                button3.Enabled = false;
            else
                button3.Enabled = true;
        }
        /// <summary>
        /// 搜索函数
        /// </summary>
        private void search()
        {
            
            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            contextMenuStrip1.Enabled = false;
            toolStripProgressBar1.ProgressBar.Show();
            isWait = true;
            toolStripStatusNoteLabel1.Text = "请稍等";            
            new Thread(searchThread).Start();
        }
        
        /// <summary>
        /// 添加到【我的书籍】
        /// </summary>
        private void addBookToList()
        {
            if (dataGridView1.SelectedRows.Count == 0)
            {
                return;
            }
            Book addbook = books[dataGridView1.SelectedRows[0].Index];
            if (myBooks.Exists(delegate(Book book) { if (book.CallNumber == addbook.CallNumber && book.Title == addbook.Title)return true; else return false; }))
            {
                if (MessageBox.Show("检测到【我的书籍】里已经存在该书籍了，确定继续添加吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == System.Windows.Forms.DialogResult.Cancel)
                {
                    listView1.Select();
                    listView1.Items[ myBooks.FindIndex(delegate(Book book) { if (book.CallNumber == addbook.CallNumber && book.Title == addbook.Title)return true; else return false; })].Selected = true;
                    return;
                }
            }
            myBooks.Add(addbook);
            listView1.Items.Add(new ListViewItem(new string[] {addbook.Title, addbook.Author, addbook.CallNumber }));
            
            saveMyBooks();
            toolStripStatusNoteLabel1.Text = "添加成功";
        }
        /// <summary>
        /// 搜索线程
        /// </summary>
        private void searchThread()
        {
            
            string returnJson = "";
            try
            {
                dataGridView1.Rows.Clear();
                returnJson = HttpBrowser.GetHttpWebRequest("http://115.159.151.32:8080/BooksSearch/hhtcbooks/search.jsp?" + "name=" + HttpUtility.UrlEncode(searchWord, System.Text.Encoding.GetEncoding("utf-8")) + "&type=" + searchType + "&page=" + searchPage + "&pagesize=" + searchPageSize).Trim();
                Root root = Newtonsoft.Json.JsonConvert.DeserializeObject<Root>(returnJson);
                books = root.Books;
                maxPage = root.MaxPage;
                nowPage = root.NowPage;
                bookNum = root.BookNum;
                //this.listView2.BeginUpdate();
                dataGridView1.Invoke(new MyDelegate(showThread));
                //this.listView2.EndUpdate();
            }
            catch (Exception e)
            {
                MessageBox.Show("无法连接网络！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            finally
            {
                toolStripProgressBar1.ProgressBar.Hide();
                isWait = false;
                toolStripStatusNoteLabel1.Text = "搜索完毕";
                textBox1.Text = nowPage.ToString();
                button1.Enabled = true;
                button2.Enabled = true;
                button3.Enabled = true;
                button4.Enabled = true;
                contextMenuStrip1.Enabled = true;
            }
            
            
            buttonEn();
        }
        private void showThread()
        {
            foreach (Book book in books)
            {
                int index = dataGridView1.Rows.Add();
                dataGridView1.Rows[index].Cells[0].Value = book.Title;
                dataGridView1.Rows[index].Cells[1].Value = book.Author;
                dataGridView1.Rows[index].Cells[2].Value = book.ISBN;
                dataGridView1.Rows[index].Cells[3].Value = book.CallNumber;
                dataGridView1.Rows[index].Cells[4].Value = book.Press;
                dataGridView1.Rows[index].Cells[5].Value = book.Place;
                dataGridView1.Rows[index].Cells[6].Value = book.Date;
                dataGridView1.Rows[index].Cells[7].Value = book.Page;

            }

            toolStripStatusLabel2.Text = "关键字：" + searchWord + "  搜索结果：" + bookNum + "条书籍信息  每页显示：" + searchPageSize + "条  总页数：" + maxPage + "页  当前页数：" + nowPage + "页";
        }
        /// <summary>
        /// 显示详细信息函数（搜索结果）
        /// </summary>
        private void showBookMessage_2()
        {
            
            if (dataGridView1.SelectedRows.Count == 0)
            {
                return;
            }
            if (books[dataGridView1.SelectedRows[0].Index].Link != "")
            {
                DialogResult dr = new BookMessageForm(books[dataGridView1.SelectedRows[0].Index].Link).ShowDialog();
                if (dr == System.Windows.Forms.DialogResult.Yes)
                {
                    addBookToList();
                }
            }
            else
            {
                MessageBox.Show("该书籍暂时没有详细信息！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        /// <summary>
        /// 显示详细信息函数（我的书籍）
        /// </summary>
        private void showBookMessage_1()
        {
            
            if (listView1.SelectedItems.Count == 0)
            {
                return;
            }
            if (myBooks[listView1.SelectedItems[0].Index].Link != "")
            {
                new BookMessageForm(myBooks[listView1.SelectedItems[0].Index].Link,false).ShowDialog();
            }
            else
            {
                MessageBox.Show("该书籍暂时没有详细信息！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        
        /// <summary>
        /// 复制信息函数（我的书籍）
        /// </summary>
        private void copyBook_1()
        {
            if (listView1.SelectedItems.Count == 0)
            {
                return;
            }
            Book copybook = myBooks[listView1.SelectedItems[0].Index];
            Clipboard.SetDataObject(copybook.Author + "." + copybook.Title + "[M]." + copybook.Place + ":" + copybook.Press + "," + copybook.Date + ":1-100.");
            toolStripStatusNoteLabel1.Text = "复制成功";
        }
        /// <summary>
        /// 复制信息函数（搜索结果）
        /// </summary>
        private void copyBook_2()
        {
            if (dataGridView1.SelectedRows.Count == 0)
            {
                return;
            }
            Book copybook = books[dataGridView1.SelectedRows[0].Index];
            Clipboard.SetDataObject(copybook.Author + "." + copybook.Title + "[M]." + copybook.Place + ":" + copybook.Press + "," + copybook.Date + ":1-100.");
            toolStripStatusNoteLabel1.Text = "复制成功";
        }
        /// <summary>
        /// 下一页函数
        /// </summary>
        private void nextPage()
        {
            if (dataGridView1.Rows.Count != 0)
            {
                if (searchPage < maxPage)
                {
                    searchPage = nowPage + 1;
                    search();
                }
            
            }
            else
            {
                MessageBox.Show("请先【搜索】关键字", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            
        }
        /// <summary>
        /// 上一页函数
        /// </summary>
        private void previousPage()
        {
            if (dataGridView1.Rows.Count != 0)
            {
                if (searchPage > 1)
                {
                    searchPage = nowPage - 1;
                    search();
                }
            }
            else
            {
                MessageBox.Show("请先【搜索】关键字", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            
        }
        /// <summary>
        /// 首页函数
        /// </summary>
        private void firstPage()
        {
            if (dataGridView1.Rows.Count != 0)
            {
                searchPage = 1;
                search();
                
            }
            else
            {
                MessageBox.Show("请先【搜索】关键字", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        /// <summary>
        /// 尾页函数
        /// </summary>
        private void lastPage()
        {
            if (dataGridView1.Rows.Count != 0)
            {
                searchPage = maxPage;
                search();

            }
            else
            {
                MessageBox.Show("请先【搜索】关键字", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        /// <summary>
        /// 导出函数
        /// </summary>
        private void exportMessage()
        {
            if (myBooks.Count == 0)
            {
                if (MessageBox.Show("【我的书籍】为空，确定要导出？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.Cancel)
                    return;
            }
            this.saveFileDialog1.RestoreDirectory = true;
            if (this.saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                StreamWriter sw = new StreamWriter(this.saveFileDialog1.FileName,false);
                int i = 1;
                foreach (Book book in myBooks)
                {
                    sw.WriteLine("[" + i + "] " + book.Author + "." + book.Title + "[M]." + book.Place + ":" + book.Press + "," + book.Date + ":1-100.");
                    i++;
                }
                sw.Close();
                toolStripStatusNoteLabel1.Text = "导出成功";
            }
        }
        /// <summary>
        /// 清空我的书籍函数
        /// </summary>
        private void clearBook()
        {
            if (MessageBox.Show("确定要清空列表吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.OK)
            {
                myBooks.Clear();
                listView1.Items.Clear();
                saveMyBooks();
                toolStripStatusNoteLabel1.Text = "清空成功";
            }
        }
        /// <summary>
        /// 删除某条书籍信息
        /// </summary>
        private void deleteBook()
        {
            if (listView1.SelectedItems.Count != 0)
            {
                if (MessageBox.Show("确定要删除该条记录？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.OK)
                {
                    myBooks.RemoveAt(listView1.SelectedItems[0].Index);
                    listView1.Items.RemoveAt(listView1.SelectedItems[0].Index);
                    saveMyBooks();
                    toolStripStatusNoteLabel1.Text = "删除成功";
                }
            }
        }
        /// <summary>
        /// 关于窗口函数
        /// </summary>
        private void showAbout()
        {
            new SearchAboutForm().ShowDialog();
        }
        /// <summary>
        /// 保存我的书籍
        /// </summary>
        private void saveMyBooks()
        {
            StreamWriter sw = new StreamWriter(@"MyBooks.json",false);
            sw.Write(Newtonsoft.Json.JsonConvert.SerializeObject(myBooks));
            sw.Close();
        }
        /// <summary>
        /// 读取我的书籍
        /// </summary>
        private void readMyBooks()
        {
            if (!File.Exists(@"MyBooks.json"))
            {

                return;
            }
            StreamReader sr = new StreamReader(@"MyBooks.json");
            string jsonText = sr.ReadToEnd();
            sr.Close();
            List<Book> readBooks = Newtonsoft.Json.JsonConvert.DeserializeObject<List<Book>>(jsonText);
            myBooks = new List<Book>();
            foreach (Book book in readBooks)
            {
                listView1.Items.Add(new ListViewItem(new string[] {book.Title, book.Author, book.CallNumber }));
                myBooks.Add(book);
            }
            
        }
        /// <summary>
        /// 返回一个随机数
        /// </summary>
        /// <param name="len">随机数范围1-len</param>
        /// <param name="seed">种子</param>
        /// <returns></returns>
        private int getRandomNum(int len, int seed)
        {
            Random rd = new Random(unchecked(seed));
            return rd.Next(1, len);
        }
        /// <summary>
        /// 修改窗口返回内容
        /// </summary>
        private void returnBook()
        {
            if (myBooks.Count == 0)
            {
                if (MessageBox.Show("【我的书籍】为空，确定要保存？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.Cancel)
                    return;
            }
            int i = 1;
            string page = "1-10";
            int st = (int)DateTime.Now.Ticks, end;
            foreach (Book book in myBooks)
            {
                if (this.checkBox1.CheckState == CheckState.Checked)
                {
                    st = getRandomNum(200, st);
                    end = st + getRandomNum(50, st);
                    page = st.ToString() + "-" + end.ToString();
                }
                returnBookStr += "[" + i + "] " + book.Author + "." + book.Title + "[M]." + book.Place + ":" + book.Press + "," + book.Date +":"+ page + "." + "\n";
                i++;
            }
            returnBookStr.TrimEnd('\n');
            this.Close();
        }



        /*事件函数*/
        private void button2_Click(object sender, EventArgs e)
        {
            nextPage();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            previousPage();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            lastPage();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            firstPage();
        }

        

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F && e.Control == true)
            {
                searchBook();
            }
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {
            if (dataGridView1.SelectedRows.Count == 0)
            {
                
                e.Cancel = true;
            }
        }

        private void 查看详情ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            showBookMessage_2();
        }

        private void 添加到我的书籍ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            addBookToList();
        }

        private void 复制信息ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            copyBook_2();
        }

        private void 上一页ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            previousPage();
        }

        private void 下一页ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            nextPage();
        }

        private void 查看详情ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            showBookMessage_1();
        }

        private void 复制信息ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            copyBook_1();
        }


        private void contextMenuStrip2_Opening(object sender, CancelEventArgs e)
        {
            if (listView1.SelectedItems.Count == 0)
            {
                e.Cancel = true;
            }
        }

        private void listView1_DoubleClick(object sender, EventArgs e)
        {
            showBookMessage_1();
        }

        private void 导出全部ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            exportMessage();
        }

        private void 删除ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            deleteBook();
        }

        private void listView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                deleteBook();
            }
        }

        private void 清空列表ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            clearBook();
        }

        private void 关于AToolStripMenuItem_Click(object sender, EventArgs e)
        {
            showAbout();
        }

        private void 搜索ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            searchBook();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            Column1.Width = dataGridView1.Width / 8;
            Column2.Width = dataGridView1.Width / 8;
            Column3.Width = dataGridView1.Width / 8;
            Column4.Width = dataGridView1.Width / 8;
            Column5.Width = dataGridView1.Width / 8;
            Column6.Width = dataGridView1.Width / 8;
            Column7.Width = dataGridView1.Width / 8;
            Column8.Width = dataGridView1.Width / 8;
            columnHeader1.Width = listView1.Width / 3;
            columnHeader2.Width = listView1.Width / 3;
            columnHeader3.Width = listView1.Width / 3;
            Control.CheckForIllegalCrossThreadCalls = false;
        }

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();

        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (Char)8)
            {
                e.Handled = true;
            }
            else
            {
                toolStripStatusNoteLabel1.Text = "请输入数字";
            }
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (dataGridView1.Rows.Count != 0)
                {
                    try
                    {
                        searchPage = int.Parse(textBox1.Text);
                    }
                    catch
                    {
                        searchPage = 1;
                    }
                    search();
                }
                else
                {
                    MessageBox.Show("请先【搜索】关键字", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count == 0)
                return;
            Book selectBook = myBooks[listView1.SelectedItems[0].Index];
            toolStripStatusLabel2.Text = "序号：【"+(listView1.SelectedItems[0].Index+1)+"】"+"文献信息："+ selectBook.Author + "." + selectBook.Title + "[M]." + selectBook.Place + ":" + selectBook.Press + "," + selectBook.Date;
        }

        

        private void dataGridView1_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (e.RowIndex == -1) return;
                dataGridView1.Rows[e.RowIndex].Selected = true;
            }
        }

        
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.SelectedRows.Count == 0)
                return;
            Book selectBook = books[dataGridView1.SelectedRows[0].Index];
            toolStripStatusLabel2.Text = "文献信息：" + selectBook.Author + "." + selectBook.Title + "[M]." + selectBook.Place + ":" + selectBook.Press + "," + selectBook.Date;
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1)
                return;
            dataGridView1.Rows[e.RowIndex].Selected = true;
            showBookMessage_2();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            returnBook();
            
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (this.checkBox1.CheckState == CheckState.Unchecked)
            {
                MessageBox.Show("建议使用程序的随机页码,效果非常好!\n(不使用随机页码后，添加上的参考文献会统一使用默认页码\"1-10\"，请记得去参考文献界面修改)", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

    }
}
