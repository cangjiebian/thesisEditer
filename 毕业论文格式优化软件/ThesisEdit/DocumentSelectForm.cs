using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace thesisEditer
{
    public partial class DocumentSelectForm : Form
    {
        public DocumentSelectForm()
        {
            InitializeComponent();
        }
        public string DocumentText = "";
        

        private void buttonItem1_Click(object sender, EventArgs e)
        {
            this.tabControl1.SelectedTabIndex = 0;
        }

        private void buttonItem2_Click(object sender, EventArgs e)
        {
            this.tabControl1.SelectedTabIndex = 1;
        }

        private void buttonItem3_Click(object sender, EventArgs e)
        {
            this.tabControl1.SelectedTabIndex = 2;
        }

        private void buttonItem4_Click(object sender, EventArgs e)
        {
            this.tabControl1.SelectedTabIndex = 3;
        }

        private void buttonItem5_Click(object sender, EventArgs e)
        {
            this.tabControl1.SelectedTabIndex = 4;
        }

        private void buttonItem6_Click(object sender, EventArgs e)
        {
            this.tabControl1.SelectedTabIndex = 5;
        }

        private void buttonItem7_Click(object sender, EventArgs e)
        {
            this.tabControl1.SelectedTabIndex = 6;
        }

        private void buttonItem8_Click(object sender, EventArgs e)
        {
            this.tabControl1.SelectedTabIndex = 7;
        }

        private void buttonItem9_Click(object sender, EventArgs e)
        {
            this.tabControl1.SelectedTabIndex = 8;
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void DocumentSelectForm_Load(object sender, EventArgs e)
        {
            this.tabControl1.SelectedTabIndex = 1;
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            switch (this.tabControl1.SelectedTabIndex)
            {
                case 0:
                    {
                        if (isOver(this.tabControlPanel1) == false)
                        {
                            return;
                        }
                        DocumentText = textBoxX11.Text + "." + textBoxX12.Text + "[J]." + textBoxX13.Text + "." + textBoxX14.Text + "," + textBoxX15.Text + ":" + textBoxX16.Text + ".";
                        break;
                    }
                case 1:
                    {
                        if (isOver(this.tabControlPanel2) == false)
                        {
                            return;
                        }
                        DocumentText = textBoxX21.Text + "." + textBoxX22.Text + "[M]." + textBoxX23.Text + ":" + textBoxX24.Text + "," + textBoxX25.Text + ":" + textBoxX26.Text + ".";
                        break;
                    }
                case 2:
                    {
                        if (isOver(this.tabControlPanel3) == false)
                        {
                            return;
                        }
                        DocumentText = textBoxX31.Text + "." + textBoxX32.Text + "[A]." + textBoxX33.Text + "[C]." + textBoxX34.Text + ":" + textBoxX35.Text + "," + textBoxX36.Text + ":" + textBoxX37.Text + ".";
                        break;
                    }
                case 3:
                    {
                        if (isOver(this.tabControlPanel4) == false)
                        {
                            return;
                        }
                        DocumentText = textBoxX41.Text + "." + textBoxX42.Text + "[D]." + textBoxX43.Text + ":" + textBoxX44.Text + "," + textBoxX45.Text + ".";
                        break;
                    }
                case 4:
                    {
                        if (isOver(this.tabControlPanel5) == false)
                        {
                            return;
                        }
                        DocumentText = textBoxX51.Text + "." + textBoxX52.Text + "[R]." + textBoxX53.Text + ":" + textBoxX54.Text + "," + textBoxX55.Text + ".";
                        break;
                    }
                case 5:
                    {
                        if (isOver(this.tabControlPanel6) == false)
                        {
                            return;
                        }
                        DocumentText = textBoxX61.Text + "." + textBoxX62.Text + "[P]." + textBoxX63.Text + ":" + textBoxX64.Text + "," + textBoxX65.Text + ".";
                        break;
                    }
                case 6:
                    {
                        if (isOver(this.tabControlPanel7) == false)
                        {
                            return;
                        }
                        DocumentText = textBoxX71.Text + "," + textBoxX72.Text + "[S]." + textBoxX73.Text + ":" + textBoxX74.Text + "," + textBoxX75.Text + ".";
                        break;
                    }
                case 7:
                    {
                        if (isOver(this.tabControlPanel8) == false)
                        {
                            return;
                        }
                        DocumentText = textBoxX81.Text + "." + textBoxX82.Text + "[N]." + textBoxX83.Text + "," + textBoxX84.Text + ".";
                        break;
                    }
                case 8:
                    {
                        if (this.richTextBox1.Text == "")
                        {
                            MessageBox.Show("请输入参考信息内容", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        DocumentText = this.richTextBox1.Text;
                        break;
                    }
            }
            
            this.Close();
        }

        private bool isOver(DevComponents.DotNetBar.TabControlPanel tabControl)
        {
            foreach (Control nowCC in tabControl.Controls)
            {
                if (nowCC.GetType() == typeof(DevComponents.DotNetBar.Controls.TextBoxX))
                {
                    if (nowCC.Text == "")
                    {
                        MessageBox.Show("请将文献信息填写完整!\n如果不清楚文献信息可以选择\"其他文献\"自定义输入", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                }

            }
            return true;
        }

        private void richTextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                this.buttonX1.Focus();
            }
        }

        private void 复制ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.richTextBox1.SelectedText != "")
            {
                Clipboard.SetDataObject(this.richTextBox1.SelectedText);
            }
        }

        private void 粘贴ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.richTextBox1.SelectedText = Convert.ToString(Clipboard.GetDataObject().GetData(DataFormats.Text));
        }

        private void 全选ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.richTextBox1.SelectAll();
        }

        private void 清空ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.richTextBox1.Clear();
        }

        private void 剪切ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.richTextBox1.SelectedText != "")
            {
                Clipboard.SetDataObject(this.richTextBox1.SelectedText);
                this.richTextBox1.SelectedText = "";
            }
        }

        private void tabItem1_Click(object sender, EventArgs e)
        {
            this.itemPanelSelect.SelectedItem = buttonItem1;
        }

        private void tabItem2_Click(object sender, EventArgs e)
        {
            this.itemPanelSelect.SelectedItem = buttonItem2;

        }

        private void tabItem3_Click(object sender, EventArgs e)
        {
            this.itemPanelSelect.SelectedItem = buttonItem3;

        }

        private void tabItem4_Click(object sender, EventArgs e)
        {
            this.itemPanelSelect.SelectedItem = buttonItem4;

        }

        private void tabItem5_Click(object sender, EventArgs e)
        {
            this.itemPanelSelect.SelectedItem = buttonItem5;

        }

        private void tabItem6_Click(object sender, EventArgs e)
        {
            this.itemPanelSelect.SelectedItem = buttonItem6;

        }

        private void tabItem7_Click(object sender, EventArgs e)
        {
            this.itemPanelSelect.SelectedItem = buttonItem7;

        }

        private void tabItem8_Click(object sender, EventArgs e)
        {
            this.itemPanelSelect.SelectedItem = buttonItem8;

        }

        private void tabItem9_Click(object sender, EventArgs e)
        {
            this.itemPanelSelect.SelectedItem = buttonItem9;

        }

       
        

        

        

        
    }
}
