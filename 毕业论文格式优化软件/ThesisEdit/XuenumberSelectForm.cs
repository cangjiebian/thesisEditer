using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Globalization;

namespace thesisEditer
{
    public partial class XuenumberSelectForm : Form
    {
        public string xueNum = null;
        public string zhuanYe = null;
        public string xueYuan = null;
        string[,] xueList = { 
                            { "数学与计算科学学院", "数学与应用数学", "0701" },
                            { "数学与计算科学学院", "信息与计算科学", "0701" },
                            { "体育学院","体育教育", "0402" },
                            { "体育学院","社会体育指导与管理", "0402" }, 
                            { "生物与食品工程学院","生物工程", "0830" },
                            { "生物与食品工程学院","生物科学", "0710" }, 
                            { "生物与食品工程学院","食品科学与工程", "0827" }, 
                            { "生物与食品工程学院","食品质量与安全", "0827" }, 
                            { "生物与食品工程学院", "生物制药", "0830" },
                            { "美术学院","美术学", "1304" },
                            { "音乐舞蹈学院", "音乐学", "1302" }, 
                            { "音乐舞蹈学院","舞蹈学", "1302" },
                            { "音乐舞蹈学院","音乐表演", "1302" },
                            { "外国语学院","英语", "0502" },
                            { "外国语学院", "商务英语", "0502" },
                            { "电气与信息工程学院","电子信息科学与技术", "0807" },
                            { "电气与信息工程学院", "通信工程", "0807" },
                            { "电气与信息工程学院","广播电视工程", "0807" }, 
                            { "电气与信息工程学院","电气工程及其自动化", "0806" }, 
                            { "文学与新闻传播学院","汉语言文学", "0501" },
                            { "文学与新闻传播学院","广播电视学", "0503" },
                            { "文学与新闻传播学院","网络与新媒体", "0503" },
                            { "化学与材料工程学院","化学", "0703" },
                            { "化学与材料工程学院","科学教育", "0401" },
                            { "化学与材料工程学院", "制药工程", "0813" },
                            { "化学与材料工程学院","材料化学", "0804" },
                            { "化学与材料工程学院","材料科学与工程", "0804" },
                            { "设计艺术学院","工业设计", "0802" },
                            { "设计艺术学院","视觉传达设计", "1305" },
                            { "设计艺术学院","环境设计", "1305" },
                            { "设计艺术学院","产品设计", "1305" },
                            { "设计艺术学院","服装与服饰设计", "1305" },
                            { "设计艺术学院","数字媒体艺术", "1305" },
                            { "法学与公共管理学院","公共事业管理", "1204" },
                            { "法学与公共管理学院","社会工作", "0303" },
                            { "法学与公共管理学院","土地资源管理", "1204" },
                            { "法学与公共管理学院","法学", "0301" },
                            { "商学院","旅游管理", "1209" },
                            { "商学院", "物流管理", "1206" },
                            { "商学院","财务管理", "1202" },
                            { "商学院","酒店管理", "1209" },
                            { "经济学院","国际经济与贸易", "0204" },
                            { "经济学院","投资学", "0203" },
                            { "计算机科学与工程学院","计算机科学与技术", "0809" },
                            { "计算机科学与工程学院","网络工程", "0809" }, 
                            { "计算机科学与工程学院","软件工程", "0809" }, 
                            { "教育科学学院","小学教育", "0401" }, 
                            { "教育科学学院","学前教育", "0401" }, 
                            { "教育科学学院","人文教育", "0401" }, 
                            { "马克思主义学院", "思想政治教育", "0305" },
                            { "风景园林学院", "园林", "0905" },
                            { "风景园林学院", "风景园林", "0828" },
                            { "机械与光电物理学院","物理学", "0702" },
                            { "机械与光电物理学院","光电信息科学与工程", "0807" } 
                            };
        public XuenumberSelectForm()
        {
            InitializeComponent();

            for (int i = 0; i < xueList.Length / 3; i++)
            {
                listViewEx1.Items.Add(new ListViewItem(new string[] { xueList[i, 0], xueList[i, 1], xueList[i, 2] }));
            }
        }

        private void textBoxX1_TextChanged(object sender, EventArgs e)
        {
            if (xueList == null)
                return;
            CompareInfo comp = CultureInfo.InvariantCulture.CompareInfo;
            this.listViewEx1.BeginUpdate();
            this.listViewEx1.Items.Clear();
            for (int i = 0; i < xueList.Length / 3; i++)
            {

                listViewEx1.Items.Add(new ListViewItem(new string[] { xueList[i, 0], xueList[i, 1], xueList[i, 2] }));
            }
            this.listViewEx1.EndUpdate();
            if (this.textBoxX1.Text != "")
            {
                ListView.ListViewItemCollection list = this.listViewEx1.Items;
                this.listViewEx1.BeginUpdate();
                foreach (ListViewItem nowItem in list)
                {
                    if (comp.IndexOf(nowItem.SubItems[0].Text, this.textBoxX1.Text, CompareOptions.IgnoreCase) == -1 && comp.IndexOf(nowItem.SubItems[1].Text, this.textBoxX1.Text, CompareOptions.IgnoreCase) == -1 && comp.IndexOf(nowItem.SubItems[2].Text, this.textBoxX1.Text, CompareOptions.IgnoreCase) == -1)
                    {
                        nowItem.Remove();
                    }
                }
                this.listViewEx1.EndUpdate();
            }
        }

        private void buttonX3_Click(object sender, EventArgs e)
        {
            if (listViewEx1.SelectedItems.Count != 0)
            {
                xueYuan = listViewEx1.SelectedItems[0].SubItems[0].Text;
                zhuanYe = listViewEx1.SelectedItems[0].SubItems[1].Text;
                xueNum = listViewEx1.SelectedItems[0].SubItems[2].Text;
            }
            else
            {
                this.DialogResult = System.Windows.Forms.DialogResult.No;
            }
        }

        private void buttonX5_Click(object sender, EventArgs e)
        {
            textBoxX1.Text = "";
        }

        private void listViewEx1_DoubleClick(object sender, EventArgs e)
        {
            if (listViewEx1.SelectedItems.Count != 0)
            {
                xueNum = listViewEx1.SelectedItems[0].SubItems[2].Text;
                zhuanYe = listViewEx1.SelectedItems[0].SubItems[1].Text;
                xueYuan = listViewEx1.SelectedItems[0].SubItems[0].Text;
                this.DialogResult = System.Windows.Forms.DialogResult.Yes;
            }
        }
    }
}
