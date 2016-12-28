using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;

namespace thesisEditer
{
    public partial class SettingsForm : Form
    {
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
        
        DevComponents.Editors.ComboItem[] fontStyleItems;
        DevComponents.Editors.ComboItem[] fontSizeItems;
        MainForm father = null;
        public string settingsPath = System.AppDomain.CurrentDomain.BaseDirectory + @"\Settings.ini";
        public SettingsForm(int[] a,int[] b,MainForm f)
        {
            InitializeComponent();
            fontStyleItems = new DevComponents.Editors.ComboItem[4];
            //创建项
            fontStyleItems[0] = new DevComponents.Editors.ComboItem("黑体");
            fontStyleItems[1] = new DevComponents.Editors.ComboItem("仿宋");
            fontStyleItems[2] = new DevComponents.Editors.ComboItem("宋体");
            fontStyleItems[3] = new DevComponents.Editors.ComboItem("Times New Roman");
            fontSizeItems = new DevComponents.Editors.ComboItem[10];
            fontSizeItems[0] = new DevComponents.Editors.ComboItem("二号");
            fontSizeItems[1] = new DevComponents.Editors.ComboItem("小二号");
            fontSizeItems[2] = new DevComponents.Editors.ComboItem("三号");
            fontSizeItems[3] = new DevComponents.Editors.ComboItem("小三号");
            fontSizeItems[4] = new DevComponents.Editors.ComboItem("四号");
            fontSizeItems[5] = new DevComponents.Editors.ComboItem("小四号");
            fontSizeItems[6] = new DevComponents.Editors.ComboItem("五号");
            fontSizeItems[7] = new DevComponents.Editors.ComboItem("小五号");
            fontSizeItems[8] = new DevComponents.Editors.ComboItem("六号");
            fontSizeItems[9] = new DevComponents.Editors.ComboItem("小六号");

            //添加项
            this.comboBoxEx1.Items.AddRange(fontStyleItems);
            this.comboBoxEx3.Items.AddRange(fontStyleItems);
            this.comboBoxEx5.Items.AddRange(fontStyleItems);
            this.comboBoxEx7.Items.AddRange(fontStyleItems);
            this.comboBoxEx9.Items.AddRange(fontStyleItems);
            this.comboBoxEx11.Items.AddRange(fontStyleItems);
            this.comboBoxEx13.Items.AddRange(fontStyleItems);
            this.comboBoxEx15.Items.AddRange(fontStyleItems);
            this.comboBoxEx17.Items.AddRange(fontStyleItems);
            this.comboBoxEx19.Items.AddRange(fontStyleItems);
            this.comboBoxEx21.Items.AddRange(fontStyleItems);
            this.comboBoxEx2.Items.AddRange(fontSizeItems);
            this.comboBoxEx4.Items.AddRange(fontSizeItems);
            this.comboBoxEx6.Items.AddRange(fontSizeItems);
            this.comboBoxEx8.Items.AddRange(fontSizeItems);
            this.comboBoxEx10.Items.AddRange(fontSizeItems);
            this.comboBoxEx12.Items.AddRange(fontSizeItems);
            this.comboBoxEx14.Items.AddRange(fontSizeItems);
            this.comboBoxEx16.Items.AddRange(fontSizeItems);
            this.comboBoxEx18.Items.AddRange(fontSizeItems);
            this.comboBoxEx20.Items.AddRange(fontSizeItems);
            this.comboBoxEx22.Items.AddRange(fontSizeItems);
            initSettings(a, b);
            father = f;

        }
        private void initSettings(int[] a, int[] b)
        {


            //读取项
            this.comboBoxEx1.SelectedIndex = a[0];
            this.comboBoxEx3.SelectedIndex = a[1];
            this.comboBoxEx5.SelectedIndex = a[2];
            this.comboBoxEx7.SelectedIndex = a[3];
            this.comboBoxEx9.SelectedIndex = a[4];
            this.comboBoxEx11.SelectedIndex = a[5];
            this.comboBoxEx13.SelectedIndex = a[6];
            this.comboBoxEx15.SelectedIndex = a[7];
            this.comboBoxEx17.SelectedIndex = a[8];
            this.comboBoxEx19.SelectedIndex = a[9];
            this.comboBoxEx21.SelectedIndex = a[10];
            this.comboBoxEx2.SelectedIndex = b[0];
            this.comboBoxEx4.SelectedIndex = b[1];
            this.comboBoxEx6.SelectedIndex = b[2];
            this.comboBoxEx8.SelectedIndex = b[3];
            this.comboBoxEx10.SelectedIndex = b[4];
            this.comboBoxEx12.SelectedIndex = b[5];
            this.comboBoxEx14.SelectedIndex = b[6];
            this.comboBoxEx16.SelectedIndex = b[7];
            this.comboBoxEx18.SelectedIndex = b[8];
            this.comboBoxEx20.SelectedIndex = b[9];
            this.comboBoxEx22.SelectedIndex = b[10];
        }
        private void SettingsForm_Load(object sender, EventArgs e)
        {


            
        }
        private void saveSettings(string path)
        {
            WritePrivateProfileString("Format", "subject_font_style", comboBoxEx1.SelectedIndex.ToString(), path);
            WritePrivateProfileString("Format", "big_font_style", comboBoxEx3.SelectedIndex.ToString(), path);
            WritePrivateProfileString("Format", "small_font_style", comboBoxEx5.SelectedIndex.ToString(), path);
            WritePrivateProfileString("Format", "text_font_style", comboBoxEx7.SelectedIndex.ToString(), path);
            WritePrivateProfileString("Format", "page_font_style", comboBoxEx9.SelectedIndex.ToString(), path);
            WritePrivateProfileString("Format", "reference_font_style", comboBoxEx11.SelectedIndex.ToString(), path);
            WritePrivateProfileString("Format", "mulu_font_style", comboBoxEx13.SelectedIndex.ToString(), path);
            WritePrivateProfileString("Format", "list_font_style", comboBoxEx15.SelectedIndex.ToString(), path);
            WritePrivateProfileString("Format", "picture_font_style", comboBoxEx17.SelectedIndex.ToString(), path);
            WritePrivateProfileString("Format", "table_font_style", comboBoxEx19.SelectedIndex.ToString(), path);
            WritePrivateProfileString("Format", "intable_font_style", comboBoxEx21.SelectedIndex.ToString(), path);
            
            WritePrivateProfileString("Format", "subject_font_size", comboBoxEx2.SelectedIndex.ToString(), path);
            WritePrivateProfileString("Format", "big_font_size", comboBoxEx4.SelectedIndex.ToString(), path);
            WritePrivateProfileString("Format", "small_font_size", comboBoxEx6.SelectedIndex.ToString(), path);
            WritePrivateProfileString("Format", "text_font_size", comboBoxEx8.SelectedIndex.ToString(), path);
            WritePrivateProfileString("Format", "page_font_size", comboBoxEx10.SelectedIndex.ToString(), path);
            WritePrivateProfileString("Format", "reference_font_size", comboBoxEx12.SelectedIndex.ToString(), path);
            WritePrivateProfileString("Format", "mulu_font_size", comboBoxEx14.SelectedIndex.ToString(), path);
            WritePrivateProfileString("Format", "list_font_size", comboBoxEx16.SelectedIndex.ToString(), path);
            WritePrivateProfileString("Format", "picture_font_size", comboBoxEx18.SelectedIndex.ToString(), path);
            WritePrivateProfileString("Format", "table_font_size", comboBoxEx20.SelectedIndex.ToString(), path);
            WritePrivateProfileString("Format", "intable_font_size", comboBoxEx22.SelectedIndex.ToString(), path);
        }
        private void readSettingsFile()
        {
            try
            {
                this.openFileDialog1.RestoreDirectory = true;
                if (this.openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    File.Copy(this.openFileDialog1.FileName, MainForm.settingsPath, true);

                    father.readSettings();
                    initSettings(father.settingsStyleList, father.settingsSizeList);
                }
            }
            catch (Exception e)
            {
                new ErrorForm(e.ToString()).ShowDialog();
            }
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            saveSettings(settingsPath);
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonX3_Click(object sender, EventArgs e)
        {
            
            this.saveFileDialog1.RestoreDirectory = true;
            if (this.saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                saveSettings(this.saveFileDialog1.FileName);
            }
            
        }

        private void buttonX4_Click(object sender, EventArgs e)
        {
            readSettingsFile();
        }

        
    }
}
