using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DocumentSearch
{
    public partial class InputSearchForm : Form
    {
        public InputSearchForm()
        {
            InitializeComponent();
            comboBox1.SelectedIndex = 0;
        }
        public string word;
        public int type;
        private void button1_Click(object sender, EventArgs e)
        {
            word = textBox1.Text;
            type = comboBox1.SelectedIndex;
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                word = textBox1.Text;
                type = comboBox1.SelectedIndex;
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
            }
        }
    }
}
