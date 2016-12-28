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
    public partial class ThesisEditerAboutForm : Form
    {
        public ThesisEditerAboutForm()
        {
            InitializeComponent();
            label1.Text += "v" + MainForm.thesisEditerVersion;
            label3.Text = "Office版本：" + MainForm.officeVersion;
        }
    }
}
