using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace DocumentSearch
{
    public partial class BookMessageForm : Form
    {
        public BookMessageForm(string link,bool note = true)
        {
            InitializeComponent();
            url = "http://172.28.135.39:8089/opac/ckgc.jsp?kzh=" + link;
            groupBox1.Enabled = note;
        }
        public string url;
        private void BookMessageForm_Load(object sender, EventArgs e)
        {
            browserUrl(url);
        }
        private void browserUrl(string url)
        {
            webBrowser1.Navigate(url);
        }

        private void webBrowser1_Navigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            if (e.Url != new Uri(url))
            {
                e.Cancel = true;
            }
        }

        private void webBrowser1_NewWindow(object sender, CancelEventArgs e)
        {
            e.Cancel = true;
        }
    }
}
