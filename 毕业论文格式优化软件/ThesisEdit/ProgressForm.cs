using System;
using System.Windows.Forms;

namespace thesisEditer
{
    public partial class ProgressForm : Form
    {
        private int _current = 0;
        /// <summary>
        /// 当前值
        /// </summary>
        public int Current
        {
            get { return _current; }
            set
            {
                _current = value;
                AddValue();
            }
        }
        private int _max = 100;
        /// <summary>
        /// 最大值
        /// </summary>
        public int Max
        {
            get { return _max; }
            set 
            { 
                _max = value;
                this.progressBar1.Maximum = (_max - 1) * 10;
                this.progressBar1.Value = 0;
            }
        }
        MainForm father;
        public ProgressForm(MainForm main)
        {
            father = main;
            InitializeComponent();
            this.label1.Text = "正在写入论文信息，请稍等...";
        }

        /// <summary>
        /// 给进度条加值的方法
        /// </summary>
        private void AddValue()
        {
            if (this.progressBar1.Style == ProgressBarStyle.Marquee)
                return;
            if (_current > 0)
            {
                this.label1.Text = "正在写入论文正文，请稍等...";
            }
            this.progressBar1.PerformStep();
            if (this._current * 10 > this.progressBar1.Maximum)
            {
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
                this.Close();

            }

        }
        /// <summary>
        /// 取消按钮的方法
        /// </summary>
        private void Cancel()
        {
            father.CreateThesisIsRuning = false;
            progressBar1.Style = ProgressBarStyle.Marquee;
            label1.Text = "取消中，请稍等...";
            while (true)
            {
                if (!father.CreateThesisTh.IsAlive)
                    break;
                Application.DoEvents();
            }
        }
        /// <summary>
        /// 取消按钮事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonX1_Click(object sender, EventArgs e)
        {
            Cancel();
        }
        



    }
}
