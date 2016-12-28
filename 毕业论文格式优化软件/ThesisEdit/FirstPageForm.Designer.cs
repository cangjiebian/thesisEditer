namespace thesisEditer
{
    partial class FirstPageForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FirstPageForm));
            this.axFramerControl1 = new AxDSOFramer.AxFramerControl();
            ((System.ComponentModel.ISupportInitialize)(this.axFramerControl1)).BeginInit();
            this.SuspendLayout();
            // 
            // axFramerControl1
            // 
            this.axFramerControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.axFramerControl1.Enabled = true;
            this.axFramerControl1.Location = new System.Drawing.Point(0, 0);
            this.axFramerControl1.Name = "axFramerControl1";
            this.axFramerControl1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axFramerControl1.OcxState")));
            this.axFramerControl1.Size = new System.Drawing.Size(284, 262);
            this.axFramerControl1.TabIndex = 2;
            // 
            // FirstPageForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 262);
            this.Controls.Add(this.axFramerControl1);
            this.DoubleBuffered = true;
            this.Name = "FirstPageForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "FirstPageForm";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FirstPageForm_FormClosed);
            ((System.ComponentModel.ISupportInitialize)(this.axFramerControl1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        public AxDSOFramer.AxFramerControl axFramerControl1;

    }
}