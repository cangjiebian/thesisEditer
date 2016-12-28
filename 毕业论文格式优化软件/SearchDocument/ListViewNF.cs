using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

namespace DocumentSearch
{
    class ListViewNF : System.Windows.Forms.ListView
    {
        public ListViewNF()
        {
            // 开启双缓冲
            this.SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);

            // Enable the OnNotifyMessage event so we get a chance to filter out 
            // Windows messages before they get to the form's WndProc
            this.SetStyle(ControlStyles.EnableNotifyMessage, true);
        }

        protected override void OnNotifyMessage(Message m)
        {
            //Filter out the WM_ERASEBKGND message
            if (m.Msg != 0x14)
            {
                base.OnNotifyMessage(m);
            }

        }
        protected override void OnPrint(PaintEventArgs e)
        {
            base.OnPrint(e);
            
        }
        /*
        private void DrawText(ListViewItem item)
        {
            Graphics g = this.listView2.CreateGraphics();
            int NodeOverImageWidth = 10;
            int LeftPos, RightPos;
            LeftPos = item.Bounds.Left - NodeOverImageWidth;
            RightPos = this.listView2.Width - 4;
            Point[] LeftTriangle = new Point[5]{
												   new Point(LeftPos, item.Bounds.Top - 4),
												   new Point(LeftPos, item.Bounds.Top + 4),
												   new Point(LeftPos + 4, item.Bounds.Y),
												   new Point(LeftPos + 4, item.Bounds.Top - 1),
												   new Point(LeftPos, item.Bounds.Top - 5)};
            Point[] RightTriangle = new Point[5]{
													new Point(RightPos, item.Bounds.Top - 4),
													new Point(RightPos, item.Bounds.Top + 4),
													new Point(RightPos - 4, item.Bounds.Y),
													new Point(RightPos - 4, item.Bounds.Top - 1),
													new Point(RightPos, item.Bounds.Top - 5)};
            g.FillPolygon(System.Drawing.Brushes.Black, LeftTriangle);
            g.FillPolygon(System.Drawing.Brushes.Black, RightTriangle);
            g.DrawLine(new System.Drawing.Pen(Color.Red, 2), new Point(LeftPos, item.Bounds.Top), new Point(RightPos, item.Bounds.Top));
        }
         */
    }
}
