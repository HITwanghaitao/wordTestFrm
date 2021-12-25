using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using wordTestFrm.Model;

namespace wordTestFrm.ControlTool
{
    public partial class ucCellAlignment : UserControl
    {
        public Enum_CellAlignment cellAlignment = Enum_CellAlignment.CenterMiddle;
        private Pen pen_Normal = new Pen(Color.Black,3.0f);
        private Pen pen_Hover = new Pen(Color.Blue, 3.0f);
        private Pen pen = new Pen(Color.Red, 2.0f);
        public Label lblSelected = null;
        public event labelClickFunc LabelClick;
        public Size sizeStrand = new Size(176, 176);
        public Size sizeStrandard_label = new Size(50, 50);
        public Font font = new Font("微软雅黑", 8);

        public delegate void labelClickFunc(object sender, EventArgs e);

        public ucCellAlignment()
        {
            InitializeComponent();
            foreach (Control item in this.Controls)
            {
                if (item is Label)
                {
                    item.Font = this.font;
                }
            }
        }

        /// <summary>
        /// 设置位置
        /// </summary>
        /// <param name="cellAlignment"></param>
        public void SetAlignment(Enum_CellAlignment cellAlignment)
        {
            this.cellAlignment = cellAlignment;
        }

        private void label_Click(object sender, EventArgs e)
        {
            Control control = (Control)sender;
            this.cellAlignment = (Enum_CellAlignment)Convert.ToInt16(control.Tag);
            this.Refresh();
            if(this.LabelClick!=null)
            this.LabelClick.Invoke(sender, e);
            //this.Visible = false;
        }

        private void label_Paint(object sender, PaintEventArgs e)
        {
            Control c = (Control)sender;
            Enum_CellAlignment TempCellAlignment = (Enum_CellAlignment)Convert.ToInt16( c.Tag);
            string flag = "字体";
            Rectangle rectangle = new Rectangle(new Point(4, 4), new Size(c.Size.Width-8,c.Size.Height-8));
            Graphics g = e.Graphics;
            StringFormat sf = new StringFormat();
            switch(TempCellAlignment)
            {
                case Enum_CellAlignment.LeftUp:
                    sf.Alignment = StringAlignment.Near;
                    sf.LineAlignment = StringAlignment.Near;
                    break;
                case Enum_CellAlignment.LeftMiddle:
                    sf.Alignment = StringAlignment.Near;
                    sf.LineAlignment = StringAlignment.Center;
                    break;
                case Enum_CellAlignment.LeftBottom:
                    sf.Alignment = StringAlignment.Near;
                    sf.LineAlignment = StringAlignment.Far;
                    break;
                case Enum_CellAlignment.CenterUp:
                    sf.Alignment = StringAlignment.Center;
                    sf.LineAlignment = StringAlignment.Near;
                    break;
                case Enum_CellAlignment.CenterMiddle:
                    sf.Alignment = StringAlignment.Center;
                    sf.LineAlignment = StringAlignment.Center;
                    break;
                case Enum_CellAlignment.CenterBottom:
                    sf.Alignment = StringAlignment.Center;
                    sf.LineAlignment = StringAlignment.Far;
                    break;
                case Enum_CellAlignment.RightUp:
                    sf.Alignment = StringAlignment.Far;
                    sf.LineAlignment = StringAlignment.Near;
                    break;
                case Enum_CellAlignment.RightMiddle:
                    sf.Alignment = StringAlignment.Far;
                    sf.LineAlignment = StringAlignment.Center;
                    break;
                case Enum_CellAlignment.RightBottom:
                    sf.Alignment = StringAlignment.Far;
                    sf.LineAlignment = StringAlignment.Far;
                   
                    break;
            }

            g.DrawString(flag, c.Font,Brushes.Black, rectangle, sf);

            if(this.cellAlignment== TempCellAlignment)
            {
                g.DrawRectangle(pen, new Rectangle(1, 1, c.Width-2, c.Height-2));
            }
            else if(this.lblSelected!=null && this.lblSelected==c)
            {
                g.DrawRectangle(pen_Hover, new Rectangle(1, 1, c.Width - 3, c.Height - 4));
            }
            else
            {
                g.DrawRectangle(pen_Normal, new Rectangle(1, 1, c.Width - 3, c.Height - 4));
            }

        }

        private void label1_MouseMove(object sender, MouseEventArgs e)
        {
            this.lblSelected = (Label)sender;
            this.lblSelected.Cursor = Cursors.Hand;
            this.Refresh();

        }

        private void label1_MouseLeave(object sender, EventArgs e)
        {
            this.lblSelected = null;
            this.Refresh();
        }

        private void ucCellAlignment_SizeChanged(object sender, EventArgs e)
        {
            //double widthSeed = (double)this.Width / (double)this.sizeStrand.Width;
            //double heightSeed = (double)this.Height / (double)this.sizeStrand.Height;
            
            int space = 2;
            double tmp_width = (this.Width - space) / 3;
            double tmp_Height = (this.Height - space) / 3;
            foreach (Control item in this.Controls)
            {
                if(item is Label)
                {
                    // item.Size = new Size((int)(sizeStrandard_label.Width * widthSeed)-2, (int)(sizeStrandard_label.Height * heightSeed)-2);
                    item.Size = new Size((int)tmp_width, (int)tmp_Height);
                    int tag = Convert.ToInt32(item.Tag);
                    switch(tag)
                    {
                        case 1:
                            item.Location = new Point(0, 0);
                            break;
                        case 2:
                            item.Location = new Point(item.Width + space, 0);
                            break;
                        case 3:
                            item.Location = new Point( 2 *(item.Width + space)-1, 0);
                            break;
                        case 4:
                            item.Location = new Point(0, item.Height+space);
                            break;
                        case 5:
                            item.Location = new Point(item.Width + space, item.Height + space);
                            break;
                        case 6:
                            item.Location = new Point( 2 * (item.Width + space)-1, item.Height + space);
                            break;
                        case 7://左下
                            item.Location = new Point(0,  2*(item.Height + space));
                            break;
                        case 8://中下
                            item.Location = new Point(  (item.Width + space),  2 * (item.Height + space));
                            break;
                        case 9://右下
                            item.Location = new Point( 2 * (item.Width + space)-1,  2 * (item.Height + space));
                            break;
                    }
                }
            }
        }
    }
}
