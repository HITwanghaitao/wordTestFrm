using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Text;

namespace wordTestFrm.ControlTool
{
    public partial class ucFont : UserControl
    {
        public Font fontSelect = null;
        public Color fontColorSelect = Color.Black;
        public Pen penClose = new Pen(Brushes.Black, 1.5f);
        private Point[] pointsSpread = new Point[3];
        private Point[] pointsClosed = new Point[3];
        int heightSmall = 30;
        int heightMax = 90;
        bool isSpread = false;
        public Label lblOther = null;


        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
        }

        /// <summary>
        /// 加载字体 设置
        /// </summary>
        /// <param name="font"></param>
        /// <param name="color"></param>
        public void SettingsControl(Font font,Color color)
        {
            lblContent.Font= this.fontSelect = font;
            lblContent.ForeColor= btnColor.BackColor=this.fontColorSelect = color;
            lblFont.Text = this.fontSelect.Name + " " +CommonMethods.GetFontSize(this.fontSelect.Size);
            if(this.lblOther!=null)
            {
                lblOther.Font = this.fontSelect;
                lblOther.ForeColor = this.fontColorSelect;
            }
            this.Refresh();
        }

        protected override void OnMouseHover(EventArgs e)
        {
            base.OnMouseHover(e);
            this.btnColor.Cursor=this.btnFont.Cursor= this.pnlClose.Cursor = Cursors.Hand;
        }

        public ucFont()
        {
            InitializeComponent();
            fontSelect = lblFont.Font;
            lblFont.Text = lblFont.Font.Name + "  " + lblFont.Font.Size;
            lblContent.Font = lblFont.Font;
            lblContent.ForeColor = lblFont.ForeColor;
            pointsSpread[0] = new Point(2,2 );
            pointsSpread[1] = new Point(pnlSpread.Width/2-1,pnlSpread.Height-4);
            pointsSpread[2] = new Point(pnlSpread.Width-4,2);
            pnlSpread.Cursor = pnlClose.Cursor = Cursors.Hand;

            pointsClosed[0] = new Point(pnlClose.Width / 2 - 1, 2);
            pointsClosed[1] = new Point(2, pnlSpread.Height - 4);
            pointsClosed[2] = new Point(pnlClose.Width - 4, pnlSpread.Height - 4);
        }

        private void btnFont_Click(object sender, EventArgs e)
        {
            if(fontDialog1.ShowDialog()==DialogResult.OK)
            {
                this.fontSelect= lblContent.Font = fontDialog1.Font;
                lblFont.Text = this.fontSelect.Name + " " +CommonMethods.GetFontSize(this.fontSelect.Size);
                lblContent.ForeColor = this.fontColorSelect;
                if (this.lblOther != null)
                {
                    lblOther.Font = this.fontSelect;
                    lblOther.ForeColor = this.fontColorSelect;
                }
            }
            //else
            //{
            //    this.SendToBack();
            //}
        }

        private void btnColor_Click(object sender, EventArgs e)
        {
            if(colorDialog1.ShowDialog()==DialogResult.OK)
            {
                this.fontColorSelect = btnColor.BackColor = colorDialog1.Color;
                lblContent.ForeColor = this.fontColorSelect;
                if (this.lblOther != null)
                {
                    lblOther.Font = this.fontSelect;
                    lblOther.ForeColor = this.fontColorSelect;
                }

            }
            //else
            //{
            //    this.SendToBack();
            //}
        }

        private void ucFont2_Load(object sender, EventArgs e)
        {
            this.Height = this.heightSmall;
            this.lblContent.Height = 0;
        }

        private void btnFont_MouseHover(object sender, EventArgs e)
        {
            
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
          
        }

        private void pnlClose_Click(object sender, EventArgs e)
        {
            this.Height = this.heightSmall;
            this.lblContent.Height = 0;
            this.SendToBack();
        }


        private void pnlClose_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            g.FillPolygon(Brushes.Black, pointsClosed);

            //g.DrawLine(penClose, new Point(3, 3), new Point(pnlClose.Width-5, pnlClose.Height-5));
            //g.DrawLine(penClose, new Point(pnlClose.Width-5, 3), new Point(3, pnlClose.Height-5));
        }

        private void pnlClose_MouseHover(object sender, EventArgs e)
        {
            ((Control)sender).BackColor = Color.Red;
        }

        private void pnlClose_MouseLeave(object sender, EventArgs e)
        {
            ((Control)sender).BackColor = Color.Transparent;
        }

        private void pnlSpread_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            if(isSpread)
            g.FillPolygon(Brushes.Black, pointsClosed);
            else
             g.FillPolygon(Brushes.Black, pointsSpread);
            //g.DrawLine(penClose, new Point(3, 3), new Point(pnlClose.Width - 5, pnlClose.Height - 5));
            //g.DrawLine(penClose, new Point(pnlClose.Width - 5, 3), new Point(3, pnlClose.Height - 5));
        }

        private void pnlSpread_Click(object sender, EventArgs e)
        {
            isSpread = !isSpread;
            if (isSpread)
            {
                //this.Height = this.heightMax;
                //this.lblContent.Height = 50;
                //this.BringToFront();
                if (this.lblOther != null)
                {
                    lblOther.Font = this.fontSelect;
                    lblOther.ForeColor = this.fontColorSelect;
                }
            }
            else
            {
                //this.Height = this.heightSmall;
                //this.lblContent.Height = 0;
                //this.SendToBack();
                if (this.lblOther != null)
                {
                    lblOther.Font = this.fontSelect;
                    lblOther.ForeColor = this.fontColorSelect;
                }
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            if (this.lblOther != null)
            {
                lblOther.Font = this.fontSelect;
                lblOther.ForeColor = this.fontColorSelect;
            }
           
            base.OnMouseMove(e);
        }
    }
}
