namespace wordTestFrm.ControlTool
{
    partial class ucFont
    {
        /// <summary> 
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary> 
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btnColor = new System.Windows.Forms.Button();
            this.btnFont = new System.Windows.Forms.Button();
            this.lblFont = new System.Windows.Forms.Label();
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.fontDialog1 = new System.Windows.Forms.FontDialog();
            this.lblContent = new System.Windows.Forms.Label();
            this.pnlClose = new System.Windows.Forms.Panel();
            this.pnlSpread = new System.Windows.Forms.Panel();
            this.SuspendLayout();
            // 
            // btnColor
            // 
            this.btnColor.BackColor = System.Drawing.Color.Black;
            this.btnColor.Location = new System.Drawing.Point(34, 4);
            this.btnColor.Name = "btnColor";
            this.btnColor.Size = new System.Drawing.Size(23, 23);
            this.btnColor.TabIndex = 0;
            this.btnColor.Text = "button1";
            this.btnColor.UseVisualStyleBackColor = false;
            this.btnColor.Click += new System.EventHandler(this.btnColor_Click);
            this.btnColor.MouseHover += new System.EventHandler(this.btnFont_MouseHover);
            // 
            // btnFont
            // 
            this.btnFont.Location = new System.Drawing.Point(3, 4);
            this.btnFont.Name = "btnFont";
            this.btnFont.Size = new System.Drawing.Size(25, 23);
            this.btnFont.TabIndex = 1;
            this.btnFont.Text = "字";
            this.btnFont.UseVisualStyleBackColor = true;
            this.btnFont.Click += new System.EventHandler(this.btnFont_Click);
            this.btnFont.MouseHover += new System.EventHandler(this.btnFont_MouseHover);
            // 
            // lblFont
            // 
            this.lblFont.AutoSize = true;
            this.lblFont.Location = new System.Drawing.Point(63, 9);
            this.lblFont.Name = "lblFont";
            this.lblFont.Size = new System.Drawing.Size(29, 12);
            this.lblFont.TabIndex = 2;
            this.lblFont.Text = "字体";
            this.lblFont.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblFont.MouseHover += new System.EventHandler(this.btnFont_MouseHover);
            // 
            // lblContent
            // 
            this.lblContent.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.lblContent.Location = new System.Drawing.Point(0, 31);
            this.lblContent.Name = "lblContent";
            this.lblContent.Size = new System.Drawing.Size(141, 51);
            this.lblContent.TabIndex = 3;
            this.lblContent.Text = "字体";
            this.lblContent.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lblContent.MouseHover += new System.EventHandler(this.btnFont_MouseHover);
            // 
            // pnlClose
            // 
            this.pnlClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.pnlClose.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlClose.Location = new System.Drawing.Point(123, 32);
            this.pnlClose.Name = "pnlClose";
            this.pnlClose.Size = new System.Drawing.Size(15, 16);
            this.pnlClose.TabIndex = 4;
            this.pnlClose.Visible = false;
            this.pnlClose.Click += new System.EventHandler(this.pnlClose_Click);
            this.pnlClose.Paint += new System.Windows.Forms.PaintEventHandler(this.pnlClose_Paint);
            this.pnlClose.MouseLeave += new System.EventHandler(this.pnlClose_MouseLeave);
            this.pnlClose.MouseHover += new System.EventHandler(this.pnlClose_MouseHover);
            // 
            // pnlSpread
            // 
            this.pnlSpread.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.pnlSpread.Location = new System.Drawing.Point(123, 5);
            this.pnlSpread.Name = "pnlSpread";
            this.pnlSpread.Size = new System.Drawing.Size(15, 16);
            this.pnlSpread.TabIndex = 5;
            this.pnlSpread.Visible = false;
            this.pnlSpread.Click += new System.EventHandler(this.pnlSpread_Click);
            this.pnlSpread.Paint += new System.Windows.Forms.PaintEventHandler(this.pnlSpread_Paint);
            this.pnlSpread.MouseLeave += new System.EventHandler(this.pnlClose_MouseLeave);
            this.pnlSpread.MouseHover += new System.EventHandler(this.pnlClose_MouseHover);
            // 
            // ucFont
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Controls.Add(this.pnlSpread);
            this.Controls.Add(this.pnlClose);
            this.Controls.Add(this.lblContent);
            this.Controls.Add(this.lblFont);
            this.Controls.Add(this.btnFont);
            this.Controls.Add(this.btnColor);
            this.Name = "ucFont";
            this.Size = new System.Drawing.Size(141, 82);
            this.Load += new System.EventHandler(this.ucFont2_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnColor;
        private System.Windows.Forms.Button btnFont;
        private System.Windows.Forms.Label lblFont;
        private System.Windows.Forms.ColorDialog colorDialog1;
        private System.Windows.Forms.FontDialog fontDialog1;
        private System.Windows.Forms.Label lblContent;
        private System.Windows.Forms.Panel pnlClose;
        private System.Windows.Forms.Panel pnlSpread;
    }
}
