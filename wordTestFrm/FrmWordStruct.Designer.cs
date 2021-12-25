namespace wordTestFrm
{
    partial class FrmWordStruct
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
            this.tvStruct = new System.Windows.Forms.TreeView();
            this.txtContent = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.lblStyle = new System.Windows.Forms.Label();
            this.chkOnlyHeading = new System.Windows.Forms.CheckBox();
            this.lblFont = new System.Windows.Forms.Label();
            this.lblColor = new System.Windows.Forms.Label();
            this.fontDialog1 = new System.Windows.Forms.FontDialog();
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.btn_Apply = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // tvStruct
            // 
            this.tvStruct.Location = new System.Drawing.Point(13, 13);
            this.tvStruct.Name = "tvStruct";
            this.tvStruct.Size = new System.Drawing.Size(256, 425);
            this.tvStruct.TabIndex = 0;
            this.tvStruct.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.tvStruct_NodeMouseClick);
            this.tvStruct.NodeMouseDoubleClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.tvStruct_NodeMouseDoubleClick);
            // 
            // txtContent
            // 
            this.txtContent.Location = new System.Drawing.Point(367, 13);
            this.txtContent.Multiline = true;
            this.txtContent.Name = "txtContent";
            this.txtContent.Size = new System.Drawing.Size(421, 304);
            this.txtContent.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(367, 336);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "标题样式";
            // 
            // lblStyle
            // 
            this.lblStyle.AutoSize = true;
            this.lblStyle.Location = new System.Drawing.Point(490, 336);
            this.lblStyle.Name = "lblStyle";
            this.lblStyle.Size = new System.Drawing.Size(53, 12);
            this.lblStyle.TabIndex = 3;
            this.lblStyle.Text = "标题样式";
            // 
            // chkOnlyHeading
            // 
            this.chkOnlyHeading.AutoSize = true;
            this.chkOnlyHeading.Checked = true;
            this.chkOnlyHeading.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkOnlyHeading.Location = new System.Drawing.Point(369, 379);
            this.chkOnlyHeading.Name = "chkOnlyHeading";
            this.chkOnlyHeading.Size = new System.Drawing.Size(72, 16);
            this.chkOnlyHeading.TabIndex = 4;
            this.chkOnlyHeading.Text = "只有标题";
            this.chkOnlyHeading.UseVisualStyleBackColor = true;
            this.chkOnlyHeading.CheckedChanged += new System.EventHandler(this.chkOnlyHeading_CheckedChanged);
            // 
            // lblFont
            // 
            this.lblFont.AutoSize = true;
            this.lblFont.Location = new System.Drawing.Point(567, 364);
            this.lblFont.Name = "lblFont";
            this.lblFont.Size = new System.Drawing.Size(29, 12);
            this.lblFont.TabIndex = 29;
            this.lblFont.Text = "字体";
            this.lblFont.Click += new System.EventHandler(this.lblFont_Click);
            // 
            // lblColor
            // 
            this.lblColor.AutoSize = true;
            this.lblColor.Location = new System.Drawing.Point(490, 364);
            this.lblColor.Name = "lblColor";
            this.lblColor.Size = new System.Drawing.Size(53, 12);
            this.lblColor.TabIndex = 28;
            this.lblColor.Text = "字体颜色";
            this.lblColor.Click += new System.EventHandler(this.lblColor_Click);
            // 
            // btn_Apply
            // 
            this.btn_Apply.Location = new System.Drawing.Point(480, 396);
            this.btn_Apply.Name = "btn_Apply";
            this.btn_Apply.Size = new System.Drawing.Size(75, 23);
            this.btn_Apply.TabIndex = 30;
            this.btn_Apply.Text = "应用";
            this.btn_Apply.UseVisualStyleBackColor = true;
            this.btn_Apply.Click += new System.EventHandler(this.btn_Apply_Click);
            // 
            // FrmWordStruct
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 495);
            this.Controls.Add(this.btn_Apply);
            this.Controls.Add(this.lblFont);
            this.Controls.Add(this.lblColor);
            this.Controls.Add(this.chkOnlyHeading);
            this.Controls.Add(this.lblStyle);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtContent);
            this.Controls.Add(this.tvStruct);
            this.Name = "FrmWordStruct";
            this.Text = "FrmWordStruct";
            this.Load += new System.EventHandler(this.FrmWordStruct_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TreeView tvStruct;
        private System.Windows.Forms.TextBox txtContent;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblStyle;
        private System.Windows.Forms.CheckBox chkOnlyHeading;
        private System.Windows.Forms.Label lblFont;
        private System.Windows.Forms.Label lblColor;
        private System.Windows.Forms.FontDialog fontDialog1;
        private System.Windows.Forms.ColorDialog colorDialog1;
        private System.Windows.Forms.Button btn_Apply;
    }
}