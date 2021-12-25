namespace wordTestFrm
{
    partial class FrmPDFToImgs
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
            this.btn_openPDF = new System.Windows.Forms.Button();
            this.btn_start = new System.Windows.Forms.Button();
            this.label11 = new System.Windows.Forms.Label();
            this.cbxDPI = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txtQuality = new System.Windows.Forms.TextBox();
            this.ucLoading1 = new wordTestFrm.ControlTool.ucLoading();
            this.tb_openPath = new System.Windows.Forms.TextBox();
            this.tb_resultPath = new System.Windows.Forms.TextBox();
            this.cbxPrase = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btn_openPDF
            // 
            this.btn_openPDF.Location = new System.Drawing.Point(34, 35);
            this.btn_openPDF.Name = "btn_openPDF";
            this.btn_openPDF.Size = new System.Drawing.Size(97, 23);
            this.btn_openPDF.TabIndex = 0;
            this.btn_openPDF.Text = "选择pdf文件";
            this.btn_openPDF.UseVisualStyleBackColor = true;
            this.btn_openPDF.Click += new System.EventHandler(this.btn_openPDF_Click);
            // 
            // btn_start
            // 
            this.btn_start.Location = new System.Drawing.Point(34, 222);
            this.btn_start.Name = "btn_start";
            this.btn_start.Size = new System.Drawing.Size(75, 23);
            this.btn_start.TabIndex = 2;
            this.btn_start.Text = "开始转换";
            this.btn_start.UseVisualStyleBackColor = true;
            this.btn_start.Click += new System.EventHandler(this.btn_start_Click);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(246, 126);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(239, 12);
            this.label11.TabIndex = 39;
            this.label11.Text = "打印质量 220  屏幕质量 150  邮件质量 96";
            // 
            // cbxDPI
            // 
            this.cbxDPI.FormattingEnabled = true;
            this.cbxDPI.Items.AddRange(new object[] {
            "220",
            "150",
            "96"});
            this.cbxDPI.Location = new System.Drawing.Point(100, 123);
            this.cbxDPI.Name = "cbxDPI";
            this.cbxDPI.Size = new System.Drawing.Size(121, 20);
            this.cbxDPI.TabIndex = 38;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(41, 170);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 12);
            this.label2.TabIndex = 37;
            this.label2.Text = "压缩质量";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(41, 126);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(23, 12);
            this.label1.TabIndex = 36;
            this.label1.Text = "DPI";
            // 
            // txtQuality
            // 
            this.txtQuality.Location = new System.Drawing.Point(100, 167);
            this.txtQuality.Name = "txtQuality";
            this.txtQuality.Size = new System.Drawing.Size(121, 21);
            this.txtQuality.TabIndex = 35;
            this.txtQuality.Text = "90";
            // 
            // ucLoading1
            // 
            this.ucLoading1.Location = new System.Drawing.Point(491, 123);
            this.ucLoading1.Name = "ucLoading1";
            this.ucLoading1.Size = new System.Drawing.Size(291, 159);
            this.ucLoading1.TabIndex = 41;
            this.ucLoading1.Visible = false;
            // 
            // tb_openPath
            // 
            this.tb_openPath.Location = new System.Drawing.Point(31, 64);
            this.tb_openPath.Multiline = true;
            this.tb_openPath.Name = "tb_openPath";
            this.tb_openPath.Size = new System.Drawing.Size(549, 53);
            this.tb_openPath.TabIndex = 42;
            this.tb_openPath.Text = "选择的文件目录";
            // 
            // tb_resultPath
            // 
            this.tb_resultPath.Location = new System.Drawing.Point(34, 260);
            this.tb_resultPath.Multiline = true;
            this.tb_resultPath.Name = "tb_resultPath";
            this.tb_resultPath.Size = new System.Drawing.Size(250, 165);
            this.tb_resultPath.TabIndex = 43;
            this.tb_resultPath.Text = "结果目录";
            // 
            // cbxPrase
            // 
            this.cbxPrase.FormattingEnabled = true;
            this.cbxPrase.Items.AddRange(new object[] {
            "最低质量",
            "打印质量",
            "完美转换"});
            this.cbxPrase.Location = new System.Drawing.Point(330, 167);
            this.cbxPrase.Name = "cbxPrase";
            this.cbxPrase.Size = new System.Drawing.Size(121, 20);
            this.cbxPrase.TabIndex = 44;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(246, 170);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 12);
            this.label3.TabIndex = 45;
            this.label3.Text = "转换方式";
            // 
            // FrmPDFToImgs
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.cbxPrase);
            this.Controls.Add(this.tb_resultPath);
            this.Controls.Add(this.tb_openPath);
            this.Controls.Add(this.ucLoading1);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.cbxDPI);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtQuality);
            this.Controls.Add(this.btn_start);
            this.Controls.Add(this.btn_openPDF);
            this.Name = "FrmPDFToImgs";
            this.Text = " ";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FrmPDFToImgs_FormClosed);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_openPDF;
        private System.Windows.Forms.Button btn_start;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.ComboBox cbxDPI;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtQuality;
        private ControlTool.ucLoading ucLoading1;
        private System.Windows.Forms.TextBox tb_openPath;
        private System.Windows.Forms.TextBox tb_resultPath;
        private System.Windows.Forms.ComboBox cbxPrase;
        private System.Windows.Forms.Label label3;
    }
}