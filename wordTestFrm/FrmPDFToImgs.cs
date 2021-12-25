using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using wordTestFrm.ControlTool;
using ThreadState = System.Threading.ThreadState;

namespace wordTestFrm
{
    public partial class FrmPDFToImgs : Form
    {
        Thread th = null;
        public FrmPDFToImgs()
        {
            InitializeComponent();
            cbxDPI.SelectedIndex = 1;
            cbxPrase.SelectedIndex = 2;
           
        }
        OpenFileDialog openfile;
        private void btn_openPDF_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.InitialDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            if (of.ShowDialog() == DialogResult.OK)
            {
                
                tb_openPath.Text = of.FileName;
                openfile = of;
            }
        }

        

        private void btn_start_Click(object sender, EventArgs e)
        {
            if(openfile==null)
            {
                MessageBox.Show("请先选择文件。", "提示");
                return;
            }
            this.btn_start.Enabled = false;
            this.btn_openPDF.Enabled = false;
            string fileName =openfile.SafeFileName;
            string filePath = openfile.FileName;
            string saveDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Pdf2Pic", 
                Path.GetFileNameWithoutExtension(openfile.SafeFileName));
            if (!Directory.Exists(saveDir))
            {
                Directory.CreateDirectory(saveDir);
            }
            tb_resultPath.Text = Path.Combine(saveDir,Path.GetFileNameWithoutExtension(fileName));

            ucLoading1.Visible = true;

            int dpi = int.Parse(cbxDPI.SelectedItem.ToString());
            int Quality = int.Parse(txtQuality.Text);
            PdfiumViewer.PdfRenderFlags pdfRenderFlags = retPicQuality(cbxPrase.SelectedIndex);
             PdfiumViewerTool pt = new PdfiumViewerTool();
            th = new Thread(() => {

                string savePathDir = pt.ConvertPDF2Pic(filePath, fileName, pdfRenderFlags, Quality, dpi);
               
                this.Invoke(new Action(()=> {

                    ucLoading1.Visible = false;
                    this.btn_start.Enabled = true;
                    this.btn_openPDF.Enabled = true;
                    if (savePathDir != "-1" && savePathDir != "-2")
                    {
                        Process.Start(saveDir);
                    }else
                    {
                        string message=(savePathDir == "-1"?"无法解析当前格式文件" :"文件被占用");
                        MessageBox.Show("图片转换失败，"+message, "提示");
                    }
                }));
            });
            th.Start();
        }


        /*
         * Annotations = 1,
        //
        // 摘要:
        //     Set if using text rendering optimized for LCD display.
        LcdText = 2,
        //
        // 摘要:
        //     Don't use the native text output available on some platforms.
        NoNativeText = 4,
        //
        // 摘要:
        //     Grayscale output.
        Grayscale = 8,
        //
        // 摘要:
        //     Limit image cache size.
        LimitImageCacheSize = 512,
        //
        // 摘要:
        //     Always use halftone for image stretching.
        ForceHalftone = 1024,
        //
        // 摘要:
        //     Render for printing.
        ForPrinting = 2048,
        //
        // 摘要:
        //     Render with a transparent background.
        Transparent = 4096,
        //
        // 摘要:
        //     Correct height/width for DPI.
        CorrectFromDpi = 8192
         */
        public PdfiumViewer.PdfRenderFlags retPicQuality(int index)
        {
            switch(index)
            {
                case 0:
                    return PdfiumViewer.PdfRenderFlags.Annotations;
                case 1:
                    return PdfiumViewer.PdfRenderFlags.ForPrinting;
                case 2:
                    return PdfiumViewer.PdfRenderFlags.CorrectFromDpi;
                    //case 0:
                    //    return PdfiumViewer.PdfRenderFlags.Annotations;
                    //case 1:
                    //      return PdfiumViewer.PdfRenderFlags.LcdText; 
                    //case 2:
                    //      return PdfiumViewer.PdfRenderFlags.NoNativeText; 
                    //case 3:
                    //      return PdfiumViewer.PdfRenderFlags.Grayscale; 
                    //case 4:
                    //      return PdfiumViewer.PdfRenderFlags.LimitImageCacheSize; 
                    //case 5:
                    //      return PdfiumViewer.PdfRenderFlags.ForceHalftone; 
                    //case 6:
                    //      return PdfiumViewer.PdfRenderFlags.ForPrinting; 
                    //case 7:
                    //      return PdfiumViewer.PdfRenderFlags.Transparent; 
                    //case 8:
                    //      return PdfiumViewer.PdfRenderFlags.CorrectFromDpi;
            }
            return PdfiumViewer.PdfRenderFlags.CorrectFromDpi; 
        }

        private void FrmPDFToImgs_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                if(this.th!=null && this.th.ThreadState!=ThreadState.Aborted)
                {
                    this.th.Abort();
                }
            }
            catch (ThreadAbortException ex)
            {

             
            }
            
        }
    }
}
