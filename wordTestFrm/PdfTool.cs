
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using PdfiumViewer;

namespace wordTestFrm
{
    public class PdfiumViewerTool
    {
        /// <summary>
        /// Pdf转图片
        /// </summary>
        /// <param name="PdfPath">PDF路径</param>
        /// <param name="fileName">文件名称</param>
        /// <param name="flag">压缩百分比</param>
        /// <param name="dpi">dpi</param>
        /// <returns> 
        /// -1 文件格式异常
        /// -2 文件被占用
        /// </returns>
        public string ConvertPDF2Pic(string PdfPath,string fileName, PdfRenderFlags pdfRenderFlags,int flag=100, int dpi=300)
        {
            try
            {
                var document = PdfiumViewer.PdfDocument.Load(PdfPath);
                int pdfPage = document.PageCount;
                IList<SizeF> itemsSize = document.PageSizes;

                string saveDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Pdf2Pic");
                if (!Directory.Exists(saveDir))
                {
                    Directory.CreateDirectory(saveDir);
                }
                saveDir = Path.Combine(saveDir, Path.GetFileNameWithoutExtension(fileName));
                if (!Directory.Exists(saveDir))
                {
                    Directory.CreateDirectory(saveDir);
                }

                for (int i = 1; i <= pdfPage; i++)
                {
                    Size size = new Size();
                    size.Height = (int)itemsSize[(i - 1)].Height;
                    size.Width = (int)itemsSize[(i - 1)].Width;
               
                    this.RenderPage(PdfPath, i, size, Path.Combine(saveDir, @"sample" + i + @".jpg"), pdfRenderFlags, flag, dpi);
                }
                return saveDir;
            }
            catch(PdfException ex)
            {
                return "-1";
            }
            catch (IOException ex)
            {

                return "-2";
            }
            
        }


        /// <summary>
        /// 将PDF转换为图片
        /// </summary>
        /// <param name="pdfPath">pdf文件位置</param>
        /// <param name="pageNumber">pdf文件张数</param>
        /// <param name="size">pdf文件尺寸</param>
        /// <param name="outputPath">输出图片位置与名称</param>
        public void RenderPage(string pdfPath, int pageNumber, System.Drawing.Size size, string outputPath, PdfRenderFlags pdfRenderFlags= PdfRenderFlags.CorrectFromDpi, int flag=100, int dpi = 300)
        {
            using (var document = PdfiumViewer.PdfDocument.Load(pdfPath))
            {
                using (var stream = new FileStream(outputPath, FileMode.Create))
                {
                    using (var image = GetPageImage(pageNumber, size, document, dpi, pdfRenderFlags))
                    {

                        if (flag != 100)
                        {
                            MemoryStream ms = (MemoryStream)Resampler.GetPicThumbnail(image, flag);
                            byte[] raw = ms.ToArray();
                            stream.Write(raw, 0, raw.Length);
                        }
                        else
                        {
                            image.Save(stream, ImageFormat.Jpeg);
                        }
                        
                    }
                }
            }
        }

        private static Image GetPageImage(int pageNumber, Size size, PdfiumViewer.PdfDocument document, int dpi, PdfRenderFlags pdfRenderFlags)
        {
            return document.Render(pageNumber - 1, size.Width, size.Height, dpi, dpi, pdfRenderFlags);
        }
        private static Image GetPageImage(int pageNumber, Size size, PdfiumViewer.PdfDocument document, int dpi)
        {
            return document.Render(pageNumber - 1, size.Width, size.Height, dpi, dpi, PdfRenderFlags.CorrectFromDpi);
        }

    }
}
