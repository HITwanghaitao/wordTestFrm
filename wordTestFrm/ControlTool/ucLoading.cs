using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;

namespace wordTestFrm.ControlTool
{
    public partial class ucLoading : UserControl
    {
        Random random = new Random();
        private const int Alpha = 111;
        public int Percentage = 0;
        Pen srcPen;
        SolidBrush srcBrush;
        public ucLoading()
        {
            this.SetStyle(System.Windows.Forms.ControlStyles.Opaque, true);
            InitializeComponent();
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                //开启 WS_EX_TRANSPARENT, 使控件支持透明
                cp.ExStyle |= 0x00000020;
                return cp;
            }
        }


        protected override void OnPaint(PaintEventArgs e)
        {
            float width;
            float height;

            Color srcColor = Color.FromArgb(Alpha, this.BackColor);

            srcPen = new Pen(srcColor, 0);

            srcBrush = new SolidBrush(srcColor);

            base.OnPaint(e);

            width = this.Size.Width;

            height = this.Size.Height;

            e.Graphics.DrawRectangle(srcPen, 0, 0, width, height);
            e.Graphics.FillRectangle(srcBrush, 0, 0, width, height);

            e.Graphics.DrawString(this.Text, this.Font, new SolidBrush(Color.Black), new Rectangle(0, 0, this.Width, this.Height), new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center });
            this.srcPen.Dispose();
            this.srcBrush.Dispose();
            base.OnPaint(e);
        }

        protected override void OnLoad(EventArgs e)
        {
            //this.Size = Program.loadImgItems[index].Size;
            this.Location = new Point((this.ParentForm.Width - this.Width) / 2,
                (this.ParentForm.Height - this.Height) / 2);
            base.OnLoad(e);
            timer1.Start();
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (Program.loadImgItems.Count == 0) return;
            this.Percentage += 8;
            int index= random.Next(0, Program.loadImgItems.Count-1);
            lblLoading.Image = Program.loadImgItems[index];
            this.Size= Program.loadImgItems[index].Size;
            this.Location = new Point((this.ParentForm.Width - this.Width) / 2,
                (this.ParentForm.Height - this.Height) / 2);
        }

        /// <summary>
        /// Gif 压缩
        /// </summary>
        /// <param name="image"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        /// <param name="imageName"></param>
        /// <returns></returns>
        public static Image reSizeForImg(Image image,int width,int height,string imageName)
        {
            string dirGif = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "gif");
            if(!Directory.Exists(dirGif))
            {
                Directory.CreateDirectory(dirGif);
            }
            string savePath = Path.Combine(dirGif, imageName+".gif");
            if(File.Exists(savePath))
            {
                return Image.FromFile(savePath);
            }
            //原图
            Image img = (Image)image.Clone();

            //按比例缩放             
            int sourWidth = img.Width;
            int sourHeight = img.Height;
            if (sourHeight > height || sourWidth > width)
            {
                if ((sourWidth * height) > (sourHeight * width))
                {
                    //width = destWidth;
                    height = (width * sourHeight) / sourWidth;
                }
                else
                {
                    //height = destHeight;
                    width = (sourWidth * height) / sourHeight;
                }
            }
            else
            {
                width = sourWidth;
                height = sourHeight;
            }

            //不够100*100的不缩放
            if (img.Width > width && img.Height > height)
            {
                //新图第一帧
                Image new_img = new Bitmap(width, height);
                //新图其他帧
                Image new_imgs = new Bitmap(width, height);
                //新图第一帧GDI+绘图对象
                Graphics g_new_img = Graphics.FromImage(new_img);
                //新图其他帧GDI+绘图对象
                Graphics g_new_imgs = Graphics.FromImage(new_imgs);
                //配置新图第一帧GDI+绘图对象
                g_new_img.CompositingMode = CompositingMode.SourceCopy;
                g_new_img.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g_new_img.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g_new_img.SmoothingMode = SmoothingMode.HighQuality;
                g_new_img.Clear(Color.FromKnownColor(KnownColor.Transparent));
                //配置其他帧GDI+绘图对象
                g_new_imgs.CompositingMode = CompositingMode.SourceCopy;
                g_new_imgs.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g_new_imgs.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g_new_imgs.SmoothingMode = SmoothingMode.HighQuality;
                g_new_imgs.Clear(Color.FromKnownColor(KnownColor.Transparent));
                //遍历维数
                foreach (Guid gid in img.FrameDimensionsList)
                {
                    //因为是缩小GIF文件所以这里要设置为Time
                    //如果是TIFF这里要设置为PAGE
                    FrameDimension f = FrameDimension.Time;
                    //获取总帧数
                    int count = img.GetFrameCount(f);
                    //保存标示参数
                    System.Drawing.Imaging.Encoder encoder = System.Drawing.Imaging.Encoder.SaveFlag;
                    //
                    EncoderParameters ep = null;
                    //图片编码、解码器
                    ImageCodecInfo ici = null;
                    //图片编码、解码器集合
                    ImageCodecInfo[] icis = ImageCodecInfo.GetImageDecoders();
                    //为 图片编码、解码器 对象 赋值
                    foreach (ImageCodecInfo ic in icis)
                    {
                        if (ic.FormatID == ImageFormat.Gif.Guid)
                        {
                            ici = ic;
                            break;
                        }
                    }
                    //每一帧
                    for (int c = 0; c < count; c++)
                    {
                        //选择由维度和索引指定的帧
                        img.SelectActiveFrame(f, c);
                        //第一帧
                        if (c == 0)
                        {
                            //将原图第一帧画给新图第一帧
                            g_new_img.DrawImage(img, new Rectangle(0, 0, width, height), new Rectangle(0, 0, img.Width, img.Height), GraphicsUnit.Pixel);
                            //把振频和透明背景调色板等设置复制给新图第一帧
                            for (int i = 0; i < img.PropertyItems.Length; i++)
                            {
                                 new_img.SetPropertyItem(img.PropertyItems[i]);
                            }
                            ep = new EncoderParameters(1);
                            //第一帧需要设置为MultiFrame
                            ep.Param[0] = new EncoderParameter( encoder,(long)EncoderValue.MultiFrame);
                            //保存第一帧
                            new_img.Save(savePath, ici, ep);
                            //new_img.Save(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"/temp/" + Path.GetFileName(imgPath), ici, ep);
                        }
                        //其他帧
                        else
                        {
                            //把原图的其他帧画给新图的其他帧
                            g_new_imgs.DrawImage(img, new Rectangle(0, 0, width, height), new Rectangle(0, 0, img.Width, img.Height), GraphicsUnit.Pixel);
                            //把振频和透明背景调色板等设置复制给新图第一帧
                            for (int i = 0; i < img.PropertyItems.Length; i++)
                            {
                                new_imgs.SetPropertyItem(img.PropertyItems[i]);
                            }
                            ep = new EncoderParameters(1);
                            //如果是GIF这里设置为FrameDimensionTime
                            //如果为TIFF则设置为FrameDimensionPage
                            ep.Param[0] = new EncoderParameter(encoder, (long)EncoderValue.FrameDimensionTime);
                            //向新图添加一帧
                            new_img.SaveAdd(new_imgs, ep);
                        }
                    }
                    ep = new EncoderParameters(1);
                    //关闭多帧文件流
                    ep.Param[0] = new EncoderParameter(encoder, (long)EncoderValue.Flush);
                    new_img.SaveAdd(ep);
                }

                //释放文件
                img.Dispose();
                new_img.Dispose();
                new_imgs.Dispose();
                g_new_img.Dispose();
                g_new_imgs.Dispose();
                return Image.FromFile(savePath);
            }
            return img;
        }

        private void lblProgress_Paint(object sender, PaintEventArgs e)
        {
            //43,145,175
            //221,83,71
            if (this.Percentage > 100) this.Percentage = 100;
            Color c = Color.FromArgb(143- this.Percentage, 45+this.Percentage,75+this.Percentage);
            if (this.Percentage <= 15) c = Color.FromArgb(221, 83, 71);
            Graphics graphics = e.Graphics;
            graphics.FillRectangle(new SolidBrush(c), new Rectangle(new Point(0, 0), new Size(lblProgress.Width * this.Percentage / 100, this.lblProgress.Height)));
        }
    }
}
