using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Diagnostics;
using System.Text.RegularExpressions;
using Aspose.Words.Saving;
using System.Data.SqlClient;
using Aspose.Words.Fields;
using Aspose.Words.Markup;
using wordTestFrm.Model;
using Aspose.Words.Tables;
using Aspose.Words.Layout;

namespace wordTestFrm
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            cbxDPI.SelectedIndex = 0;

        }
        // 220ppi Print - said to be excellent on most printers and screens.
        // 150ppi Screen - said to be good for web pages and projectors.
        // 96ppi Email - said to be good for minimal document size and sharing.
        public int desiredPpi = 150;

        // In .NET this seems to be a good compression / quality setting.
        public int jpegQuality = 90;

        /// <summary>
        /// word整体调整
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDocStyleChange_Click(object sender, EventArgs e)
        {
            this.desiredPpi = int.Parse(cbxDPI.SelectedItem.ToString());
            this.jpegQuality = int.Parse(txtQuality.Text);

            OpenFileDialog of = new OpenFileDialog();
            of.InitialDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            if (of.ShowDialog() == DialogResult.OK)
            {
                string realName = Path.GetFileName(of.FileName);
                string CopyPath = Common.CommonTool.retSavePath(of.FileName);
                Document doc = new Aspose.Words.Document(CopyPath);

                string fileName = of.FileName;
                FileInfo f = new FileInfo(fileName);

                List<string> numberlist = new List<string>();
                DocumentBuilder build = new DocumentBuilder(doc);
                int sectionCount = doc.Sections.Count;
                string content = string.Empty;

                NodeCollection nodes_Paragraph = doc.GetChildNodes(NodeType.Paragraph, true);

                //修改段落字体
                this.SetStyleForParagraph(doc, ref content);

                //添加页眉页脚
                AddHeaderFooter(doc);

                //修改 页眉/页脚
                SetStyleForHeaderFooter(doc);

                //标题添加编号
                AddNumForHeader(doc);

                //修改符号/编号
                this.SetStyleForListFlags(doc);

                //修改表格
                System.Drawing.Font fTable = (System.Drawing.Font)lblFont.Tag;
                Color colorFont = (Color)lblFontColor.Tag;
                Color colorBorder = (Color)lblBorderColor.Tag;
                this.SetStyleForTable(doc, fTable, colorBorder, colorFont, ref content);


                //修改图片
                this.desiredPpi = int.Parse(cbxDPI.SelectedItem.ToString());
                this.jpegQuality = int.Parse(txtQuality.Text);
                this.SetStyleForImage(doc, this.desiredPpi, this.jpegQuality);

                #region 获取所有章节及文字内容
                /*
                            if (p.ParagraphFormat.OutlineLevel.ToString() == "BodyText")
                            {
                                string text = p.GetText();
                                string pattern = "(.*)(\r)$";
                                Match match = Regex.Match(p.GetText(), pattern);
                                if (match.Success)
                                {
                                    int count1 = match.Groups.Count;
                                    string textStr = match.Groups[1].Value;


                                    if (!textStr.Equals("正文"))
                                    {
                                        this.SetStyleForParagraph(p, "微软雅黑", 12, Color.Red, true, true);

                                        //符号/编号修改
                                        if (p.IsListItem)
                                        {
                                            p.ListFormat.ListLevel.Font.Size = 20;
                                            p.ListFormat.ListLevel.Font.Color = Color.Blue;
                                           if(!numberlist.Contains(p.ListFormat.ListLevel.NumberStyle.ToString()))
                                           {
                                               numberlist.Add(p.ListFormat.ListLevel.NumberStyle.ToString());
                                           }
                                        }

                                    }
                                }


                            }
                            else
                            {
                                p.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                p.ParagraphFormat.Style.Font.Size = 12;
                                p.ParagraphFormat.Style.Font.Bold = false;
                                p.ParagraphFormat.Style.Font.Name = "微软雅黑";
                                p.ParagraphFormat.Style.Font.Color = Color.Green;
                            }
                            content += p.ParagraphFormat.OutlineLevel + " " + p.ParagraphFormat.StyleName + ":" + p.GetText() + "\r\n";
                        */
                #endregion

                string savePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Path.GetFileNameWithoutExtension(realName) + "_AsposeWord.docx");
                doc.Save(savePath, SaveFormat.Docx);
                Process.Start(savePath);
                textBox1.Text = content;

                return;
            }
        }


        #region 样式设置

        /// <summary>
        /// 设置边框
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="lineStyle"></param>
        /// <param name="lineWidth"></param>
        /// <param name="color"></param>
        /// <returns></returns>
        public bool setStyleForBorder(Aspose.Words.Border border, LineStyle lineStyle, double lineWidth, Color color)
        {
            border.Color = color;
            border.LineStyle = lineStyle;
            border.LineWidth = lineWidth;
            return true;
        }


        /// <summary>
        /// 设置段落字体
        /// </summary>
        /// <param name="p">段落</param>
        /// <param name="fontName">字体</param>
        /// <param name="size"></param>
        /// <param name="fontColor"></param>
        /// <param name="isBold"></param>
        /// <param name="Italic"></param>
        /// <returns></returns>
        public bool SetStyleForParagraphFont(Paragraph p, string fontName, float size, Color fontColor, bool isBold = false, bool Italic = false)
        {
            try
            {

                foreach (Run item in p.Runs)
                {
                    if (item == null) continue;
                    item.Font.Size = size;
                    item.Font.Color = fontColor;
                    item.Font.Bold = isBold;
                    item.Font.Italic = Italic;
                    item.Font.Name = fontName;
                    item.ParentParagraph.ParagraphFormat.LineSpacing = 12;
                    item.ParentParagraph.ParagraphFormat.SpaceAfter = 1;
                    item.ParentParagraph.ParagraphFormat.SpaceBefore = 1;
                    item.ParentParagraph.ParagraphFormat.LeftIndent = 100;
                    item.ParentParagraph.ParagraphFormat.RightIndent = 100;
                }
                return true;
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return false;
            }
        }

        /// <summary>
        /// 设置页眉页脚样式
        /// </summary>
        /// <param name="hf"></param>
        /// <param name="fontName"></param>
        /// <param name="size"></param>
        /// <param name="fontColor"></param>
        /// <param name="isBold"></param>
        /// <param name="Italic"></param>
        /// <returns></returns>
        public bool SetStyleForHeaderFooterFont(HeaderFooter hf, string fontName, float size, Color fontColor, bool isBold = false, bool Italic = false)
        {
            try
            {
                foreach (Run item in hf.FirstParagraph.Runs)
                {
                    string txt = item.GetText();
                    if (item == null) continue;
                    item.Font.Size = size;
                    item.Font.Color = fontColor;
                    item.Font.Bold = isBold;
                    item.Font.Italic = Italic;
                    item.Font.Name = fontName;
                    //if (hf.HeaderFooterType == HeaderFooterType.FooterPrimary)
                    //{
                    //    hf.FirstParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    //}
                    //else if (hf.HeaderFooterType == HeaderFooterType.FooterEven)
                    //{
                    //    hf.FirstParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                    //}
                    //else
                    //{
                    //    hf.FirstParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                    //}
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }


        /// <summary>
        /// 设置编号/项目符号样式
        /// </summary>
        /// <param name="list">编号/符号</param>
        /// <param name="fontName">字体</param>
        /// <param name="size"></param>
        /// <param name="fontColor"></param>
        /// <param name="isBold"></param>
        /// <param name="Italic"></param>
        /// <param name="isFlags">是否项目符号</param>
        /// <returns></returns>
        public bool SetStyleForListFormatFont(ListFormat list, string fontName, float size, Color fontColor, bool isBold = false, bool Italic = false, bool isFlags = false)
        {
            try
            {
                list.ListLevel.Font.Size = size;
                list.ListLevel.Font.Color = fontColor;
                if (!isFlags)
                {
                    list.ListLevel.Font.Bold = isBold;
                    list.ListLevel.Font.Italic = Italic;
                    list.ListLevel.Font.Name = fontName;
                }
                return true;
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return false;
            }
        }
        #endregion

        /// <summary>
        /// 图片等比缩放
        /// </summary>
        /// <param name="mg"></param>
        /// <param name="newSize"></param>
        /// <returns></returns>
        public Bitmap GetImageThumb(Bitmap mg, Size newSize)
        {
            double ratio = 0d;
            double myThumbWidth = 0d;
            double myThumbHeight = 0d;
            int x = 0;
            int y = 0;

            Bitmap bp;

            if ((mg.Width / Convert.ToDouble(newSize.Width)) > (mg.Height /
            Convert.ToDouble(newSize.Height)))
                ratio = Convert.ToDouble(mg.Width) / Convert.ToDouble(newSize.Width);
            else
                ratio = Convert.ToDouble(mg.Height) / Convert.ToDouble(newSize.Height);
            myThumbHeight = Math.Ceiling(mg.Height / ratio);
            myThumbWidth = Math.Ceiling(mg.Width / ratio);

            Size thumbSize = new Size((int)newSize.Width, (int)newSize.Height);
            bp = new Bitmap(newSize.Width, newSize.Height);
            x = (newSize.Width - thumbSize.Width) / 2;
            y = (newSize.Height - thumbSize.Height);
            System.Drawing.Graphics g = Graphics.FromImage(bp);
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            Rectangle rect = new Rectangle(x, y, thumbSize.Width, thumbSize.Height);
            g.DrawImage(mg, rect, 0, 0, mg.Width, mg.Height, GraphicsUnit.Pixel);

            return bp;
        }

        /// <summary>
        /// 图片按照长宽 缩放
        /// </summary>
        /// <param name="mg"></param>
        /// <param name="newSize"></param>
        /// <returns></returns>
        public Bitmap GetImageThumb_2(Bitmap mg, Size newSize)
        {
            double ratio = 0d;
            //double myThumbWidth = 0d;
            //double myThumbHeight = 0d;
            int x = 0;
            int y = 0;

            Bitmap bp;

            //if ((mg.Width / Convert.ToDouble(newSize.Width)) > (mg.Height /
            //Convert.ToDouble(newSize.Height)))
            //    ratio = Convert.ToDouble(mg.Width) / Convert.ToDouble(newSize.Width);
            //else
            //    ratio = Convert.ToDouble(mg.Height) / Convert.ToDouble(newSize.Height);
            //myThumbHeight = Math.Ceiling(mg.Height / ratio);
            //myThumbWidth = Math.Ceiling(mg.Width / ratio);

            Size thumbSize = newSize;
            bp = new Bitmap(newSize.Width, newSize.Height);
            x = (newSize.Width - thumbSize.Width) / 2;
            y = (newSize.Height - thumbSize.Height);
            System.Drawing.Graphics g = Graphics.FromImage(bp);
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            System.Drawing.Rectangle rect = new Rectangle(x, y, thumbSize.Width, thumbSize.Height);
            g.DrawImage(mg, rect, 0, 0, mg.Width, mg.Height, GraphicsUnit.Pixel);

            return bp;
        }

        /// <summary>
        /// 插入目录
        /// </summary>
        /// <param name="bulider_blank"></param>
        public static void InsertTOC(DocumentBuilder bulider_blank)
        {
            //设置"目录"格式
            bulider_blank.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            bulider_blank.Bold = true;
            bulider_blank.Font.Name = "SONG";
            bulider_blank.Writeln("目录");
            bulider_blank.ParagraphFormat.ClearFormatting();//清除所有样式
            bulider_blank.InsertTableOfContents("\\o\"1-3\"\\h\\z\\u");
            bulider_blank.InsertBreak(BreakType.SectionBreakNewPage);
        }

        /// <summary>
        /// 图片压缩
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPicCompress_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.InitialDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            if (of.ShowDialog() == DialogResult.OK)
            {
                string realName = Path.GetFileName(of.FileName);
                string CopyPath = Common.CommonTool.retSavePath(of.FileName);
                Document doc = new Aspose.Words.Document(CopyPath);
                this.desiredPpi = int.Parse(cbxDPI.SelectedItem.ToString());
                this.jpegQuality = int.Parse(txtQuality.Text);

                Paragraph p = doc.Sections[0].Body.FirstParagraph;
                NodeCollection nodes = p.GetChildNodes(NodeType.Shape, true);
                int count = nodes.Count;
                string txt = p.GetText();
                
                SetStyleForImage(doc, this.desiredPpi, this.jpegQuality);

                string savePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Path.GetFileNameWithoutExtension(realName) + "_AsposeWord.docx");
                doc.Save(savePath, SaveFormat.Docx);
                Process.Start(savePath);


            }
        }


        /// <summary>
        /// 表格样式
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnTable_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.InitialDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            if (of.ShowDialog() == DialogResult.OK)
            {
                string content = string.Empty;
                string realName = Path.GetFileName(of.FileName);
                string CopyPath = Common.CommonTool.retSavePath(of.FileName);
                Document doc = new Aspose.Words.Document(CopyPath);
                System.Drawing.Font f = (System.Drawing.Font)lblFont.Tag;
                Color colorFont = (Color)lblFontColor.Tag;
                Color colorBorder = (Color)lblBorderColor.Tag;

                this.SetStyleForTable(doc, f, colorBorder, colorFont, ref content);

                string savePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Path.GetFileNameWithoutExtension(realName) + "_AsposeWord.docx");
                doc.Save(savePath, SaveFormat.Docx);
                Process.Start(savePath);
                textBox1.Text = content;

            }
        }

        private void lblColorChange_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                Label lbl = ((Label)sender);
                lbl.Tag = colorDialog1.Color;
                lbl.BackColor = colorDialog1.Color;
            }
        }

        private void lblFontChange_Click(object sender, EventArgs e)
        {
            if (fontDialog1.ShowDialog() == DialogResult.OK)
            {
                Label lbl = ((Label)sender);
                lbl.Tag = fontDialog1.Font;
                lbl.Text = string.Format("字体:{0}  字号: {1}", fontDialog1.Font.Name, fontDialog1.Font.Size);
            }
        }

        /// <summary>
        /// 设置段落字体
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnParagraph_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.InitialDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            if (of.ShowDialog() == DialogResult.OK)
            {
                string content = string.Empty;
                string realName = Path.GetFileName(of.FileName);
                string CopyPath = Common.CommonTool.retSavePath(of.FileName);

                Document doc = new Document(CopyPath);
                this.SetStyleForParagraph(doc, ref content);

                string savePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Path.GetFileNameWithoutExtension(realName) + "_AsposeWord.docx");
                doc.Save(savePath, SaveFormat.Docx);
                Process.Start(savePath);
                textBox1.Text = content;
            }
        }

        #region 样式修改
        #region 样式设置

        /// <summary>
        /// 设置边框
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="lineStyle"></param>
        /// <param name="lineWidth"></param>
        /// <param name="color"></param>
        /// <returns></returns>
        public bool SetStyleForBorder(Aspose.Words.Border border, LineStyle lineStyle, double lineWidth, Color color)
        {
            border.Color = color;
            border.LineStyle = lineStyle;
            border.LineWidth = lineWidth;
            return true;
        }


        ///// <summary>
        ///// 设置段落字体
        ///// </summary>
        ///// <param name="p">段落</param>
        ///// <param name="fontName">字体</param>
        ///// <param name="size"></param>
        ///// <param name="fontColor"></param>
        ///// <param name="isBold"></param>
        ///// <param name="Italic"></param>
        ///// <returns></returns>
        //public bool SetStyleForParagraphFont(Paragraph p, string fontName, float size, Color fontColor, bool isBold = false, bool Italic = false)
        //{
        //    try
        //    {

        //        foreach (Run item in p.Runs)
        //        {
        //            if (item == null) continue;
        //            item.Font.Size = size;
        //            item.Font.Color = fontColor;
        //            item.Font.Bold = isBold;
        //            item.Font.Italic = Italic;
        //            item.Font.Name = fontName;
        //        }
        //        return true;
        //    }
        //    catch (Exception ex)
        //    {

        //        MessageBox.Show(ex.Message);
        //        return false;
        //    }
        //}


        ///// <summary>
        ///// 设置编号/项目符号样式
        ///// </summary>
        ///// <param name="list">编号/符号</param>
        ///// <param name="fontName">字体</param>
        ///// <param name="size"></param>
        ///// <param name="fontColor"></param>
        ///// <param name="isBold"></param>
        ///// <param name="Italic"></param>
        ///// <param name="isFlags">是否项目符号</param>
        ///// <returns></returns>
        //public bool SetStyleForListFormatFont(ListFormat list, string fontName, float size, Color fontColor, bool isBold = false, bool Italic = false, bool isFlags = false)
        //{
        //    try
        //    {

        //        list.ListLevel.Font.Size = size;
        //        list.ListLevel.Font.Color = fontColor;
        //        if (!isFlags)
        //        {
        //            list.ListLevel.Font.Bold = isBold;
        //            list.ListLevel.Font.Italic = Italic;
        //            list.ListLevel.Font.Name = fontName;
        //        }
        //        return true;
        //    }
        //    catch (Exception ex)
        //    {

        //        MessageBox.Show(ex.Message);
        //        return false;
        //    }
        //}
        #endregion

        /// <summary>
        /// 标记符号调整
        /// </summary>
        /// <param name="doc"></param>
        public void SetStyleForListFlags(Document doc)
        {
            NodeCollection nodes = doc.GetChildNodes(NodeType.Paragraph, true);
            ////编号设置
            //Aspose.Words.Lists.List list = this.retList(doc);

            for (int i = 0; i < nodes.Count; i++)
            {
                Paragraph p = (Paragraph)nodes[i];
                System.Drawing.Font font = (System.Drawing.Font)lblLv1FlagFont.Tag;
                System.Drawing.Color color = Color.FromArgb(0, 0, 0, 0);
                bool isFlags = false;//是否项目符号

                if (p.IsListItem)
                {
                    switch (p.ParagraphFormat.OutlineLevel)
                    {
                        case OutlineLevel.Level1:
                            font = (System.Drawing.Font)lblLv1FlagFont.Tag;
                            color = (System.Drawing.Color)lblLV1FlagColor.Tag;
                            break;
                        case OutlineLevel.Level2:
                            font = (System.Drawing.Font)lblLv2FlagFont.Tag;
                            color = (System.Drawing.Color)lblLV2FlagColor.Tag;
                            break;
                        case OutlineLevel.Level3:
                            font = (System.Drawing.Font)lblLv3FlagFont.Tag;
                            color = (System.Drawing.Color)lblLV3FlagColor.Tag;
                            break;
                        case OutlineLevel.Level4:
                            font = (System.Drawing.Font)lblLv4FlagFont.Tag;
                            color = (System.Drawing.Color)lblLV4FlagColor.Tag;
                            break;
                        case OutlineLevel.Level5:
                            font = (System.Drawing.Font)lblLv5FlagFont.Tag;
                            color = (System.Drawing.Color)lblLV5FlagColor.Tag;
                            break;
                        case OutlineLevel.BodyText:
                            font = (System.Drawing.Font)lblLvStrFlagFont.Tag;
                            color = (System.Drawing.Color)lblLVStrFlagColor.Tag;
                            break;
                    }
                    if (p.ListFormat.ListLevel.NumberStyle == NumberStyle.Bullet)
                    {
                        isFlags = true;
                    }
                    this.SetStyleForListFormatFont(p.ListFormat, font.Name, font.Size, color, font.Bold, font.Italic, isFlags);
                }
            }
        }



        /// <summary>
        /// 设置页眉页脚样式
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void SetStyleForHeaderFooter(Document doc)
        {
            doc.Sections[0].PageSetup.DifferentFirstPageHeaderFooter = true;
            doc.Sections[0].PageSetup.OddAndEvenPagesHeaderFooter = true;
            NodeCollection headfooters = doc.GetChildNodes(NodeType.HeaderFooter, true);

            for (int i = 0; i < headfooters.Count; i++)
            {
                HeaderFooter hf = (HeaderFooter)headfooters[i];
                Color cFont = Color.FromArgb(0, 0, 0, 0);
                System.Drawing.Font f = (System.Drawing.Font)lblhfFont.Tag;
                switch (hf.HeaderFooterType)
                {
                    case HeaderFooterType.HeaderFirst:
                    case HeaderFooterType.HeaderEven:
                    case HeaderFooterType.HeaderPrimary:
                        f = (System.Drawing.Font)lblhfFont.Tag;
                        //if (!IsNullOfFont(f)) return false;
                        cFont = (Color)lblhfColor.Tag;
                        //if (!IsNullOfCfont(cFont)) return false;
                        //hf.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                        hf.FirstParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                        //hf.FirstParagraph.ParagraphFormat.SpaceBefore = 30;//段间距
                        //hf.FirstParagraph.ParagraphFormat.LineSpacing = 50;//行间距
                        hf.FirstParagraph.ParagraphFormat.LeftIndent = 40;//缩进
                        //插入图标
                        //hf.FirstParagraph.InsertImage("图片绝对地址", 80, 80);//可以控制图片的宽高
                        this.SetStyleForHeaderFooterFont(hf, f.Name, f.Size, cFont, f.Bold, f.Italic);
                        break;
                    case HeaderFooterType.FooterFirst:
                        f = (System.Drawing.Font)lblfootFont.Tag;
                        cFont = (Color)lblfootColor.Tag;
                        //hf.ParagraphFormat.Alignment = ParagraphAlignment.Left; 
                        hf.FirstParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                        this.SetStyleForHeaderFooterFont(hf, f.Name, f.Size, cFont, f.Bold, f.Italic);
                        hf.FirstParagraph.InsertField("PAGE", null, true);
                        break;
                    case HeaderFooterType.FooterEven:
                        f = (System.Drawing.Font)lblfootFont.Tag;
                        cFont = (Color)lblfootColor.Tag;
                        //hf.ParagraphFormat.Alignment = ParagraphAlignment.Left;

                        hf.FirstParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Left;

                        Node node_item = hf.FirstParagraph.FirstChild;

                        Run run1 = new Run(doc, "~");
                        Run run2 = new Run(doc, "~");
                        hf.FirstParagraph.InsertBefore(run1, node_item);
                        hf.FirstParagraph.InsertBefore(run2, node_item);

                        Field field = hf.FirstParagraph.InsertField("PAGE", run1, true);
                        this.SetStyleForHeaderFooterFont(hf, f.Name, f.Size, cFont, f.Bold, f.Italic);
                        break;
                    case HeaderFooterType.FooterPrimary:
                        f = (System.Drawing.Font)lblfootFont.Tag;
                        cFont = (Color)lblfootColor.Tag;
                        //hf.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                        hf.FirstParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                        this.SetStyleForHeaderFooterFont(hf, f.Name, f.Size, cFont, f.Bold, f.Italic);
                        break;
                }
            }
        }

        /// <summary>
        /// 设置段落样式
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="content"></param>
        public void SetStyleForParagraph(Document doc, ref string content)
        {
            NodeCollection nodes = doc.GetChildNodes(NodeType.Paragraph, true);

            for (int i = 0; i < nodes.Count; i++)
            {
                System.Drawing.Font f = (System.Drawing.Font)lblLv1Font.Tag;
                Color cFont = Color.FromArgb(0, 0, 0, 0);
                Paragraph p = (Paragraph)nodes[i];
                switch (p.ParagraphFormat.OutlineLevel)
                {
                    case OutlineLevel.Level1:
                        f = (System.Drawing.Font)lblLv1Font.Tag;
                        //字体为空 提示设置字体
                        //if (!IsNullOfFont(f)) return false;
                        cFont = (Color)lblLv1Color.Tag;
                        //提示设置颜色
                        //if (!IsNullOfCfont(cFont)) return false;
                        //设置 居中/居左/居右
                        if (cbxhead1.SelectedItem == null)
                        {
                            SetParagraphAlignment("", p);//默认居左
                        }
                        else
                            SetParagraphAlignment(cbxhead1.SelectedItem.ToString(), p);
                        this.SetStyleForParagraphFont(p, f.Name, f.Size, cFont, f.Bold, f.Italic);
                        break;
                    case OutlineLevel.Level2:
                        f = (System.Drawing.Font)lblLv2Font.Tag;
                        cFont = (Color)lblLv2Color.Tag;
                        if (cbxhead2.SelectedItem == null)
                        {
                            SetParagraphAlignment("", p);
                        }
                        else
                            SetParagraphAlignment(cbxhead2.SelectedItem.ToString(), p);
                        this.SetStyleForParagraphFont(p, f.Name, f.Size, cFont, f.Bold, f.Italic);
                        break;
                    case OutlineLevel.Level3:
                        f = (System.Drawing.Font)lblLv3Font.Tag;
                        cFont = (Color)lblLv3Color.Tag;
                        if (cbxhead3.SelectedItem == null)
                        {
                            SetParagraphAlignment("", p);
                        }
                        else
                            SetParagraphAlignment(cbxhead3.SelectedItem.ToString(), p);
                        this.SetStyleForParagraphFont(p, f.Name, f.Size, cFont, f.Bold, f.Italic);
                        break;
                    case OutlineLevel.Level4:
                        f = (System.Drawing.Font)lblLv4Font.Tag;
                        cFont = (Color)lblLv4Color.Tag;
                        if (cbxhead4.SelectedItem == null)
                        {
                            SetParagraphAlignment("", p);
                        }
                        else
                            SetParagraphAlignment(cbxhead4.SelectedItem.ToString(), p);
                        this.SetStyleForParagraphFont(p, f.Name, f.Size, cFont, f.Bold, f.Italic);
                        break;
                    case OutlineLevel.Level5:
                        f = (System.Drawing.Font)lblLv5Font.Tag;
                        cFont = (Color)lblLv5Color.Tag;
                        if (cbxhead5.SelectedItem == null)
                        {
                            SetParagraphAlignment("", p);
                        }
                        else
                            SetParagraphAlignment(cbxhead5.SelectedItem.ToString(), p);
                        this.SetStyleForParagraphFont(p, f.Name, f.Size, cFont, f.Bold, f.Italic);
                        break;
                    case OutlineLevel.BodyText:
                        string ss = p.ParagraphFormat.StyleName;
                        f = (System.Drawing.Font)lblLvStrFont.Tag;
                        cFont = (Color)lblLvStrColor.Tag;

                        if (cbxbody.SelectedItem == null)
                        {
                            SetParagraphAlignment("", p);
                        }
                        else
                        {
                            SetParagraphAlignment(cbxbody.SelectedItem.ToString(), p);
                        }

                        this.SetStyleForParagraphFont(p, f.Name, f.Size, cFont, f.Bold, f.Italic);
                        break;
                    default:
                        f = (System.Drawing.Font)lblLvStrFont.Tag;
                        cFont = (Color)lblLvStrColor.Tag;
                        p.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
                        this.SetStyleForParagraphFont(p, f.Name, f.Size, cFont, f.Bold, f.Italic);
                        break;
                }
                #region 第一版
                //if (p.ParagraphFormat.IsHeading)//标题
                //{
                //    if (p.ParagraphFormat.OutlineLevel.ToString() == "Level1")
                //    {
                //        //p.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                //        f = (System.Drawing.Font)lblLv1Font.Tag;
                //        cFont = (Color)lblLv1Color.Tag;
                //        p.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                //        this.SetStyleForParagraphFont(p, f.Name, f.Size, cFont, f.Bold, f.Italic);
                //    }
                //    else if (p.ParagraphFormat.OutlineLevel.ToString() == "Level2")
                //    {
                //        f = (System.Drawing.Font)lblLv2Font.Tag;
                //        cFont = (Color)lblLv2Color.Tag;
                //        p.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                //        this.SetStyleForParagraphFont(p, f.Name, f.Size, cFont, f.Bold, f.Italic);
                //    }
                //    else if (p.ParagraphFormat.OutlineLevel.ToString() == "Level3")
                //    {
                //        f = (System.Drawing.Font)lblLv3Font.Tag;
                //        cFont = (Color)lblLv3Color.Tag;
                //        p.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                //        this.SetStyleForParagraphFont(p, f.Name, f.Size, cFont, f.Bold, f.Italic);
                //    }
                //    else if (p.ParagraphFormat.OutlineLevel.ToString() == "Level4")
                //    {
                //        f = (System.Drawing.Font)lblLv4Font.Tag;
                //        cFont = (Color)lblLv4Color.Tag;
                //        p.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                //        this.SetStyleForParagraphFont(p, f.Name, f.Size, cFont, f.Bold, f.Italic);
                //    }
                //    else if (p.ParagraphFormat.OutlineLevel.ToString() == "Level5")
                //    {
                //        f = (System.Drawing.Font)lblLv5Font.Tag;
                //        cFont = (Color)lblLv5Color.Tag;
                //        p.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                //        this.SetStyleForParagraphFont(p, f.Name, f.Size, cFont, f.Bold, f.Italic);
                //    }
                //}
                //else if (p.ParagraphFormat.OutlineLevel.ToString() == "BodyText")
                //{
                //    f = (System.Drawing.Font)lblLvStrFont.Tag;
                //    cFont = (Color)lblLvStrColor.Tag;
                //    p.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
                //    this.SetStyleForParagraphFont(p, f.Name, f.Size, cFont, f.Bold, f.Italic);
                //} 
                #endregion
            }
        }

        /// <summary>
        /// 设置段落居中/居左/居右
        /// </summary>
        /// <param name="cbxTxt"></param>
        /// <param name="p"></param>
        public void SetParagraphAlignment(string cbxTxt, Paragraph p)
        {
            switch (cbxTxt)
            {
                case "居左":
                    p.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    break;
                case "居中":
                    p.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                    break;
                case "居右":
                    p.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                    break;
                default:
                    p.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    break;
            }
        }

        /// <summary>
        /// 图片压缩
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="ppi"></param>
        /// <param name="Quality"></param>
        public void SetStyleForImage(Document doc, int ppi, int Quality)
        {
            //获取所有图片
            NodeCollection nodes_Pic = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            for (int i = 0; i < nodes_Pic.Count; i++)
            {
                Shape shape = (Shape)nodes_Pic[i];
                if (shape.HasImage)
                {
                    string time = DateTime.Now.ToString("HHmmssfff");
                    //扩展名
                    string ex = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    //文件名
                    string imgName = string.Format("{0}_{1}{2}", time, imageIndex, ex);
                    Image img = shape.ImageData.ToImage();
                    double width = shape.Width;
                    double height = shape.Height;
                    shape.HorizontalAlignment = Aspose.Words.Drawing.HorizontalAlignment.Center;
                    Image newImage = Resampler.ResampleCoreToImage(shape.ImageData, shape.SizeInPoints, ppi, Quality);
                    if (newImage == null) continue;
                    shape.ImageData.SetImage(newImage);

                    imageIndex++;
                }
            }
        }

        /// <summary>
        /// 表格样式修改
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="f"></param>
        /// <param name="colorBorder"></param>
        /// <param name="colorFont"></param>
        /// <param name="content"></param>
        public void SetStyleForTable(Document doc, System.Drawing.Font f, Color colorBorder, Color colorFont, ref string content)
        {
            NodeCollection nodes = doc.GetChildNodes(NodeType.Table, true);
            if (nodes != null && nodes.Count > 0)
            {
                for (int i = 0; i < nodes.Count; i++)
                {
                    Aspose.Words.Tables.Table table = (Aspose.Words.Tables.Table)nodes[i];
                    //table.Style = styleTable;
                    //table.f
                    for (int a = 0; a < table.Rows.Count; a++)
                    {
                        Aspose.Words.Tables.Row row = table.Rows[a];

                        for (int h = 0; h < row.Cells.Count; h++)
                        {
                            Aspose.Words.Tables.Cell cell = row.Cells[h];

                            //cell.CellFormat.Shading.BackgroundPatternColor = Color.Red;//底纹颜色

                            this.setStyleForBorder(cell.CellFormat.Borders[BorderType.Top], LineStyle.Double, 2.0, colorBorder);
                            this.setStyleForBorder(cell.CellFormat.Borders[BorderType.Bottom], LineStyle.Double, 2.0, colorBorder);
                            this.setStyleForBorder(cell.CellFormat.Borders[BorderType.Left], LineStyle.Double, 2.0, colorBorder);
                            this.setStyleForBorder(cell.CellFormat.Borders[BorderType.Right], LineStyle.Double, 2.0, colorBorder);
                            NodeCollection nodes_Cell = cell.GetChildNodes(NodeType.Paragraph, true);
                            int count = nodes_Cell.Count;
                            foreach (Paragraph item in nodes_Cell)
                            {
                                this.SetStyleForParagraphFont(item, f.Name, f.Size, colorFont, f.Bold, f.Italic);
                            }
                            string txt = cell.GetText();
                            content += txt + "\r\n";
                        }
                    }
                }
            }
        }

        #endregion


        /// <summary>
        /// PDF转JPEG
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPDF2JPEG_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.InitialDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            if (of.ShowDialog() == DialogResult.OK)
            {
                string filePath = of.FileName;
                string realName = of.SafeFileName;


                string fileName = Path.GetFileNameWithoutExtension(of.SafeFileName);

                int dpi = int.Parse(cbxDPI.SelectedItem.ToString());
                int Quality = int.Parse(txtQuality.Text);
                PdfiumViewerTool pt = new PdfiumViewerTool();
                string savePathDir = pt.ConvertPDF2Pic(filePath, fileName, PdfiumViewer.PdfRenderFlags.CorrectFromDpi,  Quality, dpi);



                //string savePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Path.GetFileNameWithoutExtension(fileName) + "_AsposeWord.jpeg");
                //doc.Save(savePath);
                Process.Start(savePathDir);

            }
        }

        /// <summary>
        /// 编号/项目符号修改
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnListFlags_Click(object sender, EventArgs e)
        {

            OpenFileDialog of = new OpenFileDialog();
            of.InitialDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            if (of.ShowDialog() == DialogResult.OK)
            {

                string realName = Path.GetFileName(of.FileName);
                string CopyPath = Common.CommonTool.retSavePath(of.FileName);
                Document doc = new Aspose.Words.Document(CopyPath);
                this.SetStyleForListFlags(doc);

                string savePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Path.GetFileNameWithoutExtension(realName) + "_AsposeWord.docx");
                doc.Save(savePath, SaveFormat.Docx);
                Process.Start(savePath);
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Color color = Color.Black;
            System.Drawing.Font font_def = new System.Drawing.Font("微软雅黑", 14.25f);

            lblfootColor.Tag = lblhfColor.Tag = lblBorderColor.Tag = lblFontColor.Tag
          = lblLv1Color.Tag = lblLv2Color.Tag = lblLv3Color.Tag = lblLv4Color.Tag = lblLv5Color.Tag = lblLvStrColor.Tag
          = lblLV1FlagColor.Tag = lblLV2FlagColor.Tag = lblLV3FlagColor.Tag = lblLV4FlagColor.Tag = lblLV5FlagColor.Tag = lblLVStrFlagColor.Tag = color;

            lblfootFont.Tag = lblhfFont.Tag = lblFont.Tag
                = lblLv1Font.Tag = lblLv2Font.Tag = lblLv3Font.Tag = lblLv4Font.Tag = lblLv5Font.Tag = lblLvStrFont.Tag
                = lblLv1FlagFont.Tag = lblLv2FlagFont.Tag = lblLv3FlagFont.Tag = lblLv4FlagFont.Tag = lblLv5FlagFont.Tag = lblLvStrFlagFont.Tag
                = font_def;
            tb_header.Text = "页眉";
            tb_footer.Text = "页脚";
        }

        private void lblhfColor_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                Label lbl = ((Label)sender);
                lbl.Tag = colorDialog1.Color;
                lbl.BackColor = colorDialog1.Color;
            }
        }

        private void lblhfFont_Click(object sender, EventArgs e)
        {
            if (fontDialog1.ShowDialog() == DialogResult.OK)
            {
                Label lbl = ((Label)sender);
                lbl.Tag = fontDialog1.Font;
                lbl.Text = string.Format("字体:{0}  字号: {1}", fontDialog1.Font.Name, fontDialog1.Font.Size);
            }
        }

        private void lblfootColor_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                Label lbl = ((Label)sender);
                lbl.Tag = colorDialog1.Color;
                lbl.BackColor = colorDialog1.Color;
            }
        }

        private void lblfootFont_Click(object sender, EventArgs e)
        {
            if (fontDialog1.ShowDialog() == DialogResult.OK)
            {
                Label lbl = ((Label)sender);
                lbl.Tag = fontDialog1.Font;
                lbl.Text = string.Format("字体:{0}  字号: {1}", fontDialog1.Font.Name, fontDialog1.Font.Size);
            }
        }

        private void btnHeaderFooter_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.InitialDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            if (of.ShowDialog() == DialogResult.OK)
            {
                string content = string.Empty;
                string realName = Path.GetFileName(of.FileName);
                string CopyPath = Common.CommonTool.retSavePath(of.FileName);
                Document doc = new Aspose.Words.Document(CopyPath);
                int count = doc.GetChildNodes(NodeType.HeaderFooter, true).Count;
                for (int i = 0; i < doc.Sections.Count; i++)
                {
                    Section section = doc.Sections[i];
                    count = section.HeadersFooters.Count;
                    section.PageSetup.RestartPageNumbering = false;
                    section.PageSetup.DifferentFirstPageHeaderFooter = true;
                    section.PageSetup.OddAndEvenPagesHeaderFooter = true;
                    if (i == 0)
                    {
                        HeaderFooter headerFooter = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
                        HeaderFooter headerFooter2 = new HeaderFooter(doc, HeaderFooterType.FooterEven);
                        HeaderFooter headerFooter3 = new HeaderFooter(doc, HeaderFooterType.FooterFirst);
                        Paragraph p = new Paragraph(doc);
                        p.InsertField(FieldType.FieldPage, true,null, true);
                        headerFooter.Paragraphs.Add(p);

                        p = new Paragraph(doc);
                        p.InsertField(FieldType.FieldPage, true, null, true);
                        headerFooter2.Paragraphs.Add(p);

                        p = new Paragraph(doc);
                        p.InsertField(FieldType.FieldPage, true, null, true);
                        headerFooter3.Paragraphs.Add(p);

                        section.HeadersFooters.Add(headerFooter);
                        section.HeadersFooters.Add(headerFooter2);
                        section.HeadersFooters.Add(headerFooter3);
                    }
                    else if(i==1)
                    {
                        section.PageSetup.RestartPageNumbering = true;
                    }
                    // count = section.HeadersFooters.Count;
                    //bool rpn = section.PageSetup.RestartPageNumbering;

                }


                //this.SetStyleForHeaderFooter(doc);

                string savePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Path.GetFileNameWithoutExtension(realName) + "_AsposeWord.docx");
                doc.Save(savePath, SaveFormat.Docx);
                Process.Start(savePath);
                textBox1.Text = content;
            }
        }

        private void btnSetNum_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.InitialDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            if (of.ShowDialog() == DialogResult.OK)
            {
                string realName = Path.GetFileName(of.FileName);
                string CopyPath = Common.CommonTool.retSavePath(of.FileName);
                Document doc = new Aspose.Words.Document(CopyPath);


                string content = string.Empty;
                AddNumForHeader(doc);
                string savePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Path.GetFileNameWithoutExtension(realName) + "_AsposeWord.docx");
                doc.Save(savePath, SaveFormat.Docx);
                Process.Start(savePath);
                textBox1.Text = content;
            }
        }

        /// <summary>
        /// 标题添加编号
        /// </summary>
        /// <param name="doc"></param>
        public void AddNumForHeader(Document doc)
        {
            string content = string.Empty;
            NodeCollection nodes = doc.GetChildNodes(NodeType.Paragraph, true);
            Aspose.Words.Lists.List list = this.retList(doc);
            for (int i = 0; i < nodes.Count; i++)
            {
                Paragraph item = (Paragraph)nodes[i];
                string text = item.GetText();
                if (item.ParagraphFormat.OutlineLevel == OutlineLevel.BodyText) continue;
                item.ListFormat.List = list;
                if (item.ParagraphFormat.OutlineLevel == OutlineLevel.BodyText)
                {
                    //item.ListFormat.RemoveNumbers();
                }
                else if (item.ParagraphFormat.OutlineLevel == OutlineLevel.Level1)
                {
                    item.ListFormat.ListLevelNumber = 0;
                }
                else if (item.ParagraphFormat.OutlineLevel == OutlineLevel.Level2)
                {
                    item.ListFormat.ListLevelNumber = 1;
                    //item.ListFormat.ListIndent();
                }
                else if (item.ParagraphFormat.OutlineLevel == OutlineLevel.Level3)
                {
                    item.ListFormat.ListLevelNumber = 2;
                    //item.ListFormat.ListIndent();
                    //item.ListFormat.ListIndent();
                }
                else if (item.ParagraphFormat.OutlineLevel == OutlineLevel.Level4)
                {
                    item.ListFormat.ListLevelNumber = 3;

                    //item.ListFormat.ListIndent();
                    //item.ListFormat.ListIndent();
                    //item.ListFormat.ListIndent();
                }
                else if (item.ParagraphFormat.OutlineLevel == OutlineLevel.Level5)
                {
                    item.ListFormat.ListLevelNumber = 4;
                    //item.ListFormat.ListIndent();
                    //item.ListFormat.ListIndent();
                    //item.ListFormat.ListIndent();
                    //item.ListFormat.ListIndent();
                }

            }
        }


        public Aspose.Words.Lists.List retList(Document doc)
        {
            Aspose.Words.Lists.List list = doc.Lists.Add(Aspose.Words.Lists.ListTemplate.NumberDefault);
            System.Drawing.Font font = (System.Drawing.Font)lblLv1FlagFont.Tag;
            System.Drawing.Color color = (System.Drawing.Color)lblLV1FlagColor.Tag;
            //
            // Completely customize one list level
            Aspose.Words.Lists.ListLevel listLevel = list.ListLevels[0];
            SetFontForAddHeader(listLevel, font, color);

            listLevel.NumberStyle = NumberStyle.Arabic;
            listLevel.StartAt = 1;
            //listLevel.NumberFormat = "\x0000";
            listLevel.NumberFormat = "\x0";

            listLevel.NumberPosition = -36;
            //listLevel.TextPosition = 144;
            //listLevel.TabPosition = 144;

            // Customize another list level
            listLevel = list.ListLevels[1];
            font = (System.Drawing.Font)lblLv2FlagFont.Tag;
            color = (System.Drawing.Color)lblLV2FlagColor.Tag;
            SetFontForAddHeader(listLevel, font, color);
            //listLevel.Alignment = Aspose.Words.Lists.ListLevelAlignment.Right;
            //listLevel.NumberStyle = NumberStyle.Bullet;
            //listLevel.Font.Name = "Wingdings";
            listLevel.NumberStyle = NumberStyle.Arabic;
            //listLevel.Font.Color = Color.Blue;
            //listLevel.Font.Size = 24;
            //listLevel.NumberFormat = "\xf0af"; // A bullet that looks like a star
            //listLevel.RestartAfterLevel = 1;
            listLevel.StartAt = 1;
            listLevel.NumberFormat = "\x0.\x1";
            //listLevel.TrailingCharacter = Aspose.Words.Lists.ListTrailingCharacter.Space;
            //listLevel.NumberPosition = 144;

            listLevel = list.ListLevels[2];
            font = (System.Drawing.Font)lblLv3FlagFont.Tag;
            color = (System.Drawing.Color)lblLV3FlagColor.Tag;
            SetFontForAddHeader(listLevel, font, color);
            //listLevel.Alignment = Aspose.Words.Lists.ListLevelAlignment.Right;
            //listLevel.NumberStyle = NumberStyle.Bullet;
            //listLevel.Font.Name = "Wingdings";
            listLevel.NumberStyle = NumberStyle.Arabic;
            //listLevel.Font.Color = Color.Blue;
            //listLevel.Font.Size = 24;
            //listLevel.NumberFormat = "\xf0af"; // A bullet that looks like a star
            //listLevel.RestartAfterLevel = 1;
            listLevel.StartAt = 1;
            listLevel.NumberFormat = "\x0.\x1.\x2";



            listLevel = list.ListLevels[3];
            font = (System.Drawing.Font)lblLv4FlagFont.Tag;
            color = (System.Drawing.Color)lblLV4FlagColor.Tag;
            SetFontForAddHeader(listLevel, font, color);
            //listLevel.Alignment = Aspose.Words.Lists.ListLevelAlignment.Right;
            //listLevel.NumberStyle = NumberStyle.Bullet;
            //listLevel.Font.Name = "Wingdings";
            listLevel.NumberStyle = NumberStyle.Arabic;
            //listLevel.Font.Color = Color.Blue;
            //listLevel.Font.Size = 24;
            //listLevel.NumberFormat = "\xf0af"; // A bullet that looks like a star
            //listLevel.RestartAfterLevel = 1;
            listLevel.StartAt = 1;
            listLevel.NumberFormat = "\x0.\x1.\x2.\x3";


            listLevel = list.ListLevels[4];
            font = (System.Drawing.Font)lblLv5FlagFont.Tag;
            color = (System.Drawing.Color)lblLV5FlagColor.Tag;
            SetFontForAddHeader(listLevel, font, color);
            //listLevel.Alignment = Aspose.Words.Lists.ListLevelAlignment.Right;
            //listLevel.NumberStyle = NumberStyle.Bullet;
            //listLevel.Font.Name = "Wingdings";
            listLevel.NumberStyle = NumberStyle.Arabic;
            //listLevel.Font.Color = Color.Blue;
            //listLevel.Font.Size = 24;
            //listLevel.NumberFormat = "\xf0af"; // A bullet that looks like a star
            //listLevel.RestartAfterLevel = 1;
            listLevel.StartAt = 1;
            listLevel.NumberFormat = "\x0.\x1.\x2.\x3.\x4";

            //listLevel = list;
            return list;
        }

        /// <summary>
        /// 为新添的编号设定样式
        /// </summary>
        /// <param name="listLevel"></param>
        /// <param name="font"></param>
        /// <param name="color"></param>
        public void SetFontForAddHeader(Aspose.Words.Lists.ListLevel listLevel, System.Drawing.Font font, Color color)
        {
            listLevel.Font.Size = font.Size;
            listLevel.Font.Color = color;
            listLevel.Font.Bold = font.Bold;
            listLevel.Font.Italic = font.Italic;
            listLevel.Font.Name = font.Name;
        }

        private void btnAddHF_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.InitialDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            if (of.ShowDialog() == DialogResult.OK)
            {
                string content = string.Empty;
                string realName = Path.GetFileName(of.FileName);
                string CopyPath = Common.CommonTool.retSavePath(of.FileName);
                Document doc = new Aspose.Words.Document(CopyPath);

                AddHeaderFooter(doc);

                string savePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Path.GetFileNameWithoutExtension(realName) + "_AsposeWord.docx");
                doc.Save(savePath, SaveFormat.Docx);
                Process.Start(savePath);
                textBox1.Text = content;
            }
        }

        /// <summary>
        /// 添加页眉/页脚
        /// </summary>
        /// <param name="doc"></param>
        public void AddHeaderFooter(Document doc)
        {
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.PageSetup.DifferentFirstPageHeaderFooter = true;
            //builder.PageSetup = true;
            // builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

            //builder.MoveToSection(1);
            PageSetup pageSetup = builder.PageSetup;
            //PageSetup pageSetup = doc.Sections[1].PageSetup;
            pageSetup.PageStartingNumber = 1;
            pageSetup.RestartPageNumbering = true;
            pageSetup.PageNumberStyle = NumberStyle.Arabic;
            pageSetup.HeaderDistance = 1.75 * 28.32;
            pageSetup.FooterDistance = 1.75 * 28.32;
            //pageSetup.
            //pageSetup



            builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
            builder.Write(tb_header.Text);
            //builder.InsertImage(@"G:\C#\wordTool\wordTestFrm\wordTestFrm\Resources\favicon-20200407023103330.ico", 30, 30);//可以控制图片的宽高
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
            builder.Write(tb_header.Text);
            builder.InsertImage(@"G:\C#\wordTool\wordTestFrm\wordTestFrm\Resources\favicon-20200407023103330.ico", 30, 30);//可以控制图片的宽高
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            //builder.Write(tb_header.Text);
            builder.InsertImage(@"G:\C#\wordTool\wordTestFrm\wordTestFrm\Resources\favicon-20200407023103330.ico", 30, 30);//可以控制图片的宽高
            builder.MoveToHeaderFooter(HeaderFooterType.FooterFirst);
            builder.Write(tb_footer.Text);
            builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);
            builder.Write(tb_footer.Text);
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            //builder.Write(tb_footer.Text);
            builder.Write(tb_footer.Text);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            builder.Write("~");
            builder.InsertField("PAGE", "");
            builder.Write("~");
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;



            ///全部页面设置
            builder.PageSetup.PaperSize = PaperSize.A4;//A4纸
            builder.PageSetup.Orientation = Aspose.Words.Orientation.Portrait;//方向
            builder.PageSetup.VerticalAlignment = Aspose.Words.PageVerticalAlignment.Top;//垂直对准
            builder.PageSetup.LeftMargin = 10;//页面左边距
            builder.PageSetup.RightMargin = 42;//页面右边距
            builder.PageSetup.TopMargin = 20;//页面上边距
            builder.PageSetup.BottomMargin = 20;//页面下边距

            //SetStyleForHeaderFooter(doc);
        }


        public string CopyToWorkDir(string filePath)
        {
            string workDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Works", "CopyDocs");
            if (!Directory.Exists(workDir))
            {
                Directory.CreateDirectory(workDir);
            }
            string fileName = Path.GetFileName(filePath);
            string savePath = Path.Combine(workDir, fileName);
            return savePath;
        }

        private void btnParagraphContent_Click(object sender, EventArgs e)
        {
            string val2 = string.Empty;
            OpenFileDialog of = new OpenFileDialog();
            of.InitialDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            if (of.ShowDialog() == DialogResult.OK)
            {
                string realName = Path.GetFileName(of.FileName);
                string CopyPath = Common.CommonTool.retSavePath(of.FileName);

                Document doc = new Document(CopyPath);

                FrmWordStruct fws = new FrmWordStruct();
                fws.doc = doc;
                fws.InitTreeView(true);
                fws.ShowDialog();

                string savePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Path.GetFileNameWithoutExtension(realName) + "_AsposeWord.docx");
                doc.Save(savePath, SaveFormat.Docx);
                Process.Start(savePath);
            }
        }



        private void btnHFTest_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.InitialDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            if (of.ShowDialog() == DialogResult.OK)
            {
                string content = string.Empty;
                string realName = Path.GetFileName(of.FileName);
                string CopyPath = Common.CommonTool.retSavePath(of.FileName);
                DocToTxtWriter myDocToTxtWriter = new DocToTxtWriter();
                Document doc = new Aspose.Words.Document(CopyPath);

                NodeCollection items = doc.GetChildNodes(NodeType.Any, true);
                NodeCollection headfooters = doc.GetChildNodes(NodeType.HeaderFooter, true);
                HeaderFooter hf = (HeaderFooter)headfooters[0];
                var aa = hf.ParentSection;
                for (int i = 0; i < aa.Count; i++)
                {
                    //var item=aa.
                }
                var aas = aa.FirstChild.GetText();
                var ss = hf.FirstChild;
                //ss.fri

                HeaderFooter h1 = new HeaderFooter(doc, HeaderFooterType.HeaderFirst);
                Paragraph p_F = new Paragraph(doc);
                p_F.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                Run run1_l = new Run(doc, "首页-左");
                Run run1_c = new Run(doc, "首页-中");
                Run run1_r = new Run(doc, "首页-右");
                run1_l.Accept(myDocToTxtWriter);
                run1_c.Accept(myDocToTxtWriter);
                run1_r.Accept(myDocToTxtWriter);

                Run run1_sp = new Run(doc, "initial text. ");
                //AbsolutePositionTab tab = (AbsolutePositionTab)run1_l;
                //p_F.ParagraphFormat.Alignment = ParagraphAlignment.ThaiDistributed;
                //p_F.AppendChild(run1_sp);
                p_F.Runs.Add(run1_l);//p_F.InsertBefore(run1_l, run1_sp); // p_F.Runs.inser(run1_l);
                p_F.Runs.Add(run1_c);//p_F.InsertBefore(run1_c, run1_sp); //p_F.Runs.Add(run1_c);
                p_F.Runs.Add(run1_r);//p_F.InsertBefore(run1_r, run1_sp); //p_F.Runs.Add(run1_r);

                h1.Paragraphs.Add(p_F);

                HeaderFooter h2 = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);

                Run run2 = new Run(doc, "奇数");

                Paragraph p_P = new Paragraph(doc);
                p_P.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                p_P.Runs.Add(run2);
                h2.Paragraphs.Add(p_P);

                HeaderFooter h3 = new HeaderFooter(doc, HeaderFooterType.HeaderEven);

                Run run3 = new Run(doc, "偶数");

                Paragraph p_E = new Paragraph(doc);
                p_E.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                p_E.Runs.Add(run3);
                p_E.InsertField("PAGE", run3, true);
                p_E.InsertField("NUMPAGES", run3, true);
                h3.Paragraphs.Add(p_E);

                PageSetup ps = doc.Sections[0].PageSetup;

                ps.DifferentFirstPageHeaderFooter = true;
                ps.OddAndEvenPagesHeaderFooter = true;
                ps.PageStartingNumber = 1;
                ps.PageNumberStyle = NumberStyle.Arabic;
                ps.RestartPageNumbering = true;


                doc.Sections[0].HeadersFooters.Add(h1);
                doc.Sections[0].HeadersFooters.Add(h2);
                doc.Sections[0].HeadersFooters.Add(h3);

                string savePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Path.GetFileNameWithoutExtension(realName) + "_AsposeWord.docx");
                doc.Save(savePath, SaveFormat.Docx);
                Process.Start(savePath);
                /*
                for (int i = 0; i < items.Count; i++)
                {
                    Node item = items[i];
                    if(item is Aspose.Words.HeaderFooter)
                    {
                        Aspose.Words.HeaderFooter hf = (Aspose.Words.HeaderFooter)item;
                        
                    }
                    if (item is Aspose.Words.Run)
                    {
                        Aspose.Words.Run run = (Aspose.Words.Run)item;
                        if (run.GetText() == "左")
                        {

                        }
                        else if (run.GetText() == "中")
                        {

                        }
                        if (run.GetText() == "右")
                        {

                        }
                    }
                    string text = item.GetText();
                }
                */
            }
        }

        private void btnHFTest2_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.InitialDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            if (of.ShowDialog() == DialogResult.OK)
            {
                string content = string.Empty;
                //string realName = Path.GetFileName(of.FileName);
                //string CopyPath = this.retSavePath(of.FileName);
                Document doc = new Aspose.Words.Document(of.FileName);

                //NodeCollection items = doc.GetChildNodes(NodeType.Paragraph, true);
                NodeCollection allItems = doc.GetChildNodes(NodeType.Any, true);
                foreach (Node item in allItems)
                {
                    NodeType ss = item.NodeType;
                }


                Document newDoc = new Document();

                #region 表格
                //NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
                //for (int i = 0; i < tables.Count; i++)
                //{
                //    Aspose.Words.Tables.Table table = (Aspose.Words.Tables.Table)tables[i];
                //    Aspose.Words.Tables.Table newTable = new Aspose.Words.Tables.Table(newDoc);
                //    for (int j = 0; j < table.Rows.Count; j++)
                //    {
                //        Aspose.Words.Tables.Row row = table.Rows[j];
                //        Aspose.Words.Tables.Row newRow = new Aspose.Words.Tables.Row(newDoc);
                //        for (int h = 0; h < row.Cells.Count; h++)
                //        {
                //            Aspose.Words.Tables.Cell cell = row.Cells[h];
                //            Aspose.Words.Tables.Cell newCell = new Aspose.Words.Tables.Cell(newDoc);
                //            NodeCollection nodes_Cell = cell.GetChildNodes(NodeType.Paragraph, true);
                //            foreach (Paragraph item in nodes_Cell)
                //            {
                //                Paragraph newP = new Paragraph(newDoc);
                //                newP = addParagrah(item, newP, newDoc);
                //                newCell.ChildNodes.Add(newP);
                //            }
                //            setStyleForNewCell(cell, newCell);
                //            newRow.Cells.Add(newCell);
                //        }
                //        newTable.Rows.Add(newRow);
                //    }
                //    newDoc.FirstSection.Body.AppendChild(newTable);
                //    Paragraph p = new Paragraph(newDoc);
                //    Run r = new Run(newDoc,"/r");
                //    p.Runs.Add(r);
                //    newDoc.FirstSection.Body.AppendChild(p);
                //}
                #endregion

                #region 段落
                NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
                foreach (Node item in paragraphs)
                {
                    Paragraph p = (Paragraph)item;
                    Paragraph newP = new Paragraph(newDoc);
                    //表格文字  TOC 1  TOC 2
                    string sName = p.ParagraphFormat.StyleName;
                    switch (p.ParagraphFormat.StyleName)
                    {
                        case "Heading 1":
                            newP.ParagraphFormat.StyleName = "Heading 1";
                            newP.ParagraphFormat.OutlineLevel = OutlineLevel.Level1;
                            newP = addParagrah(p, newP, newDoc);
                            newDoc.FirstSection.Body.AppendChild(newP);
                            break;
                        case "Heading 2":
                            newP.ParagraphFormat.StyleName = "Heading 2";
                            newP.ParagraphFormat.OutlineLevel = OutlineLevel.Level2;
                            newP = addParagrah(p, newP, newDoc);
                            newDoc.FirstSection.Body.AppendChild(newP);
                            break;
                        case "Heading 3":
                            newP.ParagraphFormat.StyleName = "Heading 3";
                            newP.ParagraphFormat.OutlineLevel = OutlineLevel.Level3;
                            newP = addParagrah(p, newP, newDoc);
                            newDoc.FirstSection.Body.AppendChild(newP);
                            break;
                        case "Heading 4":
                            newP.ParagraphFormat.StyleName = "Heading 4";
                            newP.ParagraphFormat.OutlineLevel = OutlineLevel.Level4;
                            newP = addParagrah(p, newP, newDoc);
                            newDoc.FirstSection.Body.AppendChild(newP);
                            break;
                        case "Heading 5":
                            newP.ParagraphFormat.StyleName = "Heading 5";
                            newP.ParagraphFormat.OutlineLevel = OutlineLevel.Level5;
                            newP = addParagrah(p, newP, newDoc);
                            newDoc.FirstSection.Body.AppendChild(newP);
                            break;
                        case "Normal":
                        case "Body Text":
                        case "Body Text First Indent":
                            newP.ParagraphFormat.StyleName = "Normal";
                            newP.ParagraphFormat.OutlineLevel = OutlineLevel.BodyText;
                            newP = addParagrah(p, newP, newDoc);
                            newDoc.FirstSection.Body.AppendChild(newP);
                            break;
                    }
                }
                #endregion

                string fileName = Path.Combine(@"I:\word_doc\newDocs", Path.GetFileNameWithoutExtension(of.FileName) + "_new.docx");
                newDoc.Save(fileName, SaveFormat.Docx);
                Process.Start(fileName);
                #region MyRegion
                //doc.Sections[0].ChildNodes.Add();


                //for (int i = 0; i < items.Count; i++)
                //{
                //    Node item = items[i];
                //    if (item is Aspose.Words.HeaderFooter)
                //    {
                //        Aspose.Words.HeaderFooter hf = (Aspose.Words.HeaderFooter)item;

                //    }
                //    if (item is Aspose.Words.AbsolutePositionTab)
                //    {
                //        //Aspose.Words.Drawing.Shape
                //        AbsolutePositionTab tab = (AbsolutePositionTab)item;
                //    }
                //    else if (item is Aspose.Words.Run)
                //    {


                //        Aspose.Words.Run run = (Aspose.Words.Run)item;
                //        //run.Accept(myDocToTxtWriter);
                //        if (run.GetText() == "左")
                //        {

                //        }
                //        else if (run.GetText() == "中")
                //        {

                //        }
                //        if (run.GetText() == "右")
                //        {

                //        }
                //    }
                //    string text = item.NodeType.ToString() + " " + item.GetText() + "\r\n";
                //    content += text;
                //}
                //textBox1.Text = content; 
                #endregion
            }
        }

        public void setStyleForNewCell(Aspose.Words.Tables.Cell cell, Aspose.Words.Tables.Cell newCell)
        {
            newCell.CellFormat.HorizontalMerge = cell.CellFormat.HorizontalMerge;
            newCell.CellFormat.VerticalMerge = cell.CellFormat.VerticalMerge;
            newCell.CellFormat.Borders[BorderType.Top].Color = cell.CellFormat.Borders[BorderType.Top].Color;
            newCell.CellFormat.Borders[BorderType.Top].LineStyle = cell.CellFormat.Borders[BorderType.Top].LineStyle;
            newCell.CellFormat.Borders[BorderType.Top].LineWidth = cell.CellFormat.Borders[BorderType.Top].LineWidth;
            newCell.CellFormat.Borders[BorderType.Bottom].Color = cell.CellFormat.Borders[BorderType.Bottom].Color;
            newCell.CellFormat.Borders[BorderType.Bottom].LineStyle = cell.CellFormat.Borders[BorderType.Bottom].LineStyle;
            newCell.CellFormat.Borders[BorderType.Bottom].LineWidth = cell.CellFormat.Borders[BorderType.Bottom].LineWidth;
            newCell.CellFormat.Borders[BorderType.Left].Color = cell.CellFormat.Borders[BorderType.Left].Color;
            newCell.CellFormat.Borders[BorderType.Left].LineStyle = cell.CellFormat.Borders[BorderType.Left].LineStyle;
            newCell.CellFormat.Borders[BorderType.Left].LineWidth = cell.CellFormat.Borders[BorderType.Left].LineWidth;
            newCell.CellFormat.Borders[BorderType.Right].Color = cell.CellFormat.Borders[BorderType.Right].Color;
            newCell.CellFormat.Borders[BorderType.Right].LineStyle = cell.CellFormat.Borders[BorderType.Right].LineStyle;
            newCell.CellFormat.Borders[BorderType.Right].LineWidth = cell.CellFormat.Borders[BorderType.Right].LineWidth;
        }

        private Paragraph addParagrah(Paragraph p, Paragraph newP, Document newDoc)
        {
            foreach (Run r in p.Runs)
            {
                Run newr = new Run(newDoc, r.GetText());
                newP.Runs.Add(newr);
            }
            setStyleFormP(p, newP);
            return newP;
            //newDoc.FirstSection.Body.AppendChild(newP);
        }

        private void setStyleFormP(Paragraph p, Paragraph newP)
        {
            int i = 0;
            foreach (Run item in newP.Runs)
            {
                item.Font.Size = p.Runs[i].Font.Size;
                item.Font.Color = p.Runs[i].Font.Color;
                item.Font.Bold = p.Runs[i].Font.Bold;
                item.Font.Italic = p.Runs[i].Font.Italic;
                item.Font.Name = p.Runs[i].Font.Name;
                item.ParentParagraph.ParagraphFormat.LineSpacing = p.Runs[i].ParentParagraph.ParagraphFormat.LineSpacing;
                item.ParentParagraph.ParagraphFormat.SpaceAfter = p.Runs[i].ParentParagraph.ParagraphFormat.SpaceAfter;
                item.ParentParagraph.ParagraphFormat.SpaceBefore = p.Runs[i].ParentParagraph.ParagraphFormat.SpaceBefore;
                item.ParentParagraph.ParagraphFormat.LeftIndent = p.Runs[i].ParentParagraph.ParagraphFormat.LeftIndent;
                item.ParentParagraph.ParagraphFormat.RightIndent = p.Runs[i].ParentParagraph.ParagraphFormat.RightIndent;
                i++;
            }
        }

        /// <summary>
        /// 英寸到里面
        /// </summary>
        /// <param name="inch"></param>
        /// <returns></returns>
        public double InchToCM(double inch)
        {
            return inch * 2.54;
        }

        /// <summary>
        /// 像素到英寸
        /// </summary>
        /// <param name="Pixel"></param>
        /// <param name="dpi"></param>
        /// <returns></returns>
        public double PixelToInch(double Pixel, int dpi)
        {
            return Pixel / dpi;
        }

        /// <summary>
        /// 像素到英寸
        /// </summary>
        /// <param name="Pixel"></param>
        /// <param name="dpi"></param>
        /// <returns></returns>
        public double InchToPixel(double Inch, int dpi)
        {
            return Inch * dpi;
        }


        /// <summary>
        /// 厘米到英寸
        /// </summary>
        /// <param name="cm"></param>
        /// <returns></returns>
        public double CmToInch(double cm)
        {
            return cm / 2.54;
        }

        private void btnImgInsert_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.InitialDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            if (of.ShowDialog() == DialogResult.OK)
            {
                Image image = Image.FromFile("test.bmp");
                string content = string.Empty;
                string realName = Path.GetFileName(of.FileName);
                string CopyPath = Common.CommonTool.retSavePath(of.FileName);
                Document doc = new Aspose.Words.Document(CopyPath);

                DocumentBuilder builder = new DocumentBuilder(doc);
                //string imageUrl = "c:/1.jpg";
                //if (File.Exists(imageUrl))
                //{

                double shapeWidthInches = ConvertUtil.PointToInch(200);
                double shapeHeightInches = ConvertUtil.PointToInch(150);
                double displayWidth = 200;//显示宽度
                double zoomRate = displayWidth / image.Width;//缩放比例
                double displayHeigth = image.Height * zoomRate;//显示高度
                int dpi = 96;
                double width_CM = 20;
                double height_CM = 15;
                double width_inch = CmToInch(width_CM);
                double height_inch = CmToInch(height_CM);

                displayWidth = ConvertUtil.MillimeterToPoint(15 * 10);
                displayHeigth = ConvertUtil.MillimeterToPoint(8.5 * 10);
                Shape shape = null;

                //shape.Rotation = 96;

                //builder.InsertShape(shape,displayWidth,displayWidth);
                shape = builder.InsertImage(image, displayWidth, displayHeigth);
                //shape = builder.InsertHorizontalRule();
                shape.ImageData.SetImage(image);
                SizeF s = shape.SizeInPoints;
                width_inch = ConvertUtil.PointToInch(shape.Width);
                height_inch = ConvertUtil.PointToInch(shape.Height);
                width_CM = this.InchToCM(width_inch);
                height_CM = this.InchToCM(height_inch);
                double width = 200;
                width_CM = 10;
                double v = ConvertUtil.MillimeterToPoint(10 * 10) / image.Width;
                displayHeigth = v * image.Height;
                displayWidth = ConvertUtil.MillimeterToPoint(width_CM * 10);
                //displayHeigth = ConvertUtil.MillimeterToPoint(height_CM*10);
                //shape.Width = displayWidth;
                //shape.Height = displayHeigth;
                shape.AspectRatioLocked = false;
                shape.AnchorLocked = false;
                //shape.HorizontalRuleFormat.NoShade = false;
                //shape.HorizontalRuleFormat.Height = 50;
                //shape.HorizontalRuleFormat.WidthPercent = 100;
                //shape.HorizontalRuleFormat.NoShade = false;
                //double w = shape.Width;

                //shape.Width = width;

                //shape.Height = (width / w) * shape.Height;

                //shape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;

                //shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Character;

                //shape.WrapType = WrapType.None;

                string savePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Path.GetFileNameWithoutExtension(realName) + "_AsposeWord.docx");
                doc.Save(savePath, SaveFormat.Docx);
                Process.Start(savePath);
                textBox1.Text = content;
            }

        }

        /// <summary>
        /// Visitor implementation that simply collects the Runs and AbsolutePositionTabs of a document as plain text. 
        /// </summary>
        public class DocToTxtWriter : DocumentVisitor
        {
            public DocToTxtWriter()
            {
                mBuilder = new StringBuilder();
            }

            /// <summary>
            /// Called when a Run node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitRun(Run run)
            {
                AppendText(run.Text);
                // Let the visitor continue visiting other nodes.
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when an AbsolutePositionTab node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitAbsolutePositionTab(AbsolutePositionTab tab)
            {
                // We'll treat the AbsolutePositionTab as a regular tab in this case
                mBuilder.Append("\t");
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Adds text to the current output. Honors the enabled/disabled output flag.
            /// </summary>
            private void AppendText(string text)
            {
                mBuilder.Append(text);
            }

            /// <summary>
            /// Gets the plain text of the document that was accumulated by the visitor.
            /// </summary>
            public string GetText()
            {
                return mBuilder.ToString();
            }

            private readonly StringBuilder mBuilder;
        }

        private void btnlayOut_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.InitialDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            if (of.ShowDialog() == DialogResult.OK)
            {
                string content = string.Empty;
                string realName = Path.GetFileName(of.FileName);
                string CopyPath = Common.CommonTool.retSavePath(of.FileName);
                ClassTest ct = new ClassTest();
                Document doc = new Document(CopyPath);

                double t1 = 207.65;
                double t2 = 434;

                HeaderFooter header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
                Paragraph p = new Paragraph(doc);
                TabStopCollection tabStops = p.ParagraphFormat.TabStops;
                PaperSize pSize = doc.Sections[0].PageSetup.PaperSize;

                double S_Width = doc.Sections[0].PageSetup.PageWidth;
                S_Width = doc.Sections[0].PageSetup.PageWidth - doc.Sections[0].PageSetup.HeaderDistance * 2;
                //S_Width = doc.Sections[0].PageSetup.FooterDistance;
                tabStops.Add(new Aspose.Words.TabStop(t1, Aspose.Words.TabAlignment.Left, TabLeader.None));
                tabStops.Add(new Aspose.Words.TabStop(t2, Aspose.Words.TabAlignment.Right, TabLeader.None));
                Run run_l = new Run(doc);

                GraphicsPath graphicsPathObj = new GraphicsPath();
                string stringText = "left";
                FontFamily family = new FontFamily(run_l.Font.Name);
                int fontStyle = (int)FontStyle.Regular;
                float emSize = (float)run_l.Font.Size;
                Point origin = new Point(0, 0);
                StringFormat format = StringFormat.GenericDefault;
                graphicsPathObj.AddString(stringText,
                family,
                fontStyle,
                emSize,
                origin,
                format);
                RectangleF rcBound = graphicsPathObj.GetBounds();

                run_l.Text = "left" + ControlChar.Tab;

                p.Runs.Add(run_l);



                Run run_c = new Run(doc);
                run_c.Text = "center" + ControlChar.Tab;
                p.Runs.Add(run_c);

                Run run_r = new Run(doc);
                run_r.Text = "right";
                p.Runs.Add(run_r);

                header.Paragraphs.Add(p);
                doc.Sections[0].HeadersFooters.Add(header);

                string savePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Path.GetFileNameWithoutExtension(realName) + "_AsposeWord.docx");
                doc.Save(savePath, SaveFormat.Docx);
                Process.Start(savePath);
            }
        }

        /// <summary>
        /// 生成存储目录
        /// </summary>
        /// <param name="SaveDir">存储目录</param>
        /// <param name="fileName">文件名</param>
        /// <param name="flagStr">生成的特有后缀</param>
        /// <returns>反馈唯一的存储文件路径</returns>
        public string retSaveFilePath(string SaveDir, string fileName, string flagStr = "Copy")
        {
            string filePath = string.Empty;
            if (!Directory.Exists(SaveDir)) Directory.CreateDirectory(SaveDir);
            string tmpPath = Path.GetFileNameWithoutExtension(fileName) + "_" + flagStr + "_";
            List<string> items = Directory.GetFiles(SaveDir).ToList().FindAll(item => item.Contains(tmpPath));
            string file = items.OrderByDescending(item =>
            {
                string tmpVal = item.Substring(item.LastIndexOf('_') + 1);
                try
                {
                    int result = int.Parse(tmpVal.Substring(0, tmpVal.LastIndexOf('.')));
                    return result;
                }
                catch (FormatException ex)
                {
                    return 0;
                }
                catch (Exception ex)
                {
                    return 0;
                }

            }).Max();
            string val = file.Substring(file.LastIndexOf('_') + 1);
            int maxIndex = int.Parse(val.Substring(0, val.LastIndexOf('.'))) + 1;
            string fileType = fileName.Substring(fileName.LastIndexOf('.'));
            string saveFileName = Path.GetFileNameWithoutExtension(fileName) + "_" + flagStr + "_" + maxIndex.ToString("00") + fileType;
            filePath = Path.Combine(SaveDir, saveFileName);


            /*
            int index = 1;
            while(true)
            {
                string saveFileName = Path.GetFileNameWithoutExtension(fileName) + "_" + flagStr + "_" + index.ToString("00") + fileType;
                filePath = Path.Combine(SaveDir, saveFileName);
                if(!File.Exists(filePath))
                {
                    break;
                }
                index++;
            }*/
            return filePath;
        }

        //        OpenFileDialog openFile = new OpenFileDialog();
        //            if (openFile.ShowDialog() == DialogResult.OK)
        //            {
        //                string filePath = openFile.FileName;
        //        string fileType = openFile.SafeFileName.Substring(openFile.SafeFileName.LastIndexOf('.'));
        //                for (int i = 0; i< 5; i++)
        //                {
        //                    string savePath = this.retSaveFilePath(AppDomain.CurrentDomain.BaseDirectory, openFile.SafeFileName);

        //        File.Copy(filePath, savePath);
        //                }

        //}


        private void btnPageNum_Click(object sender, EventArgs e)
        {
            Color a = this.ucFont1.fontColorSelect;
            System.Drawing.Font b = this.ucFont1.fontSelect;
            OpenFileDialog of = new OpenFileDialog();
            of.InitialDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            if (of.ShowDialog() == DialogResult.OK)
            {
                string realName = Path.GetFileName(of.FileName);
                string CopyPath = Common.CommonTool.retSavePath(of.FileName);
                Document doc = new Document(CopyPath);
                NodeCollection nodes = doc.GetChildNodes(NodeType.Paragraph,true);
                for (int i = 0; i < nodes.Count; i++)
                {
                    Paragraph p =(Paragraph) nodes[i];
                    for (int j = 0; j < p.Runs.Count; j++)
                    {
                        var txt = p.Runs[j].GetText();
                        var ss = p.Runs[j].Font.Size;
                    }
                }
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.MoveToDocumentStart();
                builder.Write("目录");
                builder.InsertBreak(BreakType.PageBreak);

                builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

                // The newly inserted table of contents will be initially empty.
                // It needs to be populated by updating the fields in the document.
                doc.UpdateFields();
                string savePath = Common.CommonTool.retSaveFilePath(
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WorkSpace", Path.GetFileNameWithoutExtension(of.SafeFileName)),
                    of.SafeFileName
                    );
                doc.Save(savePath, SaveFormat.Docx);
                Process.Start(savePath);
                #region MyRegion
                //string content = string.Empty;
                //string realName = Path.GetFileName(of.FileName);
                //string CopyPath = Common.CommonTool.retSavePath(of.FileName);
                //Document doc = new Document(CopyPath);
                //NodeCollection nodes = doc.GetChildNodes(NodeType.HeaderFooter, true);
                //for (int i = 0; i < nodes.Count; i++)
                //{
                //    HeaderFooter f = (HeaderFooter)nodes[i];
                //    if (f.HeaderFooterType == HeaderFooterType.FooterPrimary)
                //    {
                //        NodeCollection items = f.GetChildNodes(NodeType.FieldStart, true);
                //        if (items.Count > 0)
                //        {
                //            Aspose.Words.Fields.FieldStart fs = (Aspose.Words.Fields.FieldStart)items[0];
                //            if (fs.FieldType == FieldType.FieldPage)
                //            {

                //            }
                //        }
                //    }
                //} 
                #endregion
            }
        }

        private void btnLineSpacing_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.InitialDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            if (of.ShowDialog() == DialogResult.OK)
            {

                string content = string.Empty;
                string realName = Path.GetFileName(of.FileName);
                string CopyPath = Common.CommonTool.retSavePath(of.FileName);
                Document doc = new Document(CopyPath);
                NodeCollection nodes = doc.GetChildNodes(NodeType.Paragraph, true);
                Paragraph p = (Paragraph)nodes[0];
                p.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
                p.ParagraphFormat.LineSpacing = 2 * 12;

                string savePath = Common.CommonTool.retSaveFilePath(
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WorkSpace", Path.GetFileNameWithoutExtension(of.SafeFileName)),
                    of.SafeFileName
                    );
                doc.Save(savePath, SaveFormat.Docx);
                Process.Start(savePath);
            }
        }

        private void ucFont1_Load(object sender, EventArgs e)
        {

        }

        private void ucCellAlignment1_LabelClick(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.InitialDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            if (of.ShowDialog() == DialogResult.OK)
            {

                string content = string.Empty;
                string realName = Path.GetFileName(of.FileName);
                string CopyPath = Common.CommonTool.retSavePath(of.FileName);
                Document doc = new Document(CopyPath);


                Label label = (Label)sender;
                int index = Convert.ToInt16(label.Tag);

                NodeCollection nodes = doc.GetChildNodes(NodeType.Table, true);
                Aspose.Words.Tables.Table tb = (Aspose.Words.Tables.Table)nodes[0];
                CellVerticalAlignment verticalAlignment = CellVerticalAlignment.Center;
                ParagraphAlignment paragraphAlignment = ParagraphAlignment.Center;
                switch ((Enum_CellAlignment)index)
                {
                    case Enum_CellAlignment.LeftUp:
                        verticalAlignment = CellVerticalAlignment.Top;
                        paragraphAlignment = ParagraphAlignment.Left;
                        break;
                    case Enum_CellAlignment.LeftMiddle:
                        verticalAlignment = CellVerticalAlignment.Center;
                        paragraphAlignment = ParagraphAlignment.Left;
                        break;
                    case Enum_CellAlignment.LeftBottom:
                        verticalAlignment = CellVerticalAlignment.Bottom;
                        paragraphAlignment = ParagraphAlignment.Left;
                        break;
                    case Enum_CellAlignment.CenterMiddle:
                        verticalAlignment = CellVerticalAlignment.Center;
                        paragraphAlignment = ParagraphAlignment.Center;
                        break;
                    case Enum_CellAlignment.CenterUp:
                        verticalAlignment = CellVerticalAlignment.Top;
                        paragraphAlignment = ParagraphAlignment.Center;
                        break;
                    case Enum_CellAlignment.CenterBottom:
                        verticalAlignment = CellVerticalAlignment.Bottom;
                        paragraphAlignment = ParagraphAlignment.Center;
                        break;
                    case Enum_CellAlignment.RightBottom:
                        verticalAlignment = CellVerticalAlignment.Bottom;
                        paragraphAlignment = ParagraphAlignment.Right;
                        break;
                    case Enum_CellAlignment.RightMiddle:
                        verticalAlignment = CellVerticalAlignment.Center;
                        paragraphAlignment = ParagraphAlignment.Right;
                        break;
                    case Enum_CellAlignment.RightUp:
                        verticalAlignment = CellVerticalAlignment.Top;
                        paragraphAlignment = ParagraphAlignment.Right;
                        break;
                }

                for (int i = 0; i < tb.Rows.Count; i++)
                {
                    Row row = tb.Rows[i];
                    for (int a = 0; a < row.Cells.Count; a++)
                    {
                        Cell cell = row.Cells[a];
                        cell.CellFormat.VerticalAlignment = verticalAlignment;
                        NodeCollection nodesCP = cell.GetChildNodes(NodeType.Paragraph, true);
                        for (int k = 0; k < nodesCP.Count; k++)
                        {
                            Paragraph p = (Paragraph)nodesCP[k];
                            p.ParagraphFormat.Alignment = paragraphAlignment;
                        }
                    }
                }

                string savePath = Common.CommonTool.retSaveFilePath(
                   Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WorkSpace", Path.GetFileNameWithoutExtension(realName)),
                   realName
                   );
                doc.Save(savePath, SaveFormat.Docx);
                Process.Start(savePath);
            }
        }

        private void btnToc_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.InitialDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            if (of.ShowDialog() == DialogResult.OK)
            {

                string content = string.Empty;
                string realName = Path.GetFileName(of.FileName);
                string CopyPath = Common.CommonTool.retSavePath(of.FileName);
                Document doc = new Document(CopyPath);
                
                Paragraph p = new Paragraph(doc);
                //p.AppendChild(new BookmarkStart(doc, "bkToc"));
                p.AppendChild(new Run(doc, "目录"));

                Aspose.Words.Rendering.PageInfo pi = doc.GetPageInfo(0);
                int count = doc.PageCount;
                doc.Sections[0].Body.Paragraphs.Insert(1,p);
                
                //FieldToc fieldToc =  (FieldToc)p.InsertField(FieldType.FieldTOC,true, null, true);
                
                
    
                //p.AppendChild(new BookmarkEnd(doc, "bkToc"));

                //p.ParagraphFormat.PageBreakBefore = true;
                string allText = doc.Sections[0].GetText();
                int pageCount = 0;

          

                for (int i = 0; i < doc.Sections[0].Body.Paragraphs.Count; i++)
                {
                    Paragraph p1 = doc.Sections[0].Body.Paragraphs[i];
                    
                    Aspose.Words.Rendering.PageInfo pi1 = doc.GetPageInfo(0);
                    string s = p1.GetText();
                    
                    string ss= doc.Sections[0].GetText();
                    Aspose.Words.Font font1 = p1.ParagraphBreakFont;
                    if (p1.LastChild!=null && p1.GetText().Contains(ControlChar.PageBreak))
                    {
                        string contents = p1.GetText();
                        pageCount++;
                    }
        
                   
                }
                //DocumentBuilder builder = new DocumentBuilder(doc);
                //builder.MoveToField(fieldToc, true);
                //builder.InsertBreak(BreakType.PageBreak);
                
                //fieldToc.UpdatePageNumbers();
                doc.UpdateFields();

                string savePath = Common.CommonTool.retSaveFilePath(
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WorkSpace", Path.GetFileNameWithoutExtension(realName)),
                realName
                );
                doc.Save(savePath, SaveFormat.Docx);
                Process.Start(savePath);

            }
        }

        private void btnFirstPage_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.InitialDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            if (of.ShowDialog() == DialogResult.OK)
            {
                string content = string.Empty;
                string realName = Path.GetFileName(of.FileName);
                string CopyPath = Common.CommonTool.retSavePath(of.FileName);
                Document doc = new Document(CopyPath);

                bool isFirst = true;
                for (int i = 0; i < doc.Sections.Count; i++)
                {
                    Section section = doc.Sections[i];
                    for (int a = 0; a < section.Body.Paragraphs.Count; a++)
                    {
                        Paragraph p = section.Body.Paragraphs[a];
                        if (!isFirst)
                        {
                            for (int b = 0; b < p.Runs.Count; b++)
                            {
                                Run r = p.Runs[b];
                                r.Font.Color = Color.Red;
                            }
                        }
                        else if (p.GetText().Contains(ControlChar.PageBreak))//去除首页样式修改
                        {
                            isFirst = false;
                        }
                       
                    }
                }

                string savePath = Common.CommonTool.retSaveFilePath(
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WorkSpace", Path.GetFileNameWithoutExtension(realName)),
                realName
                );
                doc.Save(savePath, SaveFormat.Docx);
                Process.Start(savePath);
            }
        }

        private void btnPageSet_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.InitialDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            if (of.ShowDialog() == DialogResult.OK)
            {
                string content = string.Empty;
                string realName = Path.GetFileName(of.FileName);
                string CopyPath = Common.CommonTool.retSavePath(of.FileName);
                Document doc = new Document(CopyPath);
                
                
                int count = doc.PageCount;
                LayoutEnumerator layoutEnumerator = new LayoutEnumerator(doc);
                LayoutCollector layoutCollector = new LayoutCollector(doc);

                for (int i = 0; i < doc.Sections.Count; i++)
                {
                    Section section = doc.Sections[i];
                    string textAll_Section = section.Range.Text;
                    for (int a = 0; a < section.Body.Paragraphs.Count; a++)
                    {
                        Paragraph p = section.Body.Paragraphs[a];
                        string text = p.Range.Text;
                        layoutEnumerator.Current = layoutCollector.GetEntity(p);
                       
                        int pageNum = layoutEnumerator.PageIndex;
                        Aspose.Words.Rendering.PageInfo pi = doc.GetPageInfo(5);
                        PaperSize paperSize = pi.PaperSize;
                        SizeF sizeF = pi.SizeInPoints;
                    }
                 
                    
                    
                    //pi.GetDotNetPaperSize().sizeF = sizeF;
                }


            }
        }
    }
}
