using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Layout;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using wordTestFrm.Model;
using wordTestFrm.models;

namespace wordTestFrm
{
    public class CommonMethods
    {
        /// <summary>
        /// 拷贝文件
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static string retSavePath(string filePath)
        {
            string fileName = Path.GetFileName(filePath);
            string saveDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Work", "CopyFile");
            if (!Directory.Exists(saveDir))
            {
                Directory.CreateDirectory(saveDir);
            }

            string savePath = Path.Combine(saveDir, fileName);
            File.Copy(filePath, savePath, true);
            return savePath;
        }

        public static string retSaveFilePath(string SaveDir, string fileName, string flagStr = "Copy")
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
            return filePath;
        }


        /// <summary>
        /// 保存文档样式
        /// </summary>
        /// <param name="word"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static int saveTxtOfWordStyle(WordStyle word, string fileName)
        {
            try
            {
                string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WordStyles");
                if (!File.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                string result = JsonConvert.SerializeObject(word);
                string ss = Path.Combine(path, fileName);
                FileStream fs = new FileStream(ss, FileMode.Create, FileAccess.Write);
                StreamWriter sw = new StreamWriter(fs);
                //开始写入
                sw.Write(result);
                //清空缓冲区
                sw.Flush();
                //关闭流
                sw.Close();
                fs.Close();
                return 1;
            }
            catch (Exception ex)
            {
                MessageBox.Show("设定格式失败");
                return -1;
            }
        }

        /// <summary>
        /// 从选项获取行距
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public static double GetLineSpaceForComboBox(string item)
        {
            switch (item)
            {
                case "单倍行距":
                    return 1;
                case "1.5倍行距":
                    return 1.5;
                case "2倍行距":
                    return 2;
                case "最小值":
                    return 1;
                case "固定值":
                    return 1;
                case "多倍行距":
                    return 3;
                default:
                    return 1;
            }
        }

        public static PaperSize GetPageSizeForComboBox(string item)
        {
            switch (item)
            {
                case "A3": return PaperSize.A3;
                case "A4": return PaperSize.A4;
                case "A5": return PaperSize.A5;
                case "B4": return PaperSize.B4;
                case "B5": return PaperSize.B5;
                default: return PaperSize.A4;
            }
        }

        public static Aspose.Words.Orientation GetPageDirectionForComboBox(string item)
        {
            switch (item)
            {
                case "纵向": return Aspose.Words.Orientation.Portrait;
                case "横向": return Aspose.Words.Orientation.Landscape;
                default: return Aspose.Words.Orientation.Landscape;
            }
        }

        /// <summary>
        /// 获取表格样式
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public static LineStyle GetLineStyleForComboBox(string item)
        {
            switch (item)
            {
                case "1": return LineStyle.Single;
                case "7": return LineStyle.DashLargeGap;
                case "6": return LineStyle.Dot;
                case "8": return LineStyle.DotDash;
                case "9": return LineStyle.DotDotDash;
                case "3": return LineStyle.Double;
                case "10": return LineStyle.Triple;
                case "11": return LineStyle.ThinThickSmallGap;
                default:
                    return LineStyle.Single;
            }
        }

        /// <summary>
        /// 获取段前/段后
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public static double GetSpaceBeforeOrAfterForComboBox(string item)
        {
            switch (item)
            {
                case "0 行": return 0;
                case "0.5 行": return 0.5;
                case "1 行": return 1;
                case "1.5 行": return 1.5;
                case "2 行": return 2;
                case "2.5 行": return 2.5;
                case "3 行": return 3;
                case "3.5 行": return 3.5;
                case "4 行": return 4;
                default: return 0;
            }
        }

        public static string GetFontSize(double item)
        {
            switch (item)
            {
                #region word中大小
                //case 42: return "初号";
                //case 36: return "小初";
                //case 26: return "一号";
                //case 24: return "小一";
                //case 22: return "二号";
                //case 18: return "小二";
                //case 16: return "三号";
                //case 15: return "小三";
                //case 14: return "四号";
                //case 12: return "小四";
                //case 10.5: return "五号";
                //case 9: return "小五";
                //case 7.5: return "六号";
                //case 6.5: return "小六";
                //case 5.5: return "七号";
                //case 5: return "八号"; 
                #endregion
                case 42: return "初号";
                case 36: return "小初";
                case 26.25: return "一号";
                case 24: return "小一";
                case 21.75: return "二号";
                case 18: return "小二";
                case 15.75: return "三号";
                case 15: return "小三";
                case 14.25: return "四号";
                case 12: return "小四";
                case 10.5: return "五号";
                case 9: return "小五";
                case 7.5: return "六号";
                case 6.75: return "小六";
                case 5.25: return "七号";
                default: return "四号";
            }
        }
        /// <summary>
        /// 获取文档中的字号大小
        /// </summary>
        /// <param name="size"></param>
        /// <returns></returns>
        public static double GetRealSize(double size)
        {
            switch (size)
            {
                case 42: return 42;
                case 36: return 36;
                case 26.25: return 26;
                case 24: return 24;
                case 21.75: return 22;
                case 18: return 18;
                case 15.75: return 16;
                case 15: return 15;
                case 14.25: return 14;
                case 12: return 12;
                case 10.5: return 10.5;
                case 9: return 9;
                case 7.5: return 7.5;
                case 6.75: return 6.5;
                case 5.25: return 5.5;
                default:return 14;
            }
        }

        public static ParagraphAlignment GetAlignmentForComboBox(string item)
        {
            switch (item)
            {
                case "居左":
                    return ParagraphAlignment.Left;
                case "居中":
                    return ParagraphAlignment.Center;
                case "居右":
                    return ParagraphAlignment.Right;
                case "左右分散":
                    return ParagraphAlignment.Justify;
                default:
                    return ParagraphAlignment.Left;
            }
        }

        public static void SetParam(ref Aspose.Words.Tables.CellVerticalAlignment cali, ref ParagraphAlignment ali, string item)
        {
            switch (item)
            {
                case "LeftUp": cali = Aspose.Words.Tables.CellVerticalAlignment.Top; ali = ParagraphAlignment.Left; break;
                case "CenterUp": cali = Aspose.Words.Tables.CellVerticalAlignment.Top; ali = ParagraphAlignment.Center; break;
                case "RightUp": cali = Aspose.Words.Tables.CellVerticalAlignment.Top; ali = ParagraphAlignment.Right; break;
                case "LeftMiddle": cali = Aspose.Words.Tables.CellVerticalAlignment.Center; ali = ParagraphAlignment.Left; break;
                case "CenterMiddle": cali = Aspose.Words.Tables.CellVerticalAlignment.Center; ali = ParagraphAlignment.Center; break;
                case "RightMiddle": cali = Aspose.Words.Tables.CellVerticalAlignment.Center; ali = ParagraphAlignment.Right; break;
                case "LeftBottom": cali = Aspose.Words.Tables.CellVerticalAlignment.Bottom; ali = ParagraphAlignment.Left; break;
                case "CenterBottom": cali = Aspose.Words.Tables.CellVerticalAlignment.Bottom; ali = ParagraphAlignment.Center; break;
                case "RightBottom": cali = Aspose.Words.Tables.CellVerticalAlignment.Bottom; ali = ParagraphAlignment.Right; break;
            }
        }

        /// <summary>
        /// 获取边框的宽度
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public static double GetLineWidthForTable(string item)
        {
            switch (item)
            {
                case "0.25磅": return 0.25;
                case "0.5磅": return 0.5;
                case "0.75磅": return 0.75;
                case "1.0磅": return 1.0;
                case "1.5磅": return 1.5;
                case "2.25磅": return 2.25;
                case "3.0磅": return 3.0;
                case "4.5磅": return 4.5;
                case "6.0磅": return 6.0;
                default: return 1.0;
            }
        }

        /// <summary>
        ///     方向参数暂不用Aspose.Words.Orientation pageDirection,
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="pageType"></param>
        /// <param name="topMargin"></param>
        /// <param name="bottomMargin"></param>
        /// <param name="leftMargin"></param>
        /// <param name="rightMargin"></param>
        public static void SetPageSet(Document doc, PaperSize pageType, double topMargin, double bottomMargin, double leftMargin, double rightMargin)
        {
            for (int i = 0; i < doc.Sections.Count; i++)
            {
                PageSetup page = doc.Sections[i].PageSetup;
                //page.DifferentFirstPageHeaderFooter = true;
                //page.PaperSize = pageType;
                //page.Orientation = pageDirection;
                page.TopMargin = topMargin;
                page.BottomMargin = bottomMargin;
                page.LeftMargin = leftMargin;
                page.RightMargin = rightMargin;
            }

        }

        /// <summary>
        /// 添加页眉/页脚
        /// </summary>
        /// <param name="doc"></param>
        public static void AddHeaderFooter(Document doc, string headerTxt, string footerTxt, string path)
        {
            #region 第一版
            //DocumentBuilder builder = new DocumentBuilder(doc);
            //builder.PageSetup.DifferentFirstPageHeaderFooter = true;
            ////builder.PageSetup.OddAndEvenPagesHeaderFooter = true;
            //PageSetup pageSetup = builder.PageSetup;
            //pageSetup.PageStartingNumber = 1;
            //pageSetup.RestartPageNumbering = true;
            //pageSetup.PageNumberStyle = NumberStyle.Arabic;
            //builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
            //builder.Write(headerTxt);
            //builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
            //builder.Write(headerTxt);
            //if (!string.IsNullOrWhiteSpace(path))
            //{
            //    builder.InsertImage(path, 30, 30);//可以控制图片的宽高
            //}
            //builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            //builder.Write(headerTxt);
            //if (!string.IsNullOrWhiteSpace(path))
            //{
            //    builder.InsertImage(path, 30, 30);//可以控制图片的宽高
            //}

            //builder.MoveToHeaderFooter(HeaderFooterType.FooterFirst);
            //builder.Write(footerTxt);
            //builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);
            //builder.Write(footerTxt);
            //builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            //builder.Write(footerTxt); 
            #endregion

            
            var headerfooters = doc.Sections[0].GetChildNodes(NodeType.HeaderFooter, true);
            if (headerfooters.Count > 0) return;

            doc.Sections[0].PageSetup.DifferentFirstPageHeaderFooter = false;
            doc.Sections[0].PageSetup.OddAndEvenPagesHeaderFooter = false;
            HeaderFooter h1 = new HeaderFooter(doc, HeaderFooterType.HeaderFirst);
            Paragraph p_F = new Paragraph(doc);
            Run run1_l = new Run(doc, headerTxt);

            HeaderFooter h2 = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);

            Run run2 = new Run(doc, headerTxt);

            Paragraph p_P = new Paragraph(doc);
            if (!string.IsNullOrWhiteSpace(path))
            {
                Shape shape = new Shape(doc, ShapeType.Image);
                shape.Width = 30;
                shape.Height = 30;
                shape.ImageData.SetImage(path);
                p_P.ChildNodes.Add(shape);
            }

            HeaderFooter h3 = new HeaderFooter(doc, HeaderFooterType.HeaderEven);
            Run run3 = new Run(doc, headerTxt);
            Paragraph p_E = new Paragraph(doc);

            if (!string.IsNullOrWhiteSpace(path))
            {
                Shape shape = new Shape(doc, ShapeType.Image);
                shape.Width = 30;
                shape.Height = 30;
                shape.ImageData.SetImage(path);
                shape.HorizontalAlignment = Aspose.Words.Drawing.HorizontalAlignment.Right;//图片向右对齐
                p_F.ChildNodes.Add(shape);
                p_P.ChildNodes.Add(shape);
                p_E.ChildNodes.Add(shape);
            }
            p_F.Runs.Add(run1_l);
            h1.Paragraphs.Add(p_F);

            p_P.Runs.Add(run2);
            h2.Paragraphs.Add(p_P);

            p_E.Runs.Add(run3);
            h3.Paragraphs.Add(p_E);
            HeaderFooter f1 = new HeaderFooter(doc, HeaderFooterType.FooterFirst);
            Run r1 = new Run(doc, footerTxt);
            Paragraph p1 = new Paragraph(doc);
            p1.Runs.Add(r1);
            f1.Paragraphs.Add(p1);

            HeaderFooter f2 = new HeaderFooter(doc, HeaderFooterType.FooterEven);
            Run r2 = new Run(doc, footerTxt);
            Paragraph p2 = new Paragraph(doc);
            p2.Runs.Add(r2);
            f2.Paragraphs.Add(p2);

            HeaderFooter f3 = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
            Run r3 = new Run(doc, footerTxt);
            Paragraph p3 = new Paragraph(doc);
            p3.InsertField(Aspose.Words.Fields.FieldType.FieldPage, true, null, true);
            p3.Runs.Add(r3);
            
            f3.Paragraphs.Add(p3);
            doc.Sections[0].HeadersFooters.Add(h1);
            doc.Sections[0].HeadersFooters.Add(h2);
            doc.Sections[0].HeadersFooters.Add(h3);
            doc.Sections[0].HeadersFooters.Add(f1);
            doc.Sections[0].HeadersFooters.Add(f2);
            doc.Sections[0].HeadersFooters.Add(f3);

            //HeaderFooter headerFooter = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
            //Paragraph p = new Paragraph(doc);
            //p.InsertField(Aspose.Words.Fields.FieldType.FieldPage, true, null, true);

            //headerFooter.Paragraphs.Add(p);
            //doc.Sections[0].HeadersFooters.Add(headerFooter);









            #region 第二版
            //for (int i = 0; i < doc.Sections.Count; i++)
            //{

            //    Section section = doc.Sections[i];
            //    //List<Node> nodes = section.GetChildNodes(NodeType.FieldStart, true).ToList();
            //    //List<Node> items2=  nodes.FindAll(item => ((Aspose.Words.Fields.FieldStart)item).FieldType == Aspose.Words.Fields.FieldType.FieldPage).ToList();

            //    var headerfooters = section.GetChildNodes(NodeType.HeaderFooter,true);
            //    section.PageSetup.DifferentFirstPageHeaderFooter = false;
            //    section.PageSetup.OddAndEvenPagesHeaderFooter = false;



            //    var a = headerfooters.Count;
            //    var d = section.PageSetup.RestartPageNumbering;  
            //    if (!section.PageSetup.RestartPageNumbering&&i!=0) continue;
            //    if (section.HeadersFooters.Count == 0)
            //    {
            //        //var d = section.PageSetup.RestartPageNumbering;
            //        HeaderFooter h1 = new HeaderFooter(doc, HeaderFooterType.HeaderFirst);
            //        Paragraph p_F = new Paragraph(doc);
            //        Run run1_l = new Run(doc, headerTxt);

            //        HeaderFooter h2 = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);

            //        Run run2 = new Run(doc, headerTxt);

            //        Paragraph p_P = new Paragraph(doc);
            //        if (!string.IsNullOrWhiteSpace(path))
            //        {
            //            Shape shape = new Shape(doc, ShapeType.Image);
            //            shape.Width = 30;
            //            shape.Height = 30;
            //            shape.ImageData.SetImage(path);
            //            p_P.ChildNodes.Add(shape);
            //        }

            //        HeaderFooter h3 = new HeaderFooter(doc, HeaderFooterType.HeaderEven);
            //        Run run3 = new Run(doc, headerTxt);
            //        Paragraph p_E = new Paragraph(doc);

            //        if (!string.IsNullOrWhiteSpace(path))
            //        {
            //            Shape shape = new Shape(doc, ShapeType.Image);
            //            shape.Width = 30;
            //            shape.Height = 30;
            //            shape.ImageData.SetImage(path);
            //            shape.HorizontalAlignment = Aspose.Words.Drawing.HorizontalAlignment.Right;//图片向右对齐
            //            p_F.ChildNodes.Add(shape);
            //            p_P.ChildNodes.Add(shape);
            //            p_E.ChildNodes.Add(shape);
            //        }
            //        p_F.Runs.Add(run1_l);
            //        h1.Paragraphs.Add(p_F);

            //        p_P.Runs.Add(run2);
            //        h2.Paragraphs.Add(p_P);

            //        p_E.Runs.Add(run3);
            //        h3.Paragraphs.Add(p_E);
            //        HeaderFooter f1 = new HeaderFooter(doc, HeaderFooterType.FooterFirst);
            //        Run r1 = new Run(doc, footerTxt);
            //        Paragraph p1 = new Paragraph(doc);
            //        p1.Runs.Add(r1);
            //        f1.Paragraphs.Add(p1);

            //        HeaderFooter f2 = new HeaderFooter(doc, HeaderFooterType.FooterEven);
            //        Run r2 = new Run(doc, footerTxt);
            //        Paragraph p2 = new Paragraph(doc);
            //        p2.Runs.Add(r2);
            //        f2.Paragraphs.Add(p2);

            //        HeaderFooter f3 = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
            //        Run r3 = new Run(doc, footerTxt);
            //        Paragraph p3 = new Paragraph(doc);
            //        p3.Runs.Add(r3);
            //        f3.Paragraphs.Add(p3);

            //        //f1.FirstParagraph.AppendField(Aspose.Words.Fields.FieldType.FieldPage, true);
            //        //f2.FirstParagraph.AppendField(Aspose.Words.Fields.FieldType.FieldPage, true);
            //        //f3.FirstParagraph.AppendField(Aspose.Words.Fields.FieldType.FieldPage, true);
            //        f1.FirstParagraph.InsertField("PAGE", null, true);
            //        f2.FirstParagraph.InsertField("PAGE", null, true);
            //        f3.FirstParagraph.InsertField("PAGE", null, true);
            //        section.HeadersFooters.Add(h1);
            //        section.HeadersFooters.Add(h2);
            //        section.HeadersFooters.Add(h3);
            //        section.HeadersFooters.Add(f1);
            //        section.HeadersFooters.Add(f2);
            //        section.HeadersFooters.Add(f3);
            //    }

            //    //NodeCollection hfs = section.HeadersFooters;
            //    //for (int j = 0; j < hfs.Count; j++)
            //    //{
            //    //    HeaderFooter hf = (HeaderFooter)hfs[j];
            //    //    if (hf.HeaderFooterType == HeaderFooterType.HeaderEven || hf.HeaderFooterType == HeaderFooterType.HeaderFirst || hf.HeaderFooterType == HeaderFooterType.HeaderPrimary)
            //    //        continue;
            //    //    NodeCollection ps = hf.GetChildNodes(NodeType.FieldStart, true);

            //    //    if (ps.Count == 0)
            //    //    {
            //    //        hf.FirstParagraph.InsertField("PAGE", null, true);
            //    //    }
            //    //}


            //    //PageSetup page = doc.Sections[i].PageSetup;
            //    ////page
            //    //page.PageStartingNumber = 1;
            //    //page.RestartPageNumbering = true;
            //    //page.PageNumberStyle = NumberStyle.Arabic;


            //} 
            #endregion

        }

        /// <summary>
        /// 设置页眉页脚样式
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
        public static void SetStyleForHeaderFooter(Document doc, System.Drawing.Font headerFont, Color headerColor, ParagraphAlignment headerAli, System.Drawing.Font footerFont, Color footerColor, ParagraphAlignment footerAli, bool isParityDif, bool isFirstDif, double headerDistance, double footerDistance, double hLeftIndent, double hRightIndent, double fLeftIndent, double fRightIndent, double hLineSpace, double fLineSpace)
        {
            NodeCollection headfooters = doc.GetChildNodes(NodeType.HeaderFooter, true);
            NodeCollection items2 = doc.GetChildNodes(NodeType.Any, true);
            //是否奇偶不同
            if (isParityDif)
            {
                doc.Sections[0].PageSetup.OddAndEvenPagesHeaderFooter = true;
            }
            //是否首页不同
            if (isFirstDif)
            {
                doc.Sections[0].PageSetup.DifferentFirstPageHeaderFooter = true;
            }

            PageSetup page = doc.Sections[0].PageSetup;
            //page
            page.PageStartingNumber = 1;
            page.RestartPageNumbering = true;
            page.PageNumberStyle = NumberStyle.Arabic;

            page.HeaderDistance = headerDistance;//居上距离
            page.FooterDistance = footerDistance;//页脚底部距离
            for (int i = 0; i < headfooters.Count; i++)
            {
                HeaderFooter hf = (HeaderFooter)headfooters[i];
                Color cFont = Color.FromArgb(0, 0, 0, 0);
                System.Drawing.Font f = null;

                switch (hf.HeaderFooterType)
                {
                    case HeaderFooterType.HeaderFirst:
                    case HeaderFooterType.HeaderEven:
                    case HeaderFooterType.HeaderPrimary:
                        f = headerFont;
                        cFont = headerColor;
                        //hf.FirstParagraph.ParagraphFormat.SpaceBefore = 30;//段间距
                        hf.FirstParagraph.ParagraphFormat.LineSpacing = hLineSpace;//行间距
                        hf.FirstParagraph.ParagraphFormat.LeftIndent = hLeftIndent;//左缩进
                        hf.FirstParagraph.ParagraphFormat.RightIndent = hRightIndent;//右缩进
                        SetStyleForHeaderFooterFont(hf, f.Name, f.Size, cFont, headerAli, f.Bold, f.Italic);
                        break;
                    case HeaderFooterType.FooterFirst:
                    case HeaderFooterType.FooterEven:
                        f = footerFont;
                        cFont = footerColor;
                        hf.FirstParagraph.ParagraphFormat.LineSpacing = fLineSpace;//行间距
                        hf.FirstParagraph.ParagraphFormat.LeftIndent = fLeftIndent;//左缩进
                        hf.FirstParagraph.ParagraphFormat.RightIndent = fRightIndent;//右缩进
                        SetStyleForHeaderFooterFont(hf, f.Name, f.Size, cFont, footerAli, f.Bold, f.Italic);
                        //if (isParityDif)
                        //{
                        //    #region 带样式的页码
                        //    //Node node_item = hf.FirstParagraph.FirstChild;
                        //    //Run run1 = new Run(doc, "~");
                        //    //Run run2 = new Run(doc, "~");
                        //    //hf.FirstParagraph.InsertBefore(run1, node_item);
                        //    //hf.FirstParagraph.InsertBefore(run2, node_item);
                        //    //Aspose.Words.Fields.Field field = hf.FirstParagraph.InsertField("PAGE", run1, true); 
                        //    #endregion
                        //    NodeCollection ps = hf.GetChildNodes(NodeType.FieldStart, true);
                        //    if (ps.Count == 0)
                        //    {
                        //        hf.FirstParagraph.InsertField("PAGE", null, true);
                        //    }
                        //}
                        //NodeCollection ps = hf.GetChildNodes(NodeType.FieldStart, true);
                        //if (ps.Count == 0)
                        //{
                        //    hf.FirstParagraph.InsertField("PAGE", null, true);
                        //}

                        break;
                    case HeaderFooterType.FooterPrimary:
                        f = footerFont;
                        cFont = footerColor;
                        hf.FirstParagraph.ParagraphFormat.LineSpacing = fLineSpace;//行间距
                        hf.FirstParagraph.ParagraphFormat.LeftIndent = fLeftIndent;//左缩进
                        hf.FirstParagraph.ParagraphFormat.RightIndent = fRightIndent;//右缩进
                        SetStyleForHeaderFooterFont(hf, f.Name, f.Size, cFont, footerAli, f.Bold, f.Italic);
                        //NodeCollection items = hf.GetChildNodes(NodeType.FieldStart, true);
                        //if (items.Count == 0)
                        //{
                        //    hf.FirstParagraph.InsertField("PAGE", null, true);
                        //}
                        break;
                }
            }
        }

        /// <summary>
        /// 设置表格样式
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="f"></param>
        /// <param name="colorBorder"></param>
        /// <param name="colorFont"></param>
        /// <param name="lineSpace"></param>
        public static void SetStyleForTable(Document doc, System.Drawing.Font f, Color colorBorder, Color colorFont, double lineSpace, double borderWidth, double spaceBefore, double spaceAfter, double leftIndent, double rightIndent, bool isShading, Color shadingColor, LineStyle lineStyle, ParagraphAlignment tableali, Aspose.Words.Tables.CellVerticalAlignment tali)
        {
            NodeCollection nodes = doc.GetChildNodes(NodeType.Table, true);
            if (nodes != null && nodes.Count > 0)
            {
                for (int i = 0; i < nodes.Count; i++)
                {
                    Aspose.Words.Tables.Table table = (Aspose.Words.Tables.Table)nodes[i];
                    for (int a = 0; a < table.Rows.Count; a++)
                    {
                        Aspose.Words.Tables.Row row = table.Rows[a];
                        for (int h = 0; h < row.Cells.Count; h++)
                        {
                            Aspose.Words.Tables.Cell cell = row.Cells[h];
                            //首行加底纹
                            if (isShading)
                            {
                                if (a == 0)
                                {
                                    cell.CellFormat.Shading.BackgroundPatternColor = shadingColor;//底纹颜色
                                }
                            }

                            cell.CellFormat.VerticalAlignment = tali;
                            setStyleForBorder(cell.CellFormat.Borders[BorderType.Top], lineStyle, borderWidth, colorBorder);
                            setStyleForBorder(cell.CellFormat.Borders[BorderType.Bottom], lineStyle, borderWidth, colorBorder);
                            setStyleForBorder(cell.CellFormat.Borders[BorderType.Left], lineStyle, borderWidth, colorBorder);
                            setStyleForBorder(cell.CellFormat.Borders[BorderType.Right], lineStyle, borderWidth, colorBorder);
                            NodeCollection nodes_Cell = cell.GetChildNodes(NodeType.Paragraph, true);
                            int count = nodes_Cell.Count;

                            foreach (Paragraph item in nodes_Cell)
                            {
                                NodeCollection shapes = item.GetChildNodes(NodeType.Shape,true);
                                if (shapes.Count > 0) continue;
                                SetStyleForParagraphFont(item, f.Name, f.Size, colorFont, lineSpace, spaceBefore, spaceAfter, leftIndent, rightIndent, tableali, f.Bold, f.Italic);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 设置段落格式
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="oneT"></param>
        /// <param name="oneC"></param>
        /// <param name="oneLS"></param>
        /// <param name="twoT"></param>
        /// <param name="twoC"></param>
        /// <param name="twoLS"></param>
        /// <param name="threeT"></param>
        /// <param name="threeC"></param>
        /// <param name="threeLS"></param>
        /// <param name="fourT"></param>
        /// <param name="fourC"></param>
        /// <param name="fourLS"></param>
        /// <param name="fiveT"></param>
        /// <param name="fiveC"></param>
        /// <param name="fiveLS"></param>
        /// <param name="contentFont"></param>
        /// <param name="contentC"></param>
        /// <param name="contentLS"></param>
        public static void SetStyleForParagraph(Document doc, System.Drawing.Font oneT, Color oneC, double oneLS, System.Drawing.Font twoT, Color twoC, double twoLS, System.Drawing.Font threeT, Color threeC, double threeLS, System.Drawing.Font fourT, Color fourC, double fourLS, System.Drawing.Font fiveT, Color fiveC, double fiveLS, System.Drawing.Font contentFont, Color contentC, double contentLS, double spaceBeforeOne, double spaceAfterOne, double leftIndentOne, double rightIndentOne, double spaceBeforeTwo, double spaceAfterTwo, double leftIndentTwo, double rightIndentTwo, double spaceBeforeThree, double spaceAfterThree, double leftIndentThree, double rightIndentThree, double spaceBeforeFour, double spaceAfterFour, double leftIndentFour, double rightIndentFour, double spaceBeforeFive, double spaceAfterFive, double leftIndentFive, double rightIndentFive, double spaceBeforeContent, double spaceAfterContent, double leftIndentContent, double rightIndentContent, ParagraphAlignment aliOne, ParagraphAlignment aliTwo, ParagraphAlignment aliThree, ParagraphAlignment aliFour, ParagraphAlignment aliFive, ParagraphAlignment aliContent)
        {
            bool isFirst = true;
            LayoutEnumerator layoutEnumerator = new LayoutEnumerator(doc);
            LayoutCollector layoutCollector = new LayoutCollector(doc);
           
            for (int i = 0; i < doc.Sections.Count; i++)
            {
                //Section section = doc.Sections[i];
                //for (int a = 0; a < section.Body.Paragraphs.Count; a++)
                //{
                //    Paragraph p = section.Body.Paragraphs[a];
                //    if (!isFirst)
                //    {
                //        for (int b = 0; b < p.Runs.Count; b++)
                //        {
                //            Run r = p.Runs[b];
                //            r.Font.Color = Color.Red;
                //        }
                //    }
                //    else if (p.GetText().Contains(ControlChar.PageBreak))//去除首页样式修改
                //    {
                //        isFirst = false;
                //    }

                //}
                Section section = doc.Sections[i];
                ParagraphCollection pcItems = section.Body.Paragraphs;
                for (int a = 0; a < pcItems.Count; a++)
                {
                    Paragraph p = pcItems[a];
                    NodeCollection shapes = p.GetChildNodes(NodeType.Shape, true);
                    if (shapes.Count > 0) continue;
                    if (!isFirst)
                    {
                        if (p.ParentNode.NodeType != NodeType.Body) continue;

                        System.Drawing.Font f = null;
                        Color cFont = Color.FromArgb(0, 0, 0, 0);

                        switch (p.ParagraphFormat.OutlineLevel)
                        {
                            case OutlineLevel.Level1:
                                f = oneT;
                                cFont = oneC;
                                p.ParagraphFormat.Alignment = aliOne;
                                p.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
                                p.ParagraphFormat.LineSpacing = oneLS;
                                SetStyleForParagraphFont(p, f.Name, f.Size, cFont, 0, spaceBeforeOne, spaceAfterOne, leftIndentOne, rightIndentOne, aliOne, f.Bold, f.Italic);
                                break;
                            case OutlineLevel.Level2:
                                f = twoT;
                                cFont = twoC;
                                p.ParagraphFormat.Alignment = aliTwo;
                                p.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
                                p.ParagraphFormat.LineSpacing = twoLS;
                                SetStyleForParagraphFont(p, f.Name, f.Size, cFont, 0, spaceBeforeTwo, spaceAfterTwo, leftIndentTwo, rightIndentTwo, aliTwo, f.Bold, f.Italic);
                                break;
                            case OutlineLevel.Level3:
                                f = threeT;
                                cFont = threeC;
                                p.ParagraphFormat.Alignment = aliThree;
                                p.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
                                p.ParagraphFormat.LineSpacing = threeLS;
                                SetStyleForParagraphFont(p, f.Name, f.Size, cFont, 0, spaceBeforeThree, spaceAfterThree, leftIndentThree, rightIndentThree, aliThree, f.Bold, f.Italic);
                                break;
                            case OutlineLevel.Level4:
                                f = fourT;
                                cFont = fourC;
                                p.ParagraphFormat.Alignment = aliFour;
                                p.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
                                p.ParagraphFormat.LineSpacing = fourLS;
                                SetStyleForParagraphFont(p, f.Name, f.Size, cFont, 0, spaceBeforeFour, spaceAfterFour, leftIndentFour, rightIndentFour, aliFour, f.Bold, f.Italic);
                                break;
                            case OutlineLevel.Level5:
                                f = fiveT;
                                cFont = fiveC;
                                p.ParagraphFormat.Alignment = aliFive;
                                p.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
                                p.ParagraphFormat.LineSpacing = fiveLS;
                                SetStyleForParagraphFont(p, f.Name, f.Size, cFont, 0, spaceBeforeFive, spaceAfterFive, leftIndentFive, rightIndentFive, aliFive, f.Bold, f.Italic);
                                break;
                            case OutlineLevel.BodyText:
                                f = contentFont;
                                cFont = contentC;

                                p.ParagraphFormat.Alignment = aliContent;
                                p.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
                                p.ParagraphFormat.LineSpacing = contentLS;
                                SetStyleForParagraphFont(p, f.Name, f.Size, cFont, 0, spaceBeforeContent, spaceAfterContent, leftIndentContent, rightIndentContent, aliContent, f.Bold, f.Italic);
                                break;
                                //default:
                                //    f = contentFont;
                                //    cFont = contentC;
                                //    p.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
                                //    SetStyleForParagraphFont(p, f.Name, f.Size, cFont, contentLS,f.Bold, f.Italic);
                                //    break;
                        }
                    }
                    else if (p.GetText().Contains(ControlChar.PageBreak))//去除首页样式修改
                    {
                        isFirst = false;
                    }
                }
            }
        }

        /// <summary>
        /// 设置页眉/页脚格式
        /// </summary>
        /// <param name="hf"></param>
        /// <param name="fontName"></param>
        /// <param name="size"></param>
        /// <param name="fontColor"></param>
        /// <param name="alignment"></param>
        /// <param name="isBold"></param>
        /// <param name="Italic"></param>
        public static void SetStyleForHeaderFooterFont(HeaderFooter hf, string fontName, float size, Color fontColor, ParagraphAlignment alignment, bool isBold = false, bool Italic = false)
        {
            try
            {
                foreach (Run item in hf.FirstParagraph.Runs)
                {
                    if (item == null) continue;
                    item.Font.Size = GetRealSize(size);
                    item.Font.Color = fontColor;
                    item.Font.Bold = isBold;
                    item.Font.Italic = Italic;
                    item.Font.Name = fontName;
                    item.ParentParagraph.ParagraphFormat.Alignment = alignment;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 设置边框
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="lineStyle"></param>
        /// <param name="lineWidth"></param>
        /// <param name="color"></param>
        /// <returns></returns>
        public static void setStyleForBorder(Aspose.Words.Border border, LineStyle lineStyle, double lineWidth, Color color)
        {
            border.Color = color;
            border.LineStyle = lineStyle;
            border.LineWidth = lineWidth;
        }

        /// <summary>
        /// 设置内容格式
        /// </summary>
        /// <param name="p"></param>
        /// <param name="fontName"></param>
        /// <param name="size"></param>
        /// <param name="fontColor"></param>
        /// <param name="lineSpace"></param>
        /// <param name="isBold"></param>
        /// <param name="Italic"></param>
        public static void SetStyleForParagraphFont(Paragraph p, string fontName, float size, Color fontColor, double lineSpace, double spaceBefore, double spaceAfter, double leftIndent, double rightIndent, ParagraphAlignment ali, bool isBold = false, bool Italic = false)
        {
            try
            {
                foreach (Run item in p.Runs)
                {
                    if (item == null) continue;
                    item.Font.Size = GetRealSize(size);
                    item.Font.Color = fontColor;
                    item.Font.Bold = isBold;
                    item.Font.Italic = Italic;
                    item.Font.Name = fontName;
                    if (lineSpace > 0)
                    {
                        item.ParentParagraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
                        item.ParentParagraph.ParagraphFormat.LineSpacing = lineSpace;
                        item.ParentParagraph.ParagraphFormat.Alignment = ali;
                    }
                    item.ParentParagraph.ParagraphFormat.SpaceBefore = spaceBefore;
                    item.ParentParagraph.ParagraphFormat.SpaceAfter = spaceAfter;
                    item.ParentParagraph.ParagraphFormat.LeftIndent = leftIndent;
                    item.ParentParagraph.ParagraphFormat.RightIndent = rightIndent;

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

    }
}
