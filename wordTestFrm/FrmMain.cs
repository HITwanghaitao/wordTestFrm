using Aspose.Words;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using wordTestFrm.Common;
using wordTestFrm.models;
using wordTestFrm.Properties;

namespace wordTestFrm
{
    public partial class FrmMain : Form
    {
        Thread th = null;
        Thread wordApply = null;
        //行距常量
        public double lineSpaceConstant = 12;
        //缩进常量
        public double indentConstant = 10.5;
        /// <summary>
        /// 段前段后
        /// </summary>
        public double spaceBeforeConstant = 15.6;
        public string logoPath = string.Empty;
        //线样式图片地址
        string lineStylePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "lineStyleImgs");
        public FrmMain()
        {
            InitializeComponent();
            //Control.CheckForIllegalCrossThreadCalls = false;
            cbxDPI.SelectedIndex = 1;
            cmb_header.SelectedIndex = 0;
            cmb_footer.SelectedIndex = 0;
            cb_hLineSpace.SelectedIndex = 0;
            cb_fLineSpace.SelectedIndex = 0;
            cb_tableLineWidth.SelectedIndex = 0;
            cb_tableSpaceBefore.SelectedIndex = 0;
            cb_tableSpaceAfter.SelectedIndex = 0;
            cb_tableLineSpace.SelectedIndex = 0;
            cb_LineSpaceOne.SelectedIndex = 0;
            cb_LineSpaceTwo.SelectedIndex = 0;
            cb_LineSpaceThree.SelectedIndex = 0;
            cb_LineSpaceFour.SelectedIndex = 0;
            cb_LineSpaceFive.SelectedIndex = 0;
            cb_LineSpaceContent.SelectedIndex = 0;
            cb_spaceBeforeOne.SelectedIndex = 0;
            cb_spaceBeforeTwo.SelectedIndex = 0;
            cb_spaceBeforeThree.SelectedIndex = 0;
            cb_spaceBeforeFour.SelectedIndex = 0;
            cb_spaceBeforeFive.SelectedIndex = 0;
            cb_spaceBeforeContent.SelectedIndex = 0;
            cb_spaceAfterOne.SelectedIndex = 0;
            cb_spaceAfterTwo.SelectedIndex = 0;
            cb_spaceAfterThree.SelectedIndex = 0;
            cb_spaceAfterFour.SelectedIndex = 0;
            cb_spaceAfterFive.SelectedIndex = 0;
            cb_spaceAfterContent.SelectedIndex = 0;
            cb_headerFooter.Checked = true;
            cb_table.Checked = true;
            cb_content.Checked = true;
            cb_pageSet.Checked = true;
            cb_locationOne.SelectedIndex = 0;
            cb_locationTwo.SelectedIndex = 0;
            cb_locationThree.SelectedIndex = 0;
            cb_locationFour.SelectedIndex = 0;
            cb_locationFive.SelectedIndex = 0;
            cb_locationContent.SelectedIndex = 0;
            cb_pageType.SelectedIndex = 1;
            cb_pageDirection.SelectedIndex = 0;
            //cb_locationTable.SelectedIndex = 0;
            panel1.Visible = false;
            this.ucCellAlignment1.LabelClick += UcCellAlignment1_LabelClick;

            uc_footerFont.lblOther = lblContent;
            uc_headerFont.lblOther = lblContent;
            uc_tableFont.lblOther = lblContent;
            uc_oneFont.lblOther = lblContent;
            uc_twoFont.lblOther = lblContent;
            uc_threeFont.lblOther = lblContent;
            uc_fourFont.lblOther = lblContent;
            uc_fiveFont.lblOther = lblContent;
            uc_contentFont.lblOther = lblContent;
        }

        private void btn_PDFToImg_Click(object sender, EventArgs e)
        {
            FrmPDFToImgs pdf = new FrmPDFToImgs();
            pdf.ShowDialog();
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

        }

        OpenFileDialog openfile;
        private void btn_selectFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.InitialDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            if (of.ShowDialog() == DialogResult.OK)
            {
                tb_OpenPath.Text = of.FileName;
                openfile = of;
            }
        }
        public int desiredPpi = 150;

        // In .NET this seems to be a good compression / quality setting.
        public int jpegQuality = 90;
        public double width = 15;
        private void btn_startCompression_Click(object sender, EventArgs e)
        {
            if (openfile == null)
            {
                MessageBox.Show("请先选择文件", "提示");
                return;
            }

            string realName = Path.GetFileName(openfile.FileName);
            string CopyPath = CommonMethods.retSavePath(openfile.FileName);
            try
            {
                this.btn_selectFile.Enabled = false;
                this.btn_startCompression.Enabled = false;
                this.ucLoading1.Visible = true;
                //Application.DoEvents();

                this.desiredPpi = int.Parse(cbxDPI.SelectedItem.ToString());
                this.jpegQuality = int.Parse(txtQuality.Text);
                this.width = Convert.ToDouble(tb_imageWidth.Text);
                th = new Thread(() =>
                {
                    try
                    {
                        Aspose.Words.Document doc = new Aspose.Words.Document(CopyPath);
                        int res = Resampler.SetStyleForImage(doc, this.desiredPpi, this.jpegQuality, width);
                        string savePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Path.GetFileNameWithoutExtension(CopyPath) + "_AsposeWord.docx");
                        doc.Save(savePath, Aspose.Words.SaveFormat.Docx);
                        this.Invoke(new Action(() =>
                        {
                            this.ucLoading1.Visible = false;
                            this.btn_selectFile.Enabled = true;
                            this.btn_startCompression.Enabled = true;
                            if (res >= 0)
                            {
                                tb_ResultPath.Text = savePath;
                                Process.Start(savePath);
                            }
                            else
                            {
                                MessageBox.Show("无法解析当前格式文件");
                            }
                        }));
                    }
                    catch (Aspose.Words.FileCorruptedException ex)
                    {
                        MessageBox.Show("文件被占用，请先关闭打开的文件");
                    }
                    catch (Aspose.Words.UnsupportedFileFormatException ex)
                    {
                        MessageBox.Show("无法解析当前格式文件");
                    }
                    catch (Exception ex)
                    {

                    }
                });
                th.Start();
            }
            catch (Aspose.Words.FileCorruptedException ex)
            {
                MessageBox.Show("文件被占用，请先关闭打开的文件");
            }
            catch (Aspose.Words.UnsupportedFileFormatException ex)
            {
                MessageBox.Show("无法解析当前格式文件");
            }
            catch (Exception ex)
            {

            }

        }

        public List<string> fileNames = new List<string>();

        private void FrmMain_Load(object sender, EventArgs e)
        {
            //获取版本号
            string version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
            this.Text = "内部文档辅助工具  v" + version;

            #region 第一版
            //string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WordStyles");
            //if (!Directory.Exists(path)) Directory.CreateDirectory(path);
            //var files = Directory.GetFiles(path, "*.txt", SearchOption.AllDirectories);
            //for (int i = 0; i < files.Length; i++)
            //{
            //    var ss = Path.GetFileName(files[i]);
            //    fileNames.Add(ss);
            //    cmb_Company.Items.Add(Path.GetFileNameWithoutExtension(ss));
            //}
            #endregion
            while (cmb_Company.Items.Count > 0)
            {
                cmb_Company.Items.RemoveAt(0);
            }

            List<string> result = new List<string>();
            string nUrl = "https://" + ConfigurationManager.AppSettings["Ip"] + "/handleWord/getNames";
            string res = CommonTool.httpApi(nUrl, "", "get");
            if (res != "-1")
            {
                result = JsonConvert.DeserializeObject<List<string>>(res);
            }

            for (int i = 0; i < result.Count; i++)
            {
                cmb_Company.Items.Add(result[i]);
            }


            cmb_Company.SelectedIndex = -1;


            images.Add(Resources.single);
            images.Add(Resources.DashLargeGap);
            images.Add(Resources.dot);
            images.Add(Resources.DotDash);
            images.Add(Resources.DotDotDash);
            images.Add(Resources._double);
            images.Add(Resources.Triple);
            images.Add(Resources.ThinThickSmallGap);

            cb_tableLineStyle.Items.Add("1");
            cb_tableLineStyle.Items.Add("7");
            cb_tableLineStyle.Items.Add("6");
            cb_tableLineStyle.Items.Add("8");
            cb_tableLineStyle.Items.Add("9");
            cb_tableLineStyle.Items.Add("3");
            cb_tableLineStyle.Items.Add("10");
            cb_tableLineStyle.Items.Add("11");
            cb_tableLineStyle.SelectedIndex = 0;


            Color color = Color.Black;
            System.Drawing.Font font_def = new System.Drawing.Font("微软雅黑", 14.25f);

            lb_tableShadingColor.Tag = lblTableBorderColor.Tag = lb_tableShadingColor.Tag
          = lblLV1FlagColor.Tag = lblLV2FlagColor.Tag = lblLV3FlagColor.Tag = lblLV4FlagColor.Tag = lblLV5FlagColor.Tag = color;

            lb_tableShadingColor.BackColor = lblTableBorderColor.BackColor = lb_tableShadingColor.BackColor
          = lblLV1FlagColor.BackColor = lblLV2FlagColor.BackColor = lblLV3FlagColor.BackColor = lblLV4FlagColor.BackColor = lblLV5FlagColor.BackColor = color;

            lblLv1FlagFont.Tag = lblLv2FlagFont.Tag = lblLv3FlagFont.Tag = lblLv4FlagFont.Tag = lblLv5FlagFont.Tag
               = font_def;
            cb_hIsFirstDif.Checked = false;
            cb_hIsParityDif.Checked = false;
            cb_fIsFirstDif.Checked = false;
            cb_tableIsShading.Checked = false;

            this.ucCellAlignment1.Visible = false;
            this.cb_locationTable.Text = "CenterMiddle";
            //string path = @"G:\C#\wordTool\wordTestFrm\wordTestFrm\bin\Debug\WordStyles\诚明.txt";
            //StreamReader sr = new StreamReader(path, Encoding.UTF8);
            //string result = string.Empty;
            //string line;
            //while ((line = sr.ReadLine()) != null)
            //{
            //    result += line;
            //}
            //WordStyle ws = JsonConvert.DeserializeObject<WordStyle>(result);
            //this.lb_HeaderTxt.Text = ws.HeaderName;
            //this.lb_FooterTxt.Text = ws.FooterName;
            //this.lblHeaderFont.Tag = ws.HeaderFont;
            //this.lblHeaderFont.Text = string.Format("字体:{0}  字号: {1}", ws.HeaderFont.Name, ws.HeaderFont.Size);
            //this.lblHeaderColor.Tag = ws.HeaderColor;
            //this.lblHeaderColor.BackColor = ws.HeaderColor;
        }

        private void btn_apply_Click(object sender, EventArgs e)
        {
            if (openFile == null)
            {
                MessageBox.Show("请先选择文档");
                return;
            }
            string fileName = openFile.FileName;
            string realName = Path.GetFileName(fileName);

            string filePath = openFile.FileName;
            try
            {
                this.ucLoading1.Visible = true;
                this.btn_apply.Enabled = false;
                this.ucLoading1.Percentage = 0;
                WordStyle ws = new WordStyle();


                #region 主界面控件参数
                PaperSize pageType = CommonMethods.GetPageSizeForComboBox(cb_pageType.SelectedItem.ToString());
                //Aspose.Words.Orientation pageDirection =CommonMethods.GetPageDirectionForComboBox(cb_pageDirection.SelectedItem.ToString());
                double topMargin = ConvertUtil.MillimeterToPoint(Convert.ToDouble(tb_pageMarginUp.Text) * 10);
                double bottomMargin = ConvertUtil.MillimeterToPoint(Convert.ToDouble(tb_pageMarginDown.Text) * 10);
                double leftMargin = ConvertUtil.MillimeterToPoint(Convert.ToDouble(tb_pageMarginLeft.Text) * 10);
                double rightMargin = ConvertUtil.MillimeterToPoint(Convert.ToDouble(tb_pageMarginRight.Text) * 10);

                bool isHeaderFooter = cb_headerFooter.Checked;
                bool isTable = cb_table.Checked;
                bool isContent = cb_content.Checked;
                bool isPageSet = cb_pageSet.Checked;

                string headerName = tb_HeaderTxt.Text;
                string footerName = tb_FooterTxt.Text;
                System.Drawing.Font headerFont = this.uc_headerFont.fontSelect;
                Color headerColor = this.uc_headerFont.fontColorSelect;
                System.Drawing.Font footerFont = this.uc_footerFont.fontSelect;
                Color footerColor = this.uc_footerFont.fontColorSelect;
                bool fIsParityDif = cb_fIsparityDif.Checked;
                bool fIsFirstDif = cb_fIsFirstDif.Checked;
                double headerDistance = ConvertUtil.MillimeterToPoint(Convert.ToDouble(tb_headerDistance.Text) * 10);
                double footerDistance = ConvertUtil.MillimeterToPoint(Convert.ToDouble(tb_FooterDistance.Text) * 10);
                double hLeftIndent = ConvertUtil.MillimeterToPoint(Convert.ToDouble(tb_hLeftIndent.Text) * 10);
                double hRightIndent = Convert.ToDouble(tb_hRightIndent.Text) * indentConstant;
                double fLeftIndent = Convert.ToDouble(tb_fLeftIndent.Text) * indentConstant;
                double fRightIndent = Convert.ToDouble(tb_fRightIndent.Text) * indentConstant;
                double hLineSpace = CommonMethods.GetLineSpaceForComboBox(cb_hLineSpace.SelectedItem.ToString()) * lineSpaceConstant;
                double fLineSpace = CommonMethods.GetLineSpaceForComboBox(cb_fLineSpace.SelectedItem.ToString()) * lineSpaceConstant;

                System.Drawing.Font tableFont = this.uc_tableFont.fontSelect;
                Color tableBorderColor = (Color)lblTableBorderColor.Tag;
                Color tableFontColor = this.uc_tableFont.fontColorSelect;
                double tableLineSpace = CommonMethods.GetLineSpaceForComboBox(cb_tableLineSpace.SelectedItem.ToString()) * lineSpaceConstant;
                double tableLineWidth = CommonMethods.GetLineWidthForTable(cb_tableLineWidth.SelectedItem.ToString());
                double tableSpaceBefore = CommonMethods.GetSpaceBeforeOrAfterForComboBox(cb_tableSpaceBefore.SelectedItem.ToString()) * spaceBeforeConstant;
                double tablespaceAfter = CommonMethods.GetSpaceBeforeOrAfterForComboBox(cb_tableSpaceAfter.SelectedItem.ToString()) * spaceBeforeConstant;
                double tableLeftIndent = Convert.ToDouble(tb_tableLeftIndent.Text);
                double tableRightIndent = Convert.ToDouble(tb_tableRightIndent.Text);
                bool tableIsShading = cb_tableIsShading.Checked;
                Color tableShadingColor = (Color)lb_tableShadingColor.Tag;
                LineStyle tableLineStyle = CommonMethods.GetLineStyleForComboBox(cb_tableLineStyle.SelectedItem.ToString());


                System.Drawing.Font titleOneFont = this.uc_oneFont.fontSelect;
                Color titleOneColor = this.uc_oneFont.fontColorSelect;
                double lineSpaceOne = CommonMethods.GetLineSpaceForComboBox(cb_LineSpaceOne.SelectedItem.ToString()) * lineSpaceConstant;
                System.Drawing.Font titleTwoFont = this.uc_twoFont.fontSelect;
                Color titleTwoColor = this.uc_twoFont.fontColorSelect;
                double lineSpaceTwo = CommonMethods.GetLineSpaceForComboBox(cb_LineSpaceTwo.SelectedItem.ToString()) * lineSpaceConstant;
                System.Drawing.Font titleThreeFont = this.uc_threeFont.fontSelect;
                Color titleThreeColor = this.uc_threeFont.fontColorSelect;
                double lineSpaceThree = CommonMethods.GetLineSpaceForComboBox(cb_LineSpaceThree.SelectedItem.ToString()) * lineSpaceConstant;
                System.Drawing.Font titleFourFont = this.uc_fourFont.fontSelect;
                Color titleFourColor = this.uc_fourFont.fontColorSelect;
                double lineSpaceFour = CommonMethods.GetLineSpaceForComboBox(cb_LineSpaceFour.SelectedItem.ToString()) * lineSpaceConstant;
                System.Drawing.Font titleFiveFont = this.uc_fiveFont.fontSelect;
                Color titleFiveColor = this.uc_fiveFont.fontColorSelect;
                double lineSpaceFive = CommonMethods.GetLineSpaceForComboBox(cb_LineSpaceFive.SelectedItem.ToString()) * lineSpaceConstant;
                System.Drawing.Font contentFont = this.uc_contentFont.fontSelect;
                Color contentColor = this.uc_contentFont.fontColorSelect;
                double lineSpaceContent = CommonMethods.GetLineSpaceForComboBox(cb_LineSpaceContent.SelectedItem.ToString()) * lineSpaceConstant;
                double spaceBeforeOne = CommonMethods.GetSpaceBeforeOrAfterForComboBox(cb_spaceBeforeOne.SelectedItem.ToString()) * spaceBeforeConstant;
                double spaceAfterOne = CommonMethods.GetSpaceBeforeOrAfterForComboBox(cb_spaceAfterOne.SelectedItem.ToString()) * spaceBeforeConstant;
                double leftIndentOne = Convert.ToDouble(tb_leftIndentOne.Text) * indentConstant;
                double rightIndentOne = Convert.ToDouble(tb_rightIndentOne.Text) * indentConstant;
                double spaceBeforeTwo = CommonMethods.GetSpaceBeforeOrAfterForComboBox(cb_spaceBeforeTwo.SelectedItem.ToString()) * spaceBeforeConstant;
                double spaceAfterTwo = CommonMethods.GetSpaceBeforeOrAfterForComboBox(cb_spaceAfterTwo.SelectedItem.ToString()) * spaceBeforeConstant;
                double leftIndentTwo = Convert.ToDouble(tb_leftIndentOne.Text) * indentConstant;
                double rightIndentTwo = Convert.ToDouble(tb_rightIndentTwo.Text) * indentConstant;
                double spaceBeforeThree = CommonMethods.GetSpaceBeforeOrAfterForComboBox(cb_spaceBeforeThree.SelectedItem.ToString()) * spaceBeforeConstant;
                double spaceAfterThree = CommonMethods.GetSpaceBeforeOrAfterForComboBox(cb_spaceAfterThree.SelectedItem.ToString()) * spaceBeforeConstant;
                double leftIndentThree = Convert.ToDouble(tb_leftIndentThree.Text) * indentConstant;
                double rightIndentThree = Convert.ToDouble(tb_rightIndentThree.Text) * indentConstant;
                double spaceBeforeFour = CommonMethods.GetSpaceBeforeOrAfterForComboBox(cb_spaceBeforeFour.SelectedItem.ToString()) * spaceBeforeConstant;
                double spaceAfterFour = CommonMethods.GetSpaceBeforeOrAfterForComboBox(cb_spaceAfterFour.SelectedItem.ToString()) * spaceBeforeConstant;
                double leftIndentFour = Convert.ToDouble(tb_leftIndentFour.Text) * indentConstant;
                double rightIndentFour = Convert.ToDouble(tb_rightIndentFour.Text) * indentConstant;
                double spaceBeforeFive = CommonMethods.GetSpaceBeforeOrAfterForComboBox(cb_spaceBeforeFive.SelectedItem.ToString()) * spaceBeforeConstant;
                double spaceAfterFive = CommonMethods.GetSpaceBeforeOrAfterForComboBox(cb_spaceAfterFive.SelectedItem.ToString()) * spaceBeforeConstant;
                double leftIndentFive = Convert.ToDouble(tb_leftIndentFive.Text) * indentConstant;
                double rightIndentFive = Convert.ToDouble(tb_rightIndentFive.Text) * indentConstant;
                double spaceBeforeContent = CommonMethods.GetSpaceBeforeOrAfterForComboBox(cb_spaceBeforeContent.SelectedItem.ToString()) * spaceBeforeConstant;
                double spaceAfterContent = CommonMethods.GetSpaceBeforeOrAfterForComboBox(cb_spaceAfterContent.SelectedItem.ToString()) * spaceBeforeConstant;
                double leftIndentContent = Convert.ToDouble(tb_leftIndentContent.Text) * indentConstant;
                double rightIndentContent = Convert.ToDouble(tb_rightIndentContent.Text) * indentConstant;
                ParagraphAlignment aliOne = CommonMethods.GetAlignmentForComboBox(cb_locationOne.SelectedItem.ToString());
                ParagraphAlignment aliTwo = CommonMethods.GetAlignmentForComboBox(cb_locationTwo.SelectedItem.ToString());
                ParagraphAlignment aliThree = CommonMethods.GetAlignmentForComboBox(cb_locationThree.SelectedItem.ToString());
                ParagraphAlignment aliFour = CommonMethods.GetAlignmentForComboBox(cb_locationFour.SelectedItem.ToString());
                ParagraphAlignment aliFive = CommonMethods.GetAlignmentForComboBox(cb_locationFive.SelectedItem.ToString());
                ParagraphAlignment aliContent = CommonMethods.GetAlignmentForComboBox(cb_locationContent.SelectedItem.ToString());
                ParagraphAlignment headerAli = CommonMethods.GetAlignmentForComboBox(cmb_header.SelectedItem.ToString());
                ParagraphAlignment footerAli = CommonMethods.GetAlignmentForComboBox(cmb_footer.SelectedItem.ToString());
                ParagraphAlignment tableAli = ParagraphAlignment.Center;
                Aspose.Words.Tables.CellVerticalAlignment tableva = Aspose.Words.Tables.CellVerticalAlignment.Center;
                CommonMethods.SetParam(ref tableva, ref tableAli, cb_locationTable.Text);
                #endregion
                wordApply = new Thread(() =>
                {
                    try
                    {
                        string CopyPath = Common.CommonTool.retSavePath(fileName);
                        Document doc = new Document(CopyPath);
                        DocumentBuilder builder = new DocumentBuilder(doc);

                        #region 应用文档样式
                        if (isPageSet)
                        {

                            //页面布局
                            //CommonMethods.SetPageSet(doc, pageType, pageDirection, topMargin, bottomMargin, leftMargin, rightMargin);
                            CommonMethods.SetPageSet(doc, pageType, topMargin, bottomMargin, leftMargin, rightMargin);
                        }
                        this.ucLoading1.Percentage = 22;
                        if (isHeaderFooter)
                        {
                            NodeCollection nodes = doc.GetChildNodes(NodeType.HeaderFooter, true);
                            //添加页眉/页脚
                            CommonMethods.AddHeaderFooter(doc, headerName, footerName, logoPath);
                        }
                        this.ucLoading1.Percentage = 44;
                        if (isTable)
                        {
                            //设置表格
                            CommonMethods.SetStyleForTable(doc, tableFont, tableBorderColor, tableFontColor, tableLineSpace, tableLineWidth, tableSpaceBefore, tablespaceAfter, tableLeftIndent, tableRightIndent, tableIsShading, tableShadingColor, tableLineStyle, tableAli, tableva);
                        }
                        this.ucLoading1.Percentage = 66;
                        if (isContent)
                        {
                            //设置内容样式
                            CommonMethods.SetStyleForParagraph(doc, titleOneFont, titleOneColor, lineSpaceOne, titleTwoFont, titleTwoColor, lineSpaceTwo, titleThreeFont, titleThreeColor, lineSpaceThree, titleFourFont, titleFourColor, lineSpaceFour, titleFiveFont, titleFiveColor, lineSpaceFive, contentFont, contentColor, lineSpaceContent, spaceBeforeOne, spaceAfterOne, leftIndentOne, rightIndentOne, spaceBeforeTwo, spaceAfterTwo, leftIndentTwo, rightIndentTwo, spaceBeforeThree, spaceAfterThree, leftIndentThree, rightIndentThree, spaceBeforeFour, spaceAfterFour, leftIndentFour, rightIndentFour, spaceBeforeFive, spaceAfterFive, leftIndentFive, rightIndentFive, spaceBeforeContent, spaceAfterContent, leftIndentContent, rightIndentContent, aliOne, aliTwo, aliThree, aliFour, aliFive, aliContent);
                        }
                        this.ucLoading1.Percentage = 88;

                        if (isHeaderFooter)
                        {
                            //设置页眉 / 页脚格式
                            CommonMethods.SetStyleForHeaderFooter(doc, headerFont, headerColor, headerAli, footerFont, footerColor, footerAli, fIsParityDif, fIsFirstDif, headerDistance, footerDistance, hLeftIndent, hRightIndent, fLeftIndent, fRightIndent, hLineSpace, fLineSpace);
                        }
                        this.ucLoading1.Percentage = 100;
                        #endregion
                        string rootDir = Path.GetDirectoryName(filePath);
                        string savePath = Common.CommonTool.retSaveFilePath(
                    Path.Combine(rootDir, Path.GetFileNameWithoutExtension(realName)),
                    realName, ""
                    );
                        doc.Save(savePath, SaveFormat.Docx);
                        this.Invoke(new Action(() =>
                        {
                            this.ucLoading1.Visible = false;
                            this.btn_apply.Enabled = true;
                            Process.Start(savePath);
                        }));
                    }
                    catch (Exception ex)
                    {

                        this.Invoke(new Action(() =>
                        {
                            this.ucLoading1.Visible = false;
                            this.btn_apply.Enabled = true;

                        }));
                    }
                });
                wordApply.Start();
            }
            catch (Aspose.Words.FileCorruptedException ex)
            {
                MessageBox.Show("文件被占用，请先关闭打开的文件");
                this.ucLoading1.Visible = false;
                this.btn_apply.Enabled = true;
            }
            catch (Aspose.Words.UnsupportedFileFormatException ex)
            {
                MessageBox.Show("无法解析当前格式文件");
                this.ucLoading1.Visible = false;
                this.btn_apply.Enabled = true;
            }
            catch (Exception ex)
            {
                this.ucLoading1.Visible = false;
                this.btn_apply.Enabled = true;
            }

        }
        OpenFileDialog openFile = null;
        private void btn_OpenFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.InitialDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            if (of.ShowDialog() == DialogResult.OK)
            {
                openFile = of;
                tb_FilePath.Text = of.FileName;
            }
        }

        private void cmb_Company_SelectedIndexChanged(object sender, EventArgs e)
        {
            #region 第一版
            //string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WordStyles");
            //path = Path.Combine(path, cmb_Company.SelectedItem.ToString() + ".txt");
            //StreamReader sr = new StreamReader(path, Encoding.UTF8);
            //string result = string.Empty;
            //string line;
            //while ((line = sr.ReadLine()) != null)
            //{
            //    result += line;
            //}
            //sr.Close(); 
            #endregion
            try
            {
                if (cmb_Company.SelectedItem == null || string.IsNullOrWhiteSpace(cmb_Company.SelectedItem.ToString())) return;
                this.tb_styleName.Text = cmb_Company.SelectedItem.ToString();
                WordStyle ws = new WordStyle();
                string url = "https://" + ConfigurationManager.AppSettings["Ip"] + "/handleWord/getContent";

                string param = "{\"name\":\"" + cmb_Company.SelectedItem.ToString() + "\"}";
                var result = CommonTool.httpApi(url, param);
                if (result != "-1")
                {
                    ws = JsonConvert.DeserializeObject<WordStyle>(result);
                }

                #region 页眉
                this.tb_HeaderTxt.Text = ws.HeaderName;
                this.uc_headerFont.fontSelect = ws.HeaderFont;
                this.cmb_header.SelectedIndex = (int)ws.HeaderAlignment;
                this.uc_headerFont.fontColorSelect = ws.HeaderColor;
                this.uc_headerFont.SettingsControl(ws.HeaderFont, ws.HeaderColor);
                //this.uc_headerFont.ForeColor = ws.HeaderColor;
                this.tb_hLeftIndent.Text = ws.HLeftIndent.ToString();
                this.tb_hRightIndent.Text = ws.HRightIndent.ToString();
                this.tb_headerDistance.Text = ws.HeaderDistance.ToString();
                if (File.Exists(ws.HImgPath))
                {
                    this.lb_Logo.Image = new Bitmap(ws.HImgPath);
                    logoPath = ws.HImgPath;
                }

                switch (ws.HLineSpace)
                {
                    case 1:
                        this.cb_hLineSpace.SelectedIndex = 0;
                        break;
                    case 1.5:
                        this.cb_hLineSpace.SelectedIndex = 1;
                        break;
                    case 2:
                        this.cb_hLineSpace.SelectedIndex = 2;
                        break;
                    case 3:
                        this.cb_hLineSpace.SelectedIndex = 5;
                        break;
                    default:
                        this.cb_hLineSpace.SelectedIndex = 0;
                        break;
                }
                this.cb_hIsFirstDif.Checked = ws.HIsFirstDif;
                this.cb_hIsParityDif.Checked = ws.HIsParityDif;
                if (!string.IsNullOrWhiteSpace(ws.HImgPath))
                {
                    if (File.Exists(ws.HImgPath))
                        this.lb_Logo.Image = new Bitmap(ws.HImgPath);
                }

                #endregion

                #region 页脚
                this.tb_FooterTxt.Text = ws.FooterName;
                this.uc_footerFont.fontSelect = ws.FooterFont;
                this.uc_footerFont.fontColorSelect = ws.FooterColor;
                this.uc_footerFont.ForeColor = ws.FooterColor;
                this.uc_footerFont.SettingsControl(ws.FooterFont, ws.FooterColor);
                this.cmb_footer.SelectedIndex = (int)ws.FooterAlignment;
                this.tb_fLeftIndent.Text = ws.FLeftIndent.ToString();
                this.tb_fRightIndent.Text = ws.FRightIndent.ToString();
                this.tb_FooterDistance.Text = ws.FooterDistance.ToString();
                switch (ws.FLineSpace)
                {
                    case 1:
                        this.cb_fLineSpace.SelectedIndex = 0;
                        break;
                    case 1.5:
                        this.cb_fLineSpace.SelectedIndex = 1;
                        break;
                    case 2:
                        this.cb_fLineSpace.SelectedIndex = 2;
                        break;
                    case 3:
                        this.cb_fLineSpace.SelectedIndex = 5;
                        break;
                    default:
                        this.cb_fLineSpace.SelectedIndex = 0;
                        break;
                }
                this.cb_fIsFirstDif.Checked = ws.FIsFirstDif;
                this.cb_fIsparityDif.Checked = ws.FIsParityDif;
                #endregion

                #region 表格
                this.uc_tableFont.fontSelect = ws.TableFont;
                this.uc_tableFont.fontColorSelect = ws.TableFontColor;
                //this.uc_tableFont.ForeColor = ws.TableFontColor;
                this.uc_tableFont.SettingsControl(ws.TableFont, ws.TableFontColor);
                this.lblTableBorderColor.Tag = ws.TableBorderColor;
                this.lblTableBorderColor.BackColor = ws.TableBorderColor;
                switch (ws.TableLineWidth)
                {
                    case 0.25: this.cb_tableLineWidth.SelectedIndex = 0; break;
                    case 0.5: this.cb_tableLineWidth.SelectedIndex = 1; break;
                    case 0.75: this.cb_tableLineWidth.SelectedIndex = 2; break;
                    case 1.0: this.cb_tableLineWidth.SelectedIndex = 3; break;
                    case 1.5: this.cb_tableLineWidth.SelectedIndex = 4; break;
                    case 2.25: this.cb_tableLineWidth.SelectedIndex = 5; break;
                    case 3.0: this.cb_tableLineWidth.SelectedIndex = 6; break;
                    case 4.5: this.cb_tableLineWidth.SelectedIndex = 7; break;
                    case 6.0: this.cb_tableLineWidth.SelectedIndex = 8; break;
                    default: this.cb_tableLineWidth.SelectedIndex = 0; break;
                }
                switch (ws.TableSpaceBefore)
                {
                    case 0: this.cb_tableSpaceBefore.SelectedIndex = 0; break;
                    case 0.5: this.cb_tableSpaceBefore.SelectedIndex = 1; break;
                    case 1: this.cb_tableSpaceBefore.SelectedIndex = 2; break;
                    case 1.5: this.cb_tableSpaceBefore.SelectedIndex = 3; break;
                    case 2: this.cb_tableSpaceBefore.SelectedIndex = 4; break;
                    case 2.5: this.cb_tableSpaceBefore.SelectedIndex = 5; break;
                    case 3: this.cb_tableSpaceBefore.SelectedIndex = 6; break;
                    case 3.5: this.cb_tableSpaceBefore.SelectedIndex = 7; break;
                    case 4: this.cb_tableSpaceBefore.SelectedIndex = 8; break;
                    default: this.cb_tableSpaceBefore.SelectedIndex = 0; break;
                }
                switch (ws.TableSpaceAfter)
                {
                    case 0: this.cb_tableSpaceAfter.SelectedIndex = 0; break;
                    case 0.5: this.cb_tableSpaceAfter.SelectedIndex = 1; break;
                    case 1: this.cb_tableSpaceAfter.SelectedIndex = 2; break;
                    case 1.5: this.cb_tableSpaceAfter.SelectedIndex = 3; break;
                    case 2: this.cb_tableSpaceAfter.SelectedIndex = 4; break;
                    case 2.5: this.cb_tableSpaceAfter.SelectedIndex = 5; break;
                    case 3: this.cb_tableSpaceAfter.SelectedIndex = 6; break;
                    case 3.5: this.cb_tableSpaceAfter.SelectedIndex = 7; break;
                    case 4: this.cb_tableSpaceAfter.SelectedIndex = 8; break;
                    default: this.cb_tableSpaceAfter.SelectedIndex = 0; break;
                }
                switch (ws.TableLineSpace)
                {
                    case 1: this.cb_tableLineSpace.SelectedIndex = 0; break;
                    case 1.5: this.cb_tableLineSpace.SelectedIndex = 1; break;
                    case 2: this.cb_tableLineSpace.SelectedIndex = 2; break;
                    case 3: this.cb_tableLineSpace.SelectedIndex = 5; break;
                    default: this.cb_tableLineSpace.SelectedIndex = 0; break;
                }
                if (cb_tableLineStyle.Items.Count != 0)
                {
                    switch (ws.TableLineStyle)
                    {
                        case LineStyle.Single: this.cb_tableLineStyle.SelectedIndex = 0; break;
                        case LineStyle.DashLargeGap: this.cb_tableLineStyle.SelectedIndex = 1; break;
                        case LineStyle.Dot: this.cb_tableLineStyle.SelectedIndex = 2; break;
                        case LineStyle.DotDash: this.cb_tableLineStyle.SelectedIndex = 3; break;
                        case LineStyle.DotDotDash: this.cb_tableLineStyle.SelectedIndex = 4; break;
                        case LineStyle.Double: this.cb_tableLineStyle.SelectedIndex = 5; break;
                        case LineStyle.Triple: this.cb_tableLineStyle.SelectedIndex = 6; break;
                        case LineStyle.ThinThickSmallGap: this.cb_tableLineStyle.SelectedIndex = 7; break;
                        default: this.cb_tableLineStyle.SelectedIndex = 0; break;
                    }
                }

                this.tb_tableLeftIndent.Text = ws.TableLeftIndent.ToString();
                this.tb_tableRightIndent.Text = ws.TableRightIndent.ToString();
                this.cb_tableIsShading.Checked = ws.TableIsShading;
                this.lb_tableShadingColor.Tag = ws.TableShadingColor;
                this.lb_tableShadingColor.BackColor = ws.TableShadingColor;
                this.cb_locationTable.Text = ws.TableAlignment.ToString();
                //边框样式
                #endregion

                #region 标题和内容
                this.uc_oneFont.fontSelect = ws.TitleOneFont;
                this.uc_oneFont.fontColorSelect = ws.TitleOneColor;
                //this.uc_oneFont.ForeColor = ws.TitleOneColor;
                this.uc_oneFont.SettingsControl(ws.TitleOneFont, ws.TitleOneColor);
                //this.uc_oneFont.BackColor
                //this.lblLv1Font.Text = string.Format("字体:{0}  字号: {1}", ws.TitleOneFont.Name, ws.TitleOneFont.Size);
                //this.lblLv1Font.Tag = ws.TitleOneFont;
                //this.lblLv1Color.Tag = ws.TitleOneColor;
                //this.lblLv1Color.BackColor = ws.TitleOneColor;
                switch (ws.TitleOneLineSpace)
                {
                    case 1: this.cb_LineSpaceOne.SelectedIndex = 0; break;
                    case 1.5: this.cb_LineSpaceOne.SelectedIndex = 1; break;
                    case 2: this.cb_LineSpaceOne.SelectedIndex = 2; break;
                    case 3: this.cb_LineSpaceOne.SelectedIndex = 5; break;
                    default: this.cb_LineSpaceOne.SelectedIndex = 0; break;
                }
                switch (ws.TitleOneSpaceBefore)
                {
                    case 0: this.cb_spaceBeforeOne.SelectedIndex = 0; break;
                    case 0.5: this.cb_spaceBeforeOne.SelectedIndex = 1; break;
                    case 1: this.cb_spaceBeforeOne.SelectedIndex = 2; break;
                    case 1.5: this.cb_spaceBeforeOne.SelectedIndex = 3; break;
                    case 2: this.cb_spaceBeforeOne.SelectedIndex = 4; break;
                    case 2.5: this.cb_spaceBeforeOne.SelectedIndex = 5; break;
                    case 3: this.cb_spaceBeforeOne.SelectedIndex = 6; break;
                    case 3.5: this.cb_spaceBeforeOne.SelectedIndex = 7; break;
                    case 4: this.cb_spaceBeforeOne.SelectedIndex = 8; break;
                    default: this.cb_spaceBeforeOne.SelectedIndex = 0; break;
                }
                switch (ws.TitleOneSpaceAfter)
                {
                    case 0: this.cb_spaceAfterOne.SelectedIndex = 0; break;
                    case 0.5: this.cb_spaceAfterOne.SelectedIndex = 1; break;
                    case 1: this.cb_spaceAfterOne.SelectedIndex = 2; break;
                    case 1.5: this.cb_spaceAfterOne.SelectedIndex = 3; break;
                    case 2: this.cb_spaceAfterOne.SelectedIndex = 4; break;
                    case 2.5: this.cb_spaceAfterOne.SelectedIndex = 5; break;
                    case 3: this.cb_spaceAfterOne.SelectedIndex = 6; break;
                    case 3.5: this.cb_spaceAfterOne.SelectedIndex = 7; break;
                    case 4: this.cb_spaceAfterOne.SelectedIndex = 8; break;
                    default: this.cb_spaceAfterOne.SelectedIndex = 0; break;
                }
                this.tb_leftIndentOne.Text = ws.TitleOneLeftIndent.ToString();
                this.tb_rightIndentOne.Text = ws.TitleOneRightIndent.ToString();
                this.cb_locationOne.SelectedIndex = (int)ws.AlignmentOne;

                this.uc_twoFont.fontSelect = ws.TitleTwoFont;
                this.uc_twoFont.fontColorSelect = ws.TitleTwoColor;
                //this.uc_twoFont.BackColor = ws.TitleTwoColor;
                this.uc_twoFont.SettingsControl(ws.TitleTwoFont, ws.TitleTwoColor);
                switch (ws.TitleTwoLineSpace)
                {
                    case 1: this.cb_LineSpaceTwo.SelectedIndex = 0; break;
                    case 1.5: this.cb_LineSpaceTwo.SelectedIndex = 1; break;
                    case 2: this.cb_LineSpaceTwo.SelectedIndex = 2; break;
                    case 3: this.cb_LineSpaceTwo.SelectedIndex = 5; break;
                    default: this.cb_LineSpaceTwo.SelectedIndex = 0; break;
                }
                switch (ws.TitleTwoSpaceBefore)
                {
                    case 0: this.cb_spaceBeforeTwo.SelectedIndex = 0; break;
                    case 0.5: this.cb_spaceBeforeTwo.SelectedIndex = 1; break;
                    case 1: this.cb_spaceBeforeTwo.SelectedIndex = 2; break;
                    case 1.5: this.cb_spaceBeforeTwo.SelectedIndex = 3; break;
                    case 2: this.cb_spaceBeforeTwo.SelectedIndex = 4; break;
                    case 2.5: this.cb_spaceBeforeTwo.SelectedIndex = 5; break;
                    case 3: this.cb_spaceBeforeTwo.SelectedIndex = 6; break;
                    case 3.5: this.cb_spaceBeforeTwo.SelectedIndex = 7; break;
                    case 4: this.cb_spaceBeforeTwo.SelectedIndex = 8; break;
                    default: this.cb_spaceBeforeTwo.SelectedIndex = 0; break;
                }
                switch (ws.TitleTwoSpaceAfter)
                {
                    case 0: this.cb_spaceAfterTwo.SelectedIndex = 0; break;
                    case 0.5: this.cb_spaceAfterTwo.SelectedIndex = 1; break;
                    case 1: this.cb_spaceAfterTwo.SelectedIndex = 2; break;
                    case 1.5: this.cb_spaceAfterTwo.SelectedIndex = 3; break;
                    case 2: this.cb_spaceAfterTwo.SelectedIndex = 4; break;
                    case 2.5: this.cb_spaceAfterTwo.SelectedIndex = 5; break;
                    case 3: this.cb_spaceAfterTwo.SelectedIndex = 6; break;
                    case 3.5: this.cb_spaceAfterTwo.SelectedIndex = 7; break;
                    case 4: this.cb_spaceAfterTwo.SelectedIndex = 8; break;
                    default: this.cb_spaceAfterTwo.SelectedIndex = 0; break;
                }
                this.tb_leftIndentTwo.Text = ws.TitleTwoLeftIndent.ToString();
                this.tb_rightIndentTwo.Text = ws.TitleTwoRightIndent.ToString();
                this.cb_locationTwo.SelectedIndex = (int)ws.AlignmentTwo;

                this.uc_threeFont.fontSelect = ws.TitleThreeFont;
                this.uc_threeFont.fontColorSelect = ws.TitleThreeColor;
                //this.uc_threeFont.ForeColor = ws.TitleThreeColor;
                this.uc_threeFont.SettingsControl(ws.TitleThreeFont, ws.TitleThreeColor);
                switch (ws.TitleThreeLineSpace)
                {
                    case 1: this.cb_LineSpaceThree.SelectedIndex = 0; break;
                    case 1.5: this.cb_LineSpaceThree.SelectedIndex = 1; break;
                    case 2: this.cb_LineSpaceThree.SelectedIndex = 2; break;
                    case 3: this.cb_LineSpaceThree.SelectedIndex = 5; break;
                    default: this.cb_LineSpaceThree.SelectedIndex = 0; break;
                }
                switch (ws.TitleThreeSpaceBefore)
                {
                    case 0: this.cb_spaceBeforeThree.SelectedIndex = 0; break;
                    case 0.5: this.cb_spaceBeforeThree.SelectedIndex = 1; break;
                    case 1: this.cb_spaceBeforeThree.SelectedIndex = 2; break;
                    case 1.5: this.cb_spaceBeforeThree.SelectedIndex = 3; break;
                    case 2: this.cb_spaceBeforeThree.SelectedIndex = 4; break;
                    case 2.5: this.cb_spaceBeforeThree.SelectedIndex = 5; break;
                    case 3: this.cb_spaceBeforeThree.SelectedIndex = 6; break;
                    case 3.5: this.cb_spaceBeforeThree.SelectedIndex = 7; break;
                    case 4: this.cb_spaceBeforeThree.SelectedIndex = 8; break;
                    default: this.cb_spaceBeforeThree.SelectedIndex = 0; break;
                }
                switch (ws.TitleThreeSpaceAfter)
                {
                    case 0: this.cb_spaceAfterThree.SelectedIndex = 0; break;
                    case 0.5: this.cb_spaceAfterThree.SelectedIndex = 1; break;
                    case 1: this.cb_spaceAfterThree.SelectedIndex = 2; break;
                    case 1.5: this.cb_spaceAfterThree.SelectedIndex = 3; break;
                    case 2: this.cb_spaceAfterThree.SelectedIndex = 4; break;
                    case 2.5: this.cb_spaceAfterThree.SelectedIndex = 5; break;
                    case 3: this.cb_spaceAfterThree.SelectedIndex = 6; break;
                    case 3.5: this.cb_spaceAfterThree.SelectedIndex = 7; break;
                    case 4: this.cb_spaceAfterThree.SelectedIndex = 8; break;
                    default: this.cb_spaceAfterThree.SelectedIndex = 0; break;
                }
                this.tb_leftIndentThree.Text = ws.TitleThreeLeftIndent.ToString();
                this.tb_rightIndentThree.Text = ws.TitleThreeRightIndent.ToString();
                this.cb_locationThree.SelectedIndex = (int)ws.AlignmentThree;

                this.uc_fourFont.fontSelect = ws.TitleFourFont;
                this.uc_fourFont.fontColorSelect = ws.TitleFourColor;
                //this.uc_fourFont.ForeColor = ws.TitleFourColor;
                this.uc_fourFont.SettingsControl(ws.TitleFourFont, ws.TitleFourColor);
                switch (ws.TitleFourLineSpace)
                {
                    case 1: this.cb_LineSpaceFour.SelectedIndex = 0; break;
                    case 1.5: this.cb_LineSpaceFour.SelectedIndex = 1; break;
                    case 2: this.cb_LineSpaceFour.SelectedIndex = 2; break;
                    case 3: this.cb_LineSpaceFour.SelectedIndex = 5; break;
                    default: this.cb_LineSpaceFour.SelectedIndex = 0; break;
                }
                switch (ws.TitleFourSpaceBefore)
                {
                    case 0: this.cb_spaceBeforeFour.SelectedIndex = 0; break;
                    case 0.5: this.cb_spaceBeforeFour.SelectedIndex = 1; break;
                    case 1: this.cb_spaceBeforeFour.SelectedIndex = 2; break;
                    case 1.5: this.cb_spaceBeforeFour.SelectedIndex = 3; break;
                    case 2: this.cb_spaceBeforeFour.SelectedIndex = 4; break;
                    case 2.5: this.cb_spaceBeforeFour.SelectedIndex = 5; break;
                    case 3: this.cb_spaceBeforeFour.SelectedIndex = 6; break;
                    case 3.5: this.cb_spaceBeforeFour.SelectedIndex = 7; break;
                    case 4: this.cb_spaceBeforeFour.SelectedIndex = 8; break;
                    default: this.cb_spaceBeforeFour.SelectedIndex = 0; break;
                }
                switch (ws.TitleFourSpaceAfter)
                {
                    case 0: this.cb_spaceAfterFour.SelectedIndex = 0; break;
                    case 0.5: this.cb_spaceAfterFour.SelectedIndex = 1; break;
                    case 1: this.cb_spaceAfterFour.SelectedIndex = 2; break;
                    case 1.5: this.cb_spaceAfterFour.SelectedIndex = 3; break;
                    case 2: this.cb_spaceAfterFour.SelectedIndex = 4; break;
                    case 2.5: this.cb_spaceAfterFour.SelectedIndex = 5; break;
                    case 3: this.cb_spaceAfterFour.SelectedIndex = 6; break;
                    case 3.5: this.cb_spaceAfterFour.SelectedIndex = 7; break;
                    case 4: this.cb_spaceAfterFour.SelectedIndex = 8; break;
                    default: this.cb_spaceAfterFour.SelectedIndex = 0; break;
                }
                this.tb_leftIndentFour.Text = ws.TitleFourLeftIndent.ToString();
                this.tb_rightIndentFour.Text = ws.TitleFourRightIndent.ToString();
                this.cb_locationFour.SelectedIndex = (int)ws.AlignmentFour;

                this.uc_fiveFont.fontSelect = ws.TitleFiveFont;
                this.uc_fiveFont.fontColorSelect = ws.TitleFiveColor;
                //this.uc_fiveFont.ForeColor = ws.TitleFiveColor;
                this.uc_fiveFont.SettingsControl(ws.TitleFiveFont, ws.TitleFiveColor);
                switch (ws.TitleFiveLineSpace)
                {
                    case 1: this.cb_LineSpaceFive.SelectedIndex = 0; break;
                    case 1.5: this.cb_LineSpaceFive.SelectedIndex = 1; break;
                    case 2: this.cb_LineSpaceFive.SelectedIndex = 2; break;
                    case 3: this.cb_LineSpaceFive.SelectedIndex = 5; break;
                    default: this.cb_LineSpaceFive.SelectedIndex = 0; break;
                }
                switch (ws.TitleFiveSpaceBefore)
                {
                    case 0: this.cb_spaceBeforeFive.SelectedIndex = 0; break;
                    case 0.5: this.cb_spaceBeforeFive.SelectedIndex = 1; break;
                    case 1: this.cb_spaceBeforeFive.SelectedIndex = 2; break;
                    case 1.5: this.cb_spaceBeforeFive.SelectedIndex = 3; break;
                    case 2: this.cb_spaceBeforeFive.SelectedIndex = 4; break;
                    case 2.5: this.cb_spaceBeforeFive.SelectedIndex = 5; break;
                    case 3: this.cb_spaceBeforeFive.SelectedIndex = 6; break;
                    case 3.5: this.cb_spaceBeforeFive.SelectedIndex = 7; break;
                    case 4: this.cb_spaceBeforeFive.SelectedIndex = 8; break;
                    default: this.cb_spaceBeforeFive.SelectedIndex = 0; break;
                }
                switch (ws.TitleFiveSpaceAfter)
                {
                    case 0: this.cb_spaceAfterFive.SelectedIndex = 0; break;
                    case 0.5: this.cb_spaceAfterFive.SelectedIndex = 1; break;
                    case 1: this.cb_spaceAfterFive.SelectedIndex = 2; break;
                    case 1.5: this.cb_spaceAfterFive.SelectedIndex = 3; break;
                    case 2: this.cb_spaceAfterFive.SelectedIndex = 4; break;
                    case 2.5: this.cb_spaceAfterFive.SelectedIndex = 5; break;
                    case 3: this.cb_spaceAfterFive.SelectedIndex = 6; break;
                    case 3.5: this.cb_spaceAfterFive.SelectedIndex = 7; break;
                    case 4: this.cb_spaceAfterFive.SelectedIndex = 8; break;
                    default: this.cb_spaceAfterFive.SelectedIndex = 0; break;
                }
                this.tb_leftIndentFive.Text = ws.TitleFiveLeftIndent.ToString();
                this.tb_rightIndentFive.Text = ws.TitleFiveRightIndent.ToString();
                this.cb_locationFive.SelectedIndex = (int)ws.AlignmentFive;

                this.uc_contentFont.fontSelect = ws.ContentFont;
                this.uc_contentFont.fontColorSelect = ws.ContentColor;
                //this.uc_contentFont.ForeColor = ws.ContentColor;
                this.uc_contentFont.SettingsControl(ws.ContentFont, ws.ContentColor);
                switch (ws.ContentLineSpace)
                {
                    case 1: this.cb_LineSpaceContent.SelectedIndex = 0; break;
                    case 1.5: this.cb_LineSpaceContent.SelectedIndex = 1; break;
                    case 2: this.cb_LineSpaceContent.SelectedIndex = 2; break;
                    case 3: this.cb_LineSpaceContent.SelectedIndex = 5; break;
                    default: this.cb_LineSpaceContent.SelectedIndex = 0; break;
                }
                switch (ws.ContentSpaceBefore)
                {
                    case 0: this.cb_spaceBeforeContent.SelectedIndex = 0; break;
                    case 0.5: this.cb_spaceBeforeContent.SelectedIndex = 1; break;
                    case 1: this.cb_spaceBeforeContent.SelectedIndex = 2; break;
                    case 1.5: this.cb_spaceBeforeContent.SelectedIndex = 3; break;
                    case 2: this.cb_spaceBeforeContent.SelectedIndex = 4; break;
                    case 2.5: this.cb_spaceBeforeContent.SelectedIndex = 5; break;
                    case 3: this.cb_spaceBeforeContent.SelectedIndex = 6; break;
                    case 3.5: this.cb_spaceBeforeContent.SelectedIndex = 7; break;
                    case 4: this.cb_spaceBeforeContent.SelectedIndex = 8; break;
                    default: this.cb_spaceBeforeContent.SelectedIndex = 0; break;
                }
                switch (ws.ContentSpaceAfter)
                {
                    case 0: this.cb_spaceAfterContent.SelectedIndex = 0; break;
                    case 0.5: this.cb_spaceAfterContent.SelectedIndex = 1; break;
                    case 1: this.cb_spaceAfterContent.SelectedIndex = 2; break;
                    case 1.5: this.cb_spaceAfterContent.SelectedIndex = 3; break;
                    case 2: this.cb_spaceAfterContent.SelectedIndex = 4; break;
                    case 2.5: this.cb_spaceAfterContent.SelectedIndex = 5; break;
                    case 3: this.cb_spaceAfterContent.SelectedIndex = 6; break;
                    case 3.5: this.cb_spaceAfterContent.SelectedIndex = 7; break;
                    case 4: this.cb_spaceAfterContent.SelectedIndex = 8; break;
                    default: this.cb_spaceAfterContent.SelectedIndex = 0; break;
                }
                this.tb_leftIndentContent.Text = ws.ContentLeftIndent.ToString();
                this.tb_rightIndentContent.Text = ws.ContentRightIndent.ToString();
                this.cb_locationContent.SelectedIndex = (int)ws.AlignmentContent;
                #endregion
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void btn_UploadLogo_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.Filter = "图片文件|*.bmp;*.jpg;*.jpeg;*.png;*.ico";
            if (of.ShowDialog() == DialogResult.OK)
            {
                string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Imgs");
                path = Path.Combine(path, Guid.NewGuid().ToString("N") + Path.GetExtension(of.SafeFileName));
                if (File.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                if (Path.GetExtension(of.SafeFileName) == ".ico")
                {
                    Icon icon = new Icon(of.FileName);
                    MemoryStream mStream = new MemoryStream();
                    icon.Save(mStream);
                    Image image = Image.FromStream(mStream);
                    lb_Logo.Image = image;
                    File.Copy(of.FileName, path);
                    logoPath = path;
                    //image.Save(path);
                }
                else
                {
                    lb_Logo.Image = new Bitmap(of.FileName);
                    File.Copy(of.FileName, path);
                    logoPath = path;
                    //lb_Logo.Image.Save(path);
                }
            }
        }

        private List<Image> images = new List<Image>();
        private void cb_tableLineStyle_DrawItem(object sender, DrawItemEventArgs e)
        {
            e.DrawBackground();
            if (images.Count == 0) return;
            e.Graphics.DrawImage(images[e.Index], 0, e.Bounds.Y, 90, 18);
        }

        private void cmb_Company_Click(object sender, EventArgs e)
        {
            #region 第一版
            //string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WordStyles");
            //if (!Directory.Exists(path)) Directory.CreateDirectory(path);
            //var files = Directory.GetFiles(path, "*.txt", SearchOption.AllDirectories);
            //for (int i = 0; i < files.Length; i++)
            //{
            //    var ss = Path.GetFileName(files[i]);
            //    if (!fileNames.Contains(ss))
            //    {
            //        fileNames.Add(ss);
            //        cmb_Company.Items.Add(Path.GetFileNameWithoutExtension(ss));
            //    }
            //} 
            #endregion
            try
            {
                while (cmb_Company.Items.Count > 0)
                {
                    cmb_Company.Items.RemoveAt(0);
                }
                List<string> result = new List<string>();
                string nUrl = "https://" + ConfigurationManager.AppSettings["Ip"] + "/handleWord/getNames";
                string res = CommonTool.httpApi(nUrl, "", "get");
                if (res == "-1")
                {
                    result = JsonConvert.DeserializeObject<List<string>>(res);
                }

                for (int i = 0; i < result.Count; i++)
                {
                    cmb_Company.Items.Add(result[i]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btn_AddStyle_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(tb_styleName.Text) && cmb_Company.SelectedItem == null && string.IsNullOrWhiteSpace(cmb_Company.SelectedItem.ToString()))
            {
                MessageBox.Show("请填写样式名称或者选择投标企业！"); return;
            }
            try
            {
                string name = tb_styleName.Text;
                name = "{\"name\":\"" + name + "\"}";
                string eUrl = "https://" + ConfigurationManager.AppSettings["Ip"] + "/handleWord/existWordStyle";
                string count = CommonTool.httpApi(eUrl, name);
                if (int.Parse(count) > 0)
                {
                    MessageBoxButtons messButton = MessageBoxButtons.OKCancel;
                    DialogResult dr = MessageBox.Show("已存在格式，是否替换?", "确定", messButton);
                    if (dr != DialogResult.OK)//如果点击“确定”按钮
                    {
                        return;
                    }
                }
                WordStyle ws = saveStyleToTxt();
                string url = "https://" + ConfigurationManager.AppSettings["Ip"] + "/handleWord/addWordStyle";
                WordType wt = new WordType();
                wt.Name = tb_styleName.Text;
                wt.Content = JsonConvert.SerializeObject(ws);

                var ss = CommonTool.httpApi(url, JsonConvert.SerializeObject(wt));
                int res = Convert.ToInt32(ss);
                //int res = CommonMethods.saveTxtOfWordStyle(ws, tb_styleName.Text + ".txt");
                if (res > 0)
                {
                    MessageBox.Show("添加成功");
                    tb_styleName.Text = "";
                }
                else
                {
                    MessageBox.Show("添加失败");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        /// <summary>
        /// 获取样式类
        /// </summary>
        /// <returns></returns>
        private WordStyle saveStyleToTxt()
        {
            WordStyle ws = new WordStyle();

            #region 页面设置
            switch (cb_pageType.SelectedItem.ToString())
            {
                case "A3": ws.PageType = PaperSize.A3; break;
                case "A4": ws.PageType = PaperSize.A4; break;
                case "A5": ws.PageType = PaperSize.A5; break;
                case "B4": ws.PageType = PaperSize.B4; break;
                case "B5": ws.PageType = PaperSize.B5; break;
                default: ws.PageType = PaperSize.A4; break;
            }
            switch (cb_pageDirection.SelectedItem.ToString())
            {
                case "纵向": ws.PageDirection = Aspose.Words.Orientation.Portrait; break;
                case "横向": ws.PageDirection = Aspose.Words.Orientation.Landscape; break;
                default: ws.PageDirection = Aspose.Words.Orientation.Landscape; break;
            }
            ws.TopMargin = Convert.ToDouble(tb_pageMarginUp.Text);
            ws.BottomMargin = Convert.ToDouble(tb_pageMarginDown.Text);
            ws.LeftMargin = Convert.ToDouble(tb_pageMarginLeft.Text);
            ws.RightMargin = Convert.ToDouble(tb_pageMarginRight.Text);
            #endregion

            #region 页眉
            ws.HeaderName = tb_HeaderTxt.Text;
            ws.HeaderFont = uc_headerFont.fontSelect;
            ws.HeaderColor = uc_headerFont.fontColorSelect;
            switch (cmb_header.SelectedItem.ToString())
            {
                case "居左":
                    ws.HeaderAlignment = ParagraphAlignment.Left;
                    break;
                case "居中":
                    ws.HeaderAlignment = ParagraphAlignment.Center;
                    break;
                case "居右":
                    ws.HeaderAlignment = ParagraphAlignment.Right;
                    break;
                case "左右分散":
                    ws.HeaderAlignment = ParagraphAlignment.Justify;
                    break;
                default:
                    ws.HeaderAlignment = ParagraphAlignment.Center;
                    break;
            }
            ws.HImgPath = logoPath;
            ws.HLeftIndent = Convert.ToDouble(tb_hLeftIndent.Text);
            ws.HRightIndent = Convert.ToDouble(tb_hRightIndent.Text);
            ws.HeaderDistance = Convert.ToDouble(tb_headerDistance.Text);
            switch (cb_hLineSpace.SelectedItem.ToString())
            {
                case "单倍行距":
                    ws.HLineSpace = 1;//12就是一倍行距 
                    break;
                case "1.5倍行距":
                    ws.HLineSpace = 1.5;//12就是一倍行距 
                    break;
                case "2倍行距":
                    ws.HLineSpace = 2;//12就是一倍行距 
                    break;
                case "最小值":
                    ws.HLineSpace = 1;//12就是一倍行距 
                    break;
                case "固定值":
                    ws.HLineSpace = 1;//12就是一倍行距 
                    break;
                case "多倍行距":
                    ws.HLineSpace = 3;//12就是一倍行距 
                    break;
                default:
                    ws.HLineSpace = 1;//12就是一倍行距 
                    break;
            }
            ws.HIsFirstDif = cb_fIsFirstDif.Checked;
            ws.HIsParityDif = cb_hIsParityDif.Checked;
            #endregion

            #region 页脚
            ws.FooterName = tb_FooterTxt.Text;
            ws.FooterFont = uc_footerFont.fontSelect;
            ws.FooterColor = uc_footerFont.fontColorSelect;
            switch (cmb_footer.SelectedItem.ToString())
            {
                case "居左":
                    ws.FooterAlignment = ParagraphAlignment.Left;
                    break;
                case "居中":
                    ws.FooterAlignment = ParagraphAlignment.Center;
                    break;
                case "居右":
                    ws.FooterAlignment = ParagraphAlignment.Right;
                    break;
                case "左右分散":
                    ws.HeaderAlignment = ParagraphAlignment.Justify;
                    break;
                default:
                    ws.FooterAlignment = ParagraphAlignment.Center;
                    break;
            }
            ws.FLeftIndent = Convert.ToDouble(tb_fLeftIndent.Text);
            ws.FRightIndent = Convert.ToDouble(tb_fRightIndent.Text);
            ws.FooterDistance = Convert.ToDouble(tb_FooterDistance.Text);
            switch (cb_fLineSpace.SelectedItem.ToString())
            {
                case "单倍行距":
                    ws.FLineSpace = 1;//12就是一倍行距 
                    break;
                case "1.5倍行距":
                    ws.FLineSpace = 1.5;//12就是一倍行距 
                    break;
                case "2倍行距":
                    ws.FLineSpace = 2;//12就是一倍行距 
                    break;
                case "最小值":
                    ws.FLineSpace = 1;//12就是一倍行距 
                    break;
                case "固定值":
                    ws.FLineSpace = 1;//12就是一倍行距 
                    break;
                case "多倍行距":
                    ws.FLineSpace = 3;//12就是一倍行距 
                    break;
                default:
                    ws.FLineSpace = 1;//12就是一倍行距 
                    break;
            }
            ws.FIsFirstDif = cb_fIsFirstDif.Checked;
            ws.FIsParityDif = cb_fIsparityDif.Checked;
            #endregion

            #region 表格
            ws.TableFont = uc_tableFont.fontSelect;
            ws.TableFontColor = uc_tableFont.fontColorSelect;
            ws.TableBorderColor = (Color)lblTableBorderColor.Tag;
            switch (cb_tableLineSpace.SelectedItem.ToString())
            {
                case "单倍行距":
                    ws.TableLineSpace = 1;//12就是一倍行距 
                    break;
                case "1.5倍行距":
                    ws.TableLineSpace = 1.5;//12就是一倍行距 
                    break;
                case "2倍行距":
                    ws.TableLineSpace = 2;//12就是一倍行距 
                    break;
                case "最小值":
                    ws.TableLineSpace = 1;//12就是一倍行距 
                    break;
                case "固定值":
                    ws.TableLineSpace = 1;//12就是一倍行距 
                    break;
                case "多倍行距":
                    ws.TableLineSpace = 3;//12就是一倍行距 
                    break;
                default:
                    ws.TableLineSpace = 1;//12就是一倍行距 
                    break;
            }
            switch (cb_locationTable.Text)
            {
                case "LeftUp": ws.TableAlignment = Model.Enum_CellAlignment.LeftUp; break;
                case "CenterUp": ws.TableAlignment = Model.Enum_CellAlignment.CenterUp; break;
                case "RightUp": ws.TableAlignment = Model.Enum_CellAlignment.RightUp; break;
                case "LeftMiddle": ws.TableAlignment = Model.Enum_CellAlignment.LeftMiddle; break;
                case "CenterMiddle": ws.TableAlignment = Model.Enum_CellAlignment.CenterMiddle; break;
                case "RightMiddle": ws.TableAlignment = Model.Enum_CellAlignment.RightMiddle; break;
                case "LeftBottom": ws.TableAlignment = Model.Enum_CellAlignment.LeftBottom; break;
                case "CenterBottom": ws.TableAlignment = Model.Enum_CellAlignment.CenterBottom; break;
                case "RightBottom": ws.TableAlignment = Model.Enum_CellAlignment.RightBottom; break;
                default:
                    ws.TableAlignment = Model.Enum_CellAlignment.CenterMiddle;
                    break;
            }
            switch (cb_tableLineStyle.SelectedItem.ToString())
            {
                case "1": ws.TableLineStyle = LineStyle.Single; break;
                case "7": ws.TableLineStyle = LineStyle.DashLargeGap; break;
                case "6": ws.TableLineStyle = LineStyle.Dot; break;
                case "8": ws.TableLineStyle = LineStyle.DotDash; break;
                case "9": ws.TableLineStyle = LineStyle.DotDotDash; break;
                case "3": ws.TableLineStyle = LineStyle.Double; break;
                case "10": ws.TableLineStyle = LineStyle.Triple; break;
                case "11": ws.TableLineStyle = LineStyle.ThinThickSmallGap; break;
                default:
                    ws.TableLineStyle = LineStyle.Single;
                    break;
            }

            switch (cb_tableLineWidth.SelectedItem.ToString())
            {
                case "0.25磅": ws.TableLineWidth = 0.25; break;
                case "0.5磅": ws.TableLineWidth = 0.5; break;
                case "0.75磅": ws.TableLineWidth = 0.75; break;
                case "1.0磅": ws.TableLineWidth = 1.0; break;
                case "1.5磅": ws.TableLineWidth = 1.5; break;
                case "2.25磅": ws.TableLineWidth = 2.25; break;
                case "3.0磅": ws.TableLineWidth = 3.0; break;
                case "4.5磅": ws.TableLineWidth = 4.5; break;
                case "6.0磅": ws.TableLineWidth = 6.0; break;
                default: ws.TableLineWidth = 1.0; break;
            }

            switch (cb_tableSpaceBefore.SelectedItem.ToString())
            {
                case "0 行": ws.TableSpaceBefore = 0; break;
                case "0.5 行": ws.TableSpaceBefore = 0.5; break;
                case "1 行": ws.TableSpaceBefore = 1; break;
                case "1.5 行": ws.TableSpaceBefore = 1.5; break;
                case "2 行": ws.TableSpaceBefore = 2; break;
                case "2.5 行": ws.TableSpaceBefore = 2.5; break;
                case "3 行": ws.TableSpaceBefore = 3; break;
                case "3.5 行": ws.TableSpaceBefore = 3.5; break;
                case "4 行": ws.TableSpaceBefore = 4; break;
                default: ws.TableSpaceBefore = 0; break;
            }

            switch (cb_tableSpaceAfter.SelectedItem.ToString())
            {
                case "0 行": ws.TableSpaceAfter = 0; break;
                case "0.5 行": ws.TableSpaceAfter = 0.5; break;
                case "1 行": ws.TableSpaceAfter = 1; break;
                case "1.5 行": ws.TableSpaceAfter = 1.5; break;
                case "2 行": ws.TableSpaceAfter = 2; break;
                case "2.5 行": ws.TableSpaceAfter = 2.5; break;
                case "3 行": ws.TableSpaceAfter = 3; break;
                case "3.5 行": ws.TableSpaceAfter = 3.5; break;
                case "4 行": ws.TableSpaceAfter = 4; break;
                default: ws.TableSpaceAfter = 0; break;
            }
            ws.TableLeftIndent = Convert.ToDouble(tb_tableLeftIndent.Text);
            ws.TableRightIndent = Convert.ToDouble(tb_tableRightIndent.Text);
            ws.TableIsShading = cb_tableIsShading.Checked;
            ws.TableShadingColor = (Color)lb_tableShadingColor.Tag;
            #endregion

            #region 标题和内容
            ws.TitleOneFont = uc_oneFont.fontSelect;
            ws.TitleOneColor = uc_oneFont.fontColorSelect;
            switch (cb_LineSpaceOne.SelectedItem.ToString())
            {
                case "单倍行距":
                    ws.TitleOneLineSpace = 1;//12就是一倍行距 
                    break;
                case "1.5倍行距":
                    ws.TitleOneLineSpace = 1.5;//12就是一倍行距 
                    break;
                case "2倍行距":
                    ws.TitleOneLineSpace = 2;//12就是一倍行距 
                    break;
                case "最小值":
                    ws.TitleOneLineSpace = 1;//12就是一倍行距 
                    break;
                case "固定值":
                    ws.TitleOneLineSpace = 1;//12就是一倍行距 
                    break;
                case "多倍行距":
                    ws.TitleOneLineSpace = 3;//12就是一倍行距 
                    break;
                default:
                    ws.TitleOneLineSpace = 1;//12就是一倍行距 
                    break;
            }
            switch (cb_spaceBeforeOne.SelectedItem.ToString())
            {
                case "0 行": ws.TitleOneSpaceBefore = 0; break;
                case "0.5 行": ws.TitleOneSpaceBefore = 0.5; break;
                case "1 行": ws.TitleOneSpaceBefore = 1; break;
                case "1.5 行": ws.TitleOneSpaceBefore = 1.5; break;
                case "2 行": ws.TitleOneSpaceBefore = 2; break;
                case "2.5 行": ws.TitleOneSpaceBefore = 2.5; break;
                case "3 行": ws.TitleOneSpaceBefore = 3; break;
                case "3.5 行": ws.TitleOneSpaceBefore = 3.5; break;
                case "4 行": ws.TitleOneSpaceBefore = 4; break;
                default: ws.TitleOneSpaceBefore = 0; break;
            }
            switch (cb_spaceAfterOne.SelectedItem.ToString())
            {
                case "0 行": ws.TitleOneSpaceAfter = 0; break;
                case "0.5 行": ws.TitleOneSpaceAfter = 0.5; break;
                case "1 行": ws.TitleOneSpaceAfter = 1; break;
                case "1.5 行": ws.TitleOneSpaceAfter = 1.5; break;
                case "2 行": ws.TitleOneSpaceAfter = 2; break;
                case "2.5 行": ws.TitleOneSpaceAfter = 2.5; break;
                case "3 行": ws.TitleOneSpaceAfter = 3; break;
                case "3.5 行": ws.TitleOneSpaceAfter = 3.5; break;
                case "4 行": ws.TitleOneSpaceAfter = 4; break;
                default: ws.TitleOneSpaceAfter = 0; break;
            }
            ws.TitleOneLeftIndent = Convert.ToDouble(tb_leftIndentOne.Text);
            ws.TitleOneRightIndent = Convert.ToDouble(tb_rightIndentOne.Text);
            switch (cb_locationOne.SelectedItem.ToString())
            {
                case "居左": ws.AlignmentOne = ParagraphAlignment.Left; break;
                case "居中": ws.AlignmentOne = ParagraphAlignment.Center; break;
                case "居右": ws.AlignmentOne = ParagraphAlignment.Right; break;
                case "左右分散":
                    ws.HeaderAlignment = ParagraphAlignment.Justify;
                    break;
                default: ws.AlignmentOne = ParagraphAlignment.Left; break;
            }

            ws.TitleTwoFont = uc_twoFont.fontSelect;
            ws.TitleTwoColor = uc_twoFont.fontColorSelect;
            switch (cb_LineSpaceTwo.SelectedItem.ToString())
            {
                case "单倍行距":
                    ws.TitleTwoLineSpace = 1;//12就是一倍行距 
                    break;
                case "1.5倍行距":
                    ws.TitleTwoLineSpace = 1.5;//12就是一倍行距 
                    break;
                case "2倍行距":
                    ws.TitleTwoLineSpace = 2;//12就是一倍行距 
                    break;
                case "最小值":
                    ws.TitleTwoLineSpace = 1;//12就是一倍行距 
                    break;
                case "固定值":
                    ws.TitleTwoLineSpace = 1;//12就是一倍行距 
                    break;
                case "多倍行距":
                    ws.TitleTwoLineSpace = 3;//12就是一倍行距 
                    break;
                default:
                    ws.TitleTwoLineSpace = 1;//12就是一倍行距 
                    break;
            }
            switch (cb_spaceBeforeTwo.SelectedItem.ToString())
            {
                case "0 行": ws.TitleTwoSpaceBefore = 0; break;
                case "0.5 行": ws.TitleTwoSpaceBefore = 0.5; break;
                case "1 行": ws.TitleTwoSpaceBefore = 1; break;
                case "1.5 行": ws.TitleTwoSpaceBefore = 1.5; break;
                case "2 行": ws.TitleTwoSpaceBefore = 2; break;
                case "2.5 行": ws.TitleTwoSpaceBefore = 2.5; break;
                case "3 行": ws.TitleTwoSpaceBefore = 3; break;
                case "3.5 行": ws.TitleTwoSpaceBefore = 3.5; break;
                case "4 行": ws.TitleTwoSpaceBefore = 4; break;
                default: ws.TitleTwoSpaceBefore = 0; break;
            }
            switch (cb_spaceAfterTwo.SelectedItem.ToString())
            {
                case "0 行": ws.TitleTwoSpaceAfter = 0; break;
                case "0.5 行": ws.TitleTwoSpaceAfter = 0.5; break;
                case "1 行": ws.TitleTwoSpaceAfter = 1; break;
                case "1.5 行": ws.TitleTwoSpaceAfter = 1.5; break;
                case "2 行": ws.TitleTwoSpaceAfter = 2; break;
                case "2.5 行": ws.TitleTwoSpaceAfter = 2.5; break;
                case "3 行": ws.TitleTwoSpaceAfter = 3; break;
                case "3.5 行": ws.TitleTwoSpaceAfter = 3.5; break;
                case "4 行": ws.TitleTwoSpaceAfter = 4; break;
                default: ws.TitleTwoSpaceAfter = 0; break;
            }
            ws.TitleTwoLeftIndent = Convert.ToDouble(tb_leftIndentTwo.Text);
            ws.TitleTwoRightIndent = Convert.ToDouble(tb_rightIndentTwo.Text);
            switch (cb_locationTwo.SelectedItem.ToString())
            {
                case "居左": ws.AlignmentTwo = ParagraphAlignment.Left; break;
                case "居中": ws.AlignmentTwo = ParagraphAlignment.Center; break;
                case "居右": ws.AlignmentTwo = ParagraphAlignment.Right; break;
                case "左右分散":
                    ws.HeaderAlignment = ParagraphAlignment.Justify;
                    break;
                default: ws.AlignmentTwo = ParagraphAlignment.Left; break;
            }

            ws.TitleThreeFont = uc_threeFont.fontSelect;
            ws.TitleThreeColor = uc_threeFont.fontColorSelect;
            switch (cb_LineSpaceThree.SelectedItem.ToString())
            {
                case "单倍行距":
                    ws.TitleThreeLineSpace = 1;//12就是一倍行距 
                    break;
                case "1.5倍行距":
                    ws.TitleThreeLineSpace = 1.5;//12就是一倍行距 
                    break;
                case "2倍行距":
                    ws.TitleThreeLineSpace = 2;//12就是一倍行距 
                    break;
                case "最小值":
                    ws.TitleThreeLineSpace = 1;//12就是一倍行距 
                    break;
                case "固定值":
                    ws.TitleThreeLineSpace = 1;//12就是一倍行距 
                    break;
                case "多倍行距":
                    ws.TitleThreeLineSpace = 3;//12就是一倍行距 
                    break;
                default:
                    ws.TitleThreeLineSpace = 1;//12就是一倍行距 
                    break;
            }
            switch (cb_spaceBeforeThree.SelectedItem.ToString())
            {
                case "0 行": ws.TitleThreeSpaceBefore = 0; break;
                case "0.5 行": ws.TitleThreeSpaceBefore = 0.5; break;
                case "1 行": ws.TitleThreeSpaceBefore = 1; break;
                case "1.5 行": ws.TitleThreeSpaceBefore = 1.5; break;
                case "2 行": ws.TitleThreeSpaceBefore = 2; break;
                case "2.5 行": ws.TitleThreeSpaceBefore = 2.5; break;
                case "3 行": ws.TitleThreeSpaceBefore = 3; break;
                case "3.5 行": ws.TitleThreeSpaceBefore = 3.5; break;
                case "4 行": ws.TitleThreeSpaceBefore = 4; break;
                default: ws.TitleThreeSpaceBefore = 0; break;
            }
            switch (cb_spaceAfterThree.SelectedItem.ToString())
            {
                case "0 行": ws.TitleThreeSpaceAfter = 0; break;
                case "0.5 行": ws.TitleThreeSpaceAfter = 0.5; break;
                case "1 行": ws.TitleThreeSpaceAfter = 1; break;
                case "1.5 行": ws.TitleThreeSpaceAfter = 1.5; break;
                case "2 行": ws.TitleThreeSpaceAfter = 2; break;
                case "2.5 行": ws.TitleThreeSpaceAfter = 2.5; break;
                case "3 行": ws.TitleThreeSpaceAfter = 3; break;
                case "3.5 行": ws.TitleThreeSpaceAfter = 3.5; break;
                case "4 行": ws.TitleThreeSpaceAfter = 4; break;
                default: ws.TitleThreeSpaceAfter = 0; break;
            }
            ws.TitleThreeLeftIndent = Convert.ToDouble(tb_leftIndentThree.Text);
            ws.TitleThreeRightIndent = Convert.ToDouble(tb_rightIndentThree.Text);
            switch (cb_locationThree.SelectedItem.ToString())
            {
                case "居左": ws.AlignmentThree = ParagraphAlignment.Left; break;
                case "居中": ws.AlignmentThree = ParagraphAlignment.Center; break;
                case "居右": ws.AlignmentThree = ParagraphAlignment.Right; break;
                case "左右分散":
                    ws.HeaderAlignment = ParagraphAlignment.Justify;
                    break;
                default: ws.AlignmentThree = ParagraphAlignment.Left; break;
            }

            ws.TitleFourFont = uc_fourFont.fontSelect;
            ws.TitleFourColor = uc_fourFont.fontColorSelect;
            switch (cb_LineSpaceFour.SelectedItem.ToString())
            {
                case "单倍行距":
                    ws.TitleFourLineSpace = 1;//12就是一倍行距 
                    break;
                case "1.5倍行距":
                    ws.TitleFourLineSpace = 1.5;//12就是一倍行距 
                    break;
                case "2倍行距":
                    ws.TitleFourLineSpace = 2;//12就是一倍行距 
                    break;
                case "最小值":
                    ws.TitleFourLineSpace = 1;//12就是一倍行距 
                    break;
                case "固定值":
                    ws.TitleFourLineSpace = 1;//12就是一倍行距 
                    break;
                case "多倍行距":
                    ws.TitleFourLineSpace = 3;//12就是一倍行距 
                    break;
                default:
                    ws.TitleFourLineSpace = 1;//12就是一倍行距 
                    break;
            }
            switch (cb_spaceBeforeFour.SelectedItem.ToString())
            {
                case "0 行": ws.TitleFourSpaceBefore = 0; break;
                case "0.5 行": ws.TitleFourSpaceBefore = 0.5; break;
                case "1 行": ws.TitleFourSpaceBefore = 1; break;
                case "1.5 行": ws.TitleFourSpaceBefore = 1.5; break;
                case "2 行": ws.TitleFourSpaceBefore = 2; break;
                case "2.5 行": ws.TitleFourSpaceBefore = 2.5; break;
                case "3 行": ws.TitleFourSpaceBefore = 3; break;
                case "3.5 行": ws.TitleFourSpaceBefore = 3.5; break;
                case "4 行": ws.TitleFourSpaceBefore = 4; break;
                default: ws.TitleFourSpaceBefore = 0; break;
            }
            switch (cb_spaceAfterFour.SelectedItem.ToString())
            {
                case "0 行": ws.TitleFourSpaceAfter = 0; break;
                case "0.5 行": ws.TitleFourSpaceAfter = 0.5; break;
                case "1 行": ws.TitleFourSpaceAfter = 1; break;
                case "1.5 行": ws.TitleFourSpaceAfter = 1.5; break;
                case "2 行": ws.TitleFourSpaceAfter = 2; break;
                case "2.5 行": ws.TitleFourSpaceAfter = 2.5; break;
                case "3 行": ws.TitleFourSpaceAfter = 3; break;
                case "3.5 行": ws.TitleFourSpaceAfter = 3.5; break;
                case "4 行": ws.TitleFourSpaceAfter = 4; break;
                default: ws.TitleFourSpaceAfter = 0; break;
            }
            ws.TitleFourLeftIndent = Convert.ToDouble(tb_leftIndentFour.Text);
            ws.TitleFourRightIndent = Convert.ToDouble(tb_rightIndentFour.Text);
            switch (cb_locationFour.SelectedItem.ToString())
            {
                case "居左": ws.AlignmentFour = ParagraphAlignment.Left; break;
                case "居中": ws.AlignmentFour = ParagraphAlignment.Center; break;
                case "居右": ws.AlignmentFour = ParagraphAlignment.Right; break;
                case "左右分散":
                    ws.HeaderAlignment = ParagraphAlignment.Justify;
                    break;
                default: ws.AlignmentFour = ParagraphAlignment.Left; break;
            }

            ws.TitleFiveFont = uc_fiveFont.fontSelect;
            ws.TitleFiveColor = uc_fiveFont.fontColorSelect;
            switch (cb_LineSpaceFive.SelectedItem.ToString())
            {
                case "单倍行距":
                    ws.TitleFiveLineSpace = 1;//12就是一倍行距 
                    break;
                case "1.5倍行距":
                    ws.TitleFiveLineSpace = 1.5;//12就是一倍行距 
                    break;
                case "2倍行距":
                    ws.TitleFiveLineSpace = 2;//12就是一倍行距 
                    break;
                case "最小值":
                    ws.TitleFiveLineSpace = 1;//12就是一倍行距 
                    break;
                case "固定值":
                    ws.TitleFiveLineSpace = 1;//12就是一倍行距 
                    break;
                case "多倍行距":
                    ws.TitleFiveLineSpace = 3;//12就是一倍行距 
                    break;
                default:
                    ws.TitleFiveLineSpace = 1;//12就是一倍行距 
                    break;
            }
            switch (cb_spaceBeforeFive.SelectedItem.ToString())
            {
                case "0 行": ws.TitleFiveSpaceBefore = 0; break;
                case "0.5 行": ws.TitleFiveSpaceBefore = 0.5; break;
                case "1 行": ws.TitleFiveSpaceBefore = 1; break;
                case "1.5 行": ws.TitleFiveSpaceBefore = 1.5; break;
                case "2 行": ws.TitleFiveSpaceBefore = 2; break;
                case "2.5 行": ws.TitleFiveSpaceBefore = 2.5; break;
                case "3 行": ws.TitleFiveSpaceBefore = 3; break;
                case "3.5 行": ws.TitleFiveSpaceBefore = 3.5; break;
                case "4 行": ws.TitleFiveSpaceBefore = 4; break;
                default: ws.TitleFiveSpaceBefore = 0; break;
            }
            switch (cb_spaceAfterFive.SelectedItem.ToString())
            {
                case "0 行": ws.TitleFiveSpaceAfter = 0; break;
                case "0.5 行": ws.TitleFiveSpaceAfter = 0.5; break;
                case "1 行": ws.TitleFiveSpaceAfter = 1; break;
                case "1.5 行": ws.TitleFiveSpaceAfter = 1.5; break;
                case "2 行": ws.TitleFiveSpaceAfter = 2; break;
                case "2.5 行": ws.TitleFiveSpaceAfter = 2.5; break;
                case "3 行": ws.TitleFiveSpaceAfter = 3; break;
                case "3.5 行": ws.TitleFiveSpaceAfter = 3.5; break;
                case "4 行": ws.TitleFiveSpaceAfter = 4; break;
                default: ws.TitleFiveSpaceAfter = 0; break;
            }
            ws.TitleFiveLeftIndent = Convert.ToDouble(tb_leftIndentFive.Text);
            ws.TitleFiveRightIndent = Convert.ToDouble(tb_rightIndentFive.Text);
            switch (cb_locationFive.SelectedItem.ToString())
            {
                case "居左": ws.AlignmentFive = ParagraphAlignment.Left; break;
                case "居中": ws.AlignmentFive = ParagraphAlignment.Center; break;
                case "居右": ws.AlignmentFive = ParagraphAlignment.Right; break;
                case "左右分散":
                    ws.HeaderAlignment = ParagraphAlignment.Justify;
                    break;
                default: ws.AlignmentFive = ParagraphAlignment.Left; break;
            }

            ws.ContentFont = uc_contentFont.fontSelect;
            ws.ContentColor = uc_contentFont.fontColorSelect;
            switch (cb_LineSpaceContent.SelectedItem.ToString())
            {
                case "单倍行距":
                    ws.ContentLineSpace = 1;//12就是一倍行距 
                    break;
                case "1.5倍行距":
                    ws.ContentLineSpace = 1.5;//12就是一倍行距 
                    break;
                case "2倍行距":
                    ws.ContentLineSpace = 2;//12就是一倍行距 
                    break;
                case "最小值":
                    ws.ContentLineSpace = 1;//12就是一倍行距 
                    break;
                case "固定值":
                    ws.ContentLineSpace = 1;//12就是一倍行距 
                    break;
                case "多倍行距":
                    ws.ContentLineSpace = 3;//12就是一倍行距 
                    break;
                default:
                    ws.ContentLineSpace = 1;//12就是一倍行距 
                    break;
            }
            switch (cb_spaceBeforeContent.SelectedItem.ToString())
            {
                case "0 行": ws.ContentSpaceBefore = 0; break;
                case "0.5 行": ws.ContentSpaceBefore = 0.5; break;
                case "1 行": ws.ContentSpaceBefore = 1; break;
                case "1.5 行": ws.ContentSpaceBefore = 1.5; break;
                case "2 行": ws.ContentSpaceBefore = 2; break;
                case "2.5 行": ws.ContentSpaceBefore = 2.5; break;
                case "3 行": ws.ContentSpaceBefore = 3; break;
                case "3.5 行": ws.ContentSpaceBefore = 3.5; break;
                case "4 行": ws.ContentSpaceBefore = 4; break;
                default: ws.ContentSpaceBefore = 0; break;
            }
            switch (cb_spaceAfterContent.SelectedItem.ToString())
            {
                case "0 行": ws.ContentSpaceAfter = 0; break;
                case "0.5 行": ws.ContentSpaceAfter = 0.5; break;
                case "1 行": ws.ContentSpaceAfter = 1; break;
                case "1.5 行": ws.ContentSpaceAfter = 1.5; break;
                case "2 行": ws.ContentSpaceAfter = 2; break;
                case "2.5 行": ws.ContentSpaceAfter = 2.5; break;
                case "3 行": ws.ContentSpaceAfter = 3; break;
                case "3.5 行": ws.ContentSpaceAfter = 3.5; break;
                case "4 行": ws.ContentSpaceAfter = 4; break;
                default: ws.ContentSpaceAfter = 0; break;
            }
            ws.ContentLeftIndent = Convert.ToDouble(tb_leftIndentContent.Text);
            ws.ContentRightIndent = Convert.ToDouble(tb_rightIndentContent.Text);
            switch (cb_locationContent.SelectedItem.ToString())
            {
                case "居左": ws.AlignmentContent = ParagraphAlignment.Left; break;
                case "居中": ws.AlignmentContent = ParagraphAlignment.Center; break;
                case "居右": ws.AlignmentContent = ParagraphAlignment.Right; break;
                case "左右分散":
                    ws.HeaderAlignment = ParagraphAlignment.Justify;
                    break;
                default: ws.AlignmentContent = ParagraphAlignment.Left; break;
            }
            #endregion
            return ws;
        }

        private void btn_TableBorderColor_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                this.lblTableBorderColor.Tag = colorDialog1.Color;
                this.lblTableBorderColor.BackColor = colorDialog1.Color;
            }
        }

        private void btn_tableShading_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                this.lb_tableShadingColor.Tag = colorDialog1.Color;
                this.lb_tableShadingColor.BackColor = colorDialog1.Color;
            }
        }

        private void cb_locationTable_Click(object sender, EventArgs e)
        {
            this.ucCellAlignment1.Visible = true;


        }

        private void UcCellAlignment1_LabelClick(object sender, EventArgs e)
        {
            Label lbl = (Label)sender;
            wordTestFrm.Model.Enum_CellAlignment ee = (wordTestFrm.Model.Enum_CellAlignment)Convert.ToInt32(lbl.Tag);
            this.cb_locationTable.Text = ee.ToString();
            this.ucCellAlignment1.Visible = false;

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }
        Pen p = new Pen(Color.Black);
        private void lblContent_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            p.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;
            p.DashPattern = new float[] { 3f, 4f, 3f };
            g.DrawLine(p, new Point(0, 0), new Point(lblContent.Width, lblContent.Height));
            g.DrawLine(p, new Point(lblContent.Width, 0), new Point(0, lblContent.Height));
            g.DrawLine(p, new Point(lblContent.Width / 2, 0), new Point(lblContent.Width / 2, lblContent.Height));
            g.DrawLine(p, new Point(0, lblContent.Height / 2), new Point(lblContent.Width, lblContent.Height / 2));
            StringFormat sf = new StringFormat();
            sf.Alignment = StringAlignment.Center;
            sf.LineAlignment = StringAlignment.Center;

            g.DrawString("字体样式", lblContent.Font, new SolidBrush(lblContent.ForeColor), new Rectangle(new Point(0, 0), lblContent.Size), sf);
        }

        int indexGuide = 0;
        private void btn_apply_MouseHover(object sender, EventArgs e)
        {
            lblGuide.Image = Program.guideItems[indexGuide];

            indexGuide++;
            if (Program.guideItems.Count == indexGuide)
            {
                indexGuide = 0;
            }
            lblGuide.Visible = true;

        }

        private void btn_apply_MouseLeave(object sender, EventArgs e)
        {
            lblGuide.Visible = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }
    }
}
