using Aspose.Words;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using wordTestFrm.models;
using wordTestFrm.Properties;

namespace wordTestFrm
{
    public partial class FrmAddWordStyle : Form
    {
        private ComboBox cbbImages = new ComboBox();
        private List<Image> images = new List<Image>();
        string lineStylePath =Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "lineStyleImgs");
        public FrmAddWordStyle()
        {
            InitializeComponent();
            cmb_Header.SelectedIndex = 0;
            cmb_Footer.SelectedIndex = 0;
            cb_hLineSpace.SelectedIndex = 0;
            cb_fLineSpace.SelectedIndex = 0;
            cb_tableLineWidth.SelectedIndex = 0;
            cb_tableSpaceBefore.SelectedIndex = 0;
            cb_tableSpaceAfter.SelectedIndex = 0;
            cb_tableLineSpace.SelectedIndex = 0;
            cb_LineSpaceOne.SelectedIndex = 0;
            cb_spanceBeforeOne.SelectedIndex = 0;
            cb_spanceAfterOne.SelectedIndex = 0;
            cb_LineSpaceTwo.SelectedIndex = 0;
            cb_spanceBeforeTwo.SelectedIndex = 0;
            cb_spanceAfterTwo.SelectedIndex = 0;
            cb_LineSpaceThree.SelectedIndex = 0;
            cb_spanceBeforeThree.SelectedIndex = 0;
            cb_spanceAfterThree.SelectedIndex = 0;
            cb_LineSpaceFour.SelectedIndex = 0;
            cb_spanceBeforeFour.SelectedIndex = 0;
            cb_spanceAfterFour.SelectedIndex = 0;
            cb_LineSpaceFive.SelectedIndex = 0;
            cb_spanceBeforeFive.SelectedIndex = 0;
            cb_spanceAfterFive.SelectedIndex = 0;
            cb_LineSpaceContent.SelectedIndex = 0;
            cb_spanceBeforeContent.SelectedIndex = 0;
            cb_spanceAfterContent.SelectedIndex = 0;
            panel1.Visible = false;

            images.Add(new Bitmap(Path.Combine(lineStylePath, "single.png")));
            images.Add(new Bitmap(Path.Combine(lineStylePath, "DashLargeGap.png")));
            images.Add(new Bitmap(Path.Combine(lineStylePath, "dot.png")));
            images.Add(new Bitmap(Path.Combine(lineStylePath, "DotDash.png")));
            images.Add(new Bitmap(Path.Combine(lineStylePath, "DotDotDash.png")));
            images.Add(new Bitmap(Path.Combine(lineStylePath, "double.png")));
            images.Add(new Bitmap(Path.Combine(lineStylePath, "Triple.png")));
            images.Add(new Bitmap(Path.Combine(lineStylePath, "ThinThickSmallGap.png")));

            cb_tableLineStyle.Items.Add("1");
            cb_tableLineStyle.Items.Add("7");
            cb_tableLineStyle.Items.Add("6");
            cb_tableLineStyle.Items.Add("8");
            cb_tableLineStyle.Items.Add("9");
            cb_tableLineStyle.Items.Add("3");
            cb_tableLineStyle.Items.Add("10");
            cb_tableLineStyle.Items.Add("11");
            cb_tableLineStyle.SelectedIndex = 0;
        }

        private void FrmAddWordStyle_Load(object sender, EventArgs e)
        {
            Color color = Color.Black;
            System.Drawing.Font font_def = new System.Drawing.Font("微软雅黑", 14.25f);

            lb_tableShadingColor.Tag= lblfootColor.Tag = lblhfColor.Tag = lblBorderColor.Tag = lblFontColor.Tag
          = lblLv1Color.Tag = lblLv2Color.Tag = lblLv3Color.Tag = lblLv4Color.Tag = lblLv5Color.Tag = lblLvStrColor.Tag
          = lblLV1FlagColor.Tag = lblLV2FlagColor.Tag = lblLV3FlagColor.Tag = lblLV4FlagColor.Tag = lblLV5FlagColor.Tag = lblLVStrFlagColor.Tag = color;

            lblfootFont.Tag = lblhfFont.Tag = lblFont.Tag
                = lblLv1Font.Tag = lblLv2Font.Tag = lblLv3Font.Tag = lblLv4Font.Tag = lblLv5Font.Tag = lblLvStrFont.Tag
                = lblLv1FlagFont.Tag = lblLv2FlagFont.Tag = lblLv3FlagFont.Tag = lblLv4FlagFont.Tag = lblLv5FlagFont.Tag = lblLvStrFlagFont.Tag
                = font_def;
            tb_header.Text = "页眉";
            tb_footer.Text = "页脚";
            cb_hIsFirstDif.Checked = false;
            cb_hIsParityDif.Checked = false;
            cb_fIsFirstDif.Checked = false;
            cb_fIsParityDif.Checked = false;
            cb_tableIsShading.Checked = false;
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

        private void btnSaveToTxt_Click(object sender, EventArgs e)
        {
            WordStyle ws = new WordStyle();

            #region 页眉
            ws.HeaderName = tb_header.Text;
            ws.HeaderFont = (System.Drawing.Font)lblhfFont.Tag;
            ws.HeaderColor = (Color)lblhfColor.Tag;
            switch (cmb_Header.SelectedItem.ToString())
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
                default:
                    ws.HeaderAlignment = ParagraphAlignment.Center;
                    break;
            }
            ws.HImgPath = imgId;
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
            ws.HIsParityDif = cb_fIsParityDif.Checked;
            #endregion

            #region 页脚
            ws.FooterName = tb_footer.Text;
            ws.FooterFont = (System.Drawing.Font)lblfootFont.Tag;
            ws.FooterColor = (Color)lblfootColor.Tag;
            switch (cmb_Footer.SelectedItem.ToString())
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
                default:
                    ws.FooterAlignment = ParagraphAlignment.Center;
                    break;
            }
            ws.FLeftIndent = Convert.ToDouble(tb_fLeftIndent.Text);
            ws.FRightIndent = Convert.ToDouble(tb_fRightIndent.Text);
            ws.FooterDistance = Convert.ToDouble(tb_footerDistance.Text);
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
            ws.FIsParityDif = cb_fIsParityDif.Checked;
            #endregion

            #region 表格
            ws.TableFont = (System.Drawing.Font)lblFont.Tag;
            ws.TableFontColor = (Color)lblFontColor.Tag;
            ws.TableBorderColor = (Color)lblBorderColor.Tag;
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
            //TODO WWR
            switch (cb_tableLineStyle.SelectedItem.ToString())
            {
                case "1":ws.TableLineStyle = LineStyle.Single; break;
                case "7":ws.TableLineStyle = LineStyle.DashLargeGap; break;
                case "6":ws.TableLineStyle = LineStyle.Dot; break;
                case "8":ws.TableLineStyle = LineStyle.DotDash; break;
                case "9":ws.TableLineStyle = LineStyle.DotDotDash; break;
                case "3":ws.TableLineStyle = LineStyle.Double; break;
                case "10":ws.TableLineStyle = LineStyle.Triple; break;
                case "11":ws.TableLineStyle = LineStyle.ThinThickSmallGap; break;
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
                case "0 行":ws.TableSpaceBefore = 0;break;
                case "0.5 行": ws.TableSpaceBefore = 0.5;break;
                case "1 行": ws.TableSpaceBefore = 1;break;
                case "1.5 行": ws.TableSpaceBefore = 1.5;break;
                case "2 行": ws.TableSpaceBefore = 2;break;
                case "2.5 行": ws.TableSpaceBefore = 2.5;break;
                case "3 行": ws.TableSpaceBefore = 3;break;
                case "3.5 行": ws.TableSpaceBefore = 3.5;break;
                case "4 行": ws.TableSpaceBefore = 4;break;
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
            ws.TitleOneFont = (System.Drawing.Font)lblLv1Font.Tag;
            ws.TitleOneColor = (Color)lblLv1Color.Tag;
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
            switch (cb_spanceBeforeOne.SelectedItem.ToString())
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
            switch (cb_spanceAfterOne.SelectedItem.ToString())
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

            ws.TitleTwoFont = (System.Drawing.Font)lblLv2Font.Tag;
            ws.TitleTwoColor = (Color)lblLv2Color.Tag;
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
            switch (cb_spanceBeforeTwo.SelectedItem.ToString())
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
            switch (cb_spanceAfterTwo.SelectedItem.ToString())
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

            ws.TitleThreeFont = (System.Drawing.Font)lblLv3Font.Tag;
            ws.TitleThreeColor = (Color)lblLv3Color.Tag;
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
            switch (cb_spanceBeforeThree.SelectedItem.ToString())
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
            switch (cb_spanceAfterThree.SelectedItem.ToString())
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

            ws.TitleFourFont = (System.Drawing.Font)lblLv4Font.Tag;
            ws.TitleFourColor = (Color)lblLv4Color.Tag;
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
            switch (cb_spanceBeforeFour.SelectedItem.ToString())
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
            switch (cb_spanceAfterFour.SelectedItem.ToString())
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

            ws.TitleFiveFont = (System.Drawing.Font)lblLv5Font.Tag;
            ws.TitleFiveColor = (Color)lblLv5Color.Tag;
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
            switch (cb_spanceBeforeFive.SelectedItem.ToString())
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
            switch (cb_spanceAfterFive.SelectedItem.ToString())
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

            ws.ContentFont = (System.Drawing.Font)lblLvStrFont.Tag;
            ws.ContentColor = (Color)lblLvStrColor.Tag;
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
            switch (cb_spanceBeforeContent.SelectedItem.ToString())
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
            switch (cb_spanceAfterContent.SelectedItem.ToString())
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
            #endregion

            #region 编号 待用
            ws.FlagOneFont = (System.Drawing.Font)lblLv1FlagFont.Tag;
            ws.FlagOneColor = (Color)lblLV1FlagColor.Tag;
            ws.FlagTwoFont = (System.Drawing.Font)lblLv2FlagFont.Tag;
            ws.FlagTwoColor = (Color)lblLV2FlagColor.Tag;
            ws.FlagThreeFont = (System.Drawing.Font)lblLv3FlagFont.Tag;
            ws.FlagThreeColor = (Color)lblLV3FlagColor.Tag;
            ws.FlagFourFont = (System.Drawing.Font)lblLv4FlagFont.Tag;
            ws.FlagFourColor = (Color)lblLV4FlagColor.Tag;
            ws.FlagFiveFont = (System.Drawing.Font)lblLv5FlagFont.Tag;
            ws.FlagFiveColor = (Color)lblLV5FlagColor.Tag;
            #endregion

            //string res = CommonMethods.saveTxtOfWordStyle(ws, tb_styleName.Text+DateTime.Now.ToString("yyyyMMddmmffsss") + ".txt");
            int res = CommonMethods.saveTxtOfWordStyle(ws, tb_styleName.Text + ".txt");
            if (res > 0)
            {
                MessageBox.Show("添加成功");
            }
            else
            {
                MessageBox.Show("添加失败");
            }
            
            this.Close();
        }

        string imgId = string.Empty;

        private void btn_UploadLogo_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.Filter = "图片文件|*.bmp;*.jpg;*.jpeg;*.png;*.ico";
            if (of.ShowDialog() == DialogResult.OK)
            {
                string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Imgs");
                if (!File.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                path = Path.Combine(path, Guid.NewGuid().ToString("N") + Path.GetExtension(of.SafeFileName));
                if (Path.GetExtension(of.SafeFileName) == ".ico")
                {
                    using (MemoryStream mStream = new MemoryStream())
                    {
                        Icon icon = new Icon(of.FileName);
                        icon.Save(mStream);
                        Image image = Image.FromStream(mStream);
                        lb_Logo.Image = image;
                        File.Copy(of.FileName, path);
                        //以后存入数据库中 从数据库中拿
                        //byte[] byData = new byte[mStream.Length];
                        //mStream.Position = 0;
                        //mStream.Read(byData, 0, byData.Length); 
                        //mStream.Close();
                    }  
                }
                else
                {
                    lb_Logo.Image = new Bitmap(of.FileName);
                    File.Copy(of.FileName,path);
                    //MemoryStream ms = new MemoryStream();
                    //new Bitmap(of.FileName).Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                    //byte[] bytes = ms.GetBuffer();  //byte[]   bytes=   ms.ToArray(); 这两句都可以，至于区别么，下面有解释
                    //ms.Close();
                }
                imgId = path;
            }

        }

        private void cb_tableLineStyle_DrawItem(object sender, DrawItemEventArgs e)
        {
            e.DrawBackground();
            e.Graphics.DrawImage(images[e.Index], 0, e.Bounds.Y, 90, 18);
        }
    }
}
