using Aspose.Words;
using Org.BouncyCastle.Bcpg.OpenPgp;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace wordTestFrm.models
{
    /// <summary>
    /// word样式
    /// </summary>
    public class WordStyle
    {
        #region 页面格式
        /// <summary>
        /// 页面格式 A4、A5等
        /// </summary>
        public PaperSize PageType { get; set; }
        /// <summary>
        /// 页面方向
        /// </summary>
        public Orientation PageDirection { get; set; }
        /// <summary>
        /// 页边距上
        /// </summary>
        public double TopMargin { get; set; }
        /// <summary>
        /// 页边距下
        /// </summary>
        public double BottomMargin { get; set; }
        /// <summary>
        /// 页边距左
        /// </summary>
        public double LeftMargin { get; set; }
        /// <summary>
        /// 页边距右
        /// </summary>
        public double RightMargin { get; set; } 
        #endregion

        #region 页眉
        /// <summary>
        /// 页眉名称
        /// </summary>
        public string HeaderName { get; set; }
        /// <summary>
        /// 页眉字体
        /// </summary>
        public System.Drawing.Font HeaderFont { get; set; }
        /// <summary>
        /// 页眉字体颜色
        /// </summary>
        public Color HeaderColor { get; set; }
        /// <summary>
        /// 页眉居左/居中/居右
        /// </summary>
        public ParagraphAlignment HeaderAlignment { get; set; }
        /// <summary>
        /// 页眉左缩进
        /// </summary>
        public double HLeftIndent { get; set; }
        /// <summary>
        /// 页眉右缩进
        /// </summary>
        public double HRightIndent { get; set; }
        /// <summary>
        /// 页眉是否首页不同
        /// </summary>
        public bool HIsFirstDif { get; set; }
        /// <summary>
        /// 页眉是否奇偶不同
        /// </summary>
        public bool HIsParityDif { get; set; }
        /// <summary>
        /// 页眉行距
        /// </summary>
        public double HLineSpace { get; set; }
        /// <summary>
        /// 页眉顶端距离
        /// </summary>
        public double HeaderDistance { get; set; }
        /// <summary>
        /// 页眉插入图片的路径
        /// </summary>
        public string HImgPath { get; set; }
        /// <summary>
        /// 图片数据
        /// </summary>
        public byte[] ImgData { get; set; }
        #endregion

        #region 页脚
        /// <summary>
        /// 页脚名称
        /// </summary>
        public string FooterName { get; set; }
        /// <summary>
        /// 页脚字体
        /// </summary>
        public System.Drawing.Font FooterFont { get; set; }
        /// <summary>
        /// 页脚字体颜色
        /// </summary>
        public Color FooterColor { get; set; }
        /// <summary>
        /// 页脚居左/居中/居右
        /// </summary>
        public ParagraphAlignment FooterAlignment { get; set; }
        /// <summary>
        /// 页脚左缩进
        /// </summary>
        public double FLeftIndent { get; set; }
        /// <summary>
        /// 页脚右缩进
        /// </summary>
        public double FRightIndent { get; set; }
        /// <summary>
        /// 页脚是否首页不同
        /// </summary>
        public bool FIsFirstDif { get; set; }
        /// <summary>
        /// 页脚是否奇偶不同
        /// </summary>
        public bool FIsParityDif { get; set; }
        /// <summary>
        /// 页脚行距
        /// </summary>
        public double FLineSpace { get; set; }
        /// <summary>
        /// 页脚底部距离
        /// </summary>
        public double FooterDistance { get; set; }
        #endregion

        #region 表格
        /// <summary>
        /// 表格内容字体颜色
        /// </summary>
        public Color TableFontColor { get; set; }
        /// <summary>
        /// 表格内容字体
        /// </summary>
        public System.Drawing.Font TableFont { get; set; }
        /// <summary>
        /// 表格边框颜色
        /// </summary>
        public Color TableBorderColor { get; set; }
        /// <summary>
        /// 表格行间距
        /// </summary>
        public double TableLineSpace { get; set; }
        /// <summary>
        /// 表格线样式
        /// </summary>
        public LineStyle TableLineStyle { get; set; }
        /// <summary>
        /// 表格线宽度
        /// </summary>
        public double TableLineWidth { get; set; }
        /// <summary>
        /// 表格段前
        /// </summary>
        public double TableSpaceBefore { get; set; }
        /// <summary>
        /// 表格段后
        /// </summary>
        public double TableSpaceAfter { get; set; }
        /// <summary>
        /// 表格左缩进
        /// </summary>
        public double TableLeftIndent { get; set; }
        /// <summary>
        /// 表格右缩进
        /// </summary>
        public double TableRightIndent { get; set; }
        /// <summary>
        /// 表格是否首行加底纹
        /// </summary>
        public bool TableIsShading { get; set; }
        /// <summary>
        /// 底纹颜色
        /// </summary>
        public Color TableShadingColor { get; set; }

        /// <summary>
        /// 单元格位置
        /// </summary>
        public Model.Enum_CellAlignment TableAlignment { get; set; }
        #endregion

        #region 标题和内容
        /// <summary>
        /// 标题1字体
        /// </summary>
        public System.Drawing.Font TitleOneFont { get; set; }
        /// <summary>
        /// 标题1字体颜色
        /// </summary>
        public Color TitleOneColor { get; set; }
        /// <summary>
        /// 标题1行间距
        /// </summary>
        public double TitleOneLineSpace { get; set; }
        /// <summary>
        /// 标题1段前
        /// </summary>
        public double TitleOneSpaceBefore { get; set; }
        /// <summary>
        /// 标题1段后
        /// </summary>
        public double TitleOneSpaceAfter { get; set; }
        /// <summary>
        /// 标题1左缩进
        /// </summary>
        public double TitleOneLeftIndent { get; set; }
        /// <summary>
        /// 标题1右缩进
        /// </summary>
        public double TitleOneRightIndent { get; set; }
        /// <summary>
        /// 标题1位置
        /// </summary>
        public ParagraphAlignment AlignmentOne { get; set; }
        /// <summary>
        /// 标题2字体
        /// </summary>
        public System.Drawing.Font TitleTwoFont { get; set; }
        /// <summary>
        /// 标题2字体颜色
        /// </summary>
        public Color TitleTwoColor { get; set; }
        /// <summary>
        /// 标题2行间距
        /// </summary>
        public double TitleTwoLineSpace { get; set; }
        /// <summary>
        /// 标题2段前
        /// </summary>
        public double TitleTwoSpaceBefore { get; set; }
        /// <summary>
        /// 标题2段后
        /// </summary>
        public double TitleTwoSpaceAfter { get; set; }
        /// <summary>
        /// 标题2左缩进
        /// </summary>
        public double TitleTwoLeftIndent { get; set; }
        /// <summary>
        /// 标题2右缩进
        /// </summary>
        public double TitleTwoRightIndent { get; set; }
        /// <summary>
        /// 标题2位置
        /// </summary>
        public ParagraphAlignment AlignmentTwo { get; set; }
        /// <summary>
        /// 标题3字体
        /// </summary>
        public System.Drawing.Font TitleThreeFont { get; set; }
        /// <summary>
        /// 标题3字体颜色
        /// </summary>
        public Color TitleThreeColor { get; set; }
        /// <summary>
        /// 标题3行间距
        /// </summary>
        public double TitleThreeLineSpace { get; set; }
        /// <summary>
        /// 标题3段前
        /// </summary>
        public double TitleThreeSpaceBefore { get; set; }
        /// <summary>
        /// 标题3段后
        /// </summary>
        public double TitleThreeSpaceAfter { get; set; }
        /// <summary>
        /// 标题3左缩进
        /// </summary>
        public double TitleThreeLeftIndent { get; set; }
        /// <summary>
        /// 标题3右缩进
        /// </summary>
        public double TitleThreeRightIndent { get; set; }
        /// <summary>
        /// 标题3位置
        /// </summary>
        public ParagraphAlignment AlignmentThree { get; set; }
        /// <summary>
        /// 标题4字体
        /// </summary>
        public System.Drawing.Font TitleFourFont { get; set; }
        /// <summary>
        /// 标题4字体颜色
        /// </summary>
        public Color TitleFourColor { get; set; }
        /// <summary>
        /// 标题4行间距
        /// </summary>
        public double TitleFourLineSpace { get; set; }
        /// <summary>
        /// 标题4段前
        /// </summary>
        public double TitleFourSpaceBefore { get; set; }
        /// <summary>
        /// 标题4段后
        /// </summary>
        public double TitleFourSpaceAfter { get; set; }
        /// <summary>
        /// 标题4左缩进
        /// </summary>
        public double TitleFourLeftIndent { get; set; }
        /// <summary>
        /// 标题4右缩进
        /// </summary>
        public double TitleFourRightIndent { get; set; }
        /// <summary>
        /// 标题4位置
        /// </summary>
        public ParagraphAlignment AlignmentFour { get; set; }
        /// <summary>
        /// 标题5字体
        /// </summary>
        public System.Drawing.Font TitleFiveFont { get; set; }
        /// <summary>
        /// 标题5字体颜色
        /// </summary>
        public Color TitleFiveColor { get; set; }
        /// <summary>
        /// 标题5行间距
        /// </summary>
        public double TitleFiveLineSpace { get; set; }
        /// <summary>
        /// 标题5段前
        /// </summary>
        public double TitleFiveSpaceBefore { get; set; }
        /// <summary>
        /// 标题5段后
        /// </summary>
        public double TitleFiveSpaceAfter { get; set; }
        /// <summary>
        /// 标题5左缩进
        /// </summary>
        public double TitleFiveLeftIndent { get; set; }
        /// <summary>
        /// 标题5右缩进
        /// </summary>
        public double TitleFiveRightIndent { get; set; }
        /// <summary>
        /// 标题5位置
        /// </summary>
        public ParagraphAlignment AlignmentFive { get; set; }
        /// <summary>
        /// 内容字体
        /// </summary>
        public System.Drawing.Font ContentFont { get; set; }
        /// <summary>
        /// 内容字体颜色
        /// </summary>
        public Color ContentColor { get; set; }
        /// <summary>
        /// 内容行间距
        /// </summary>
        public double ContentLineSpace { get; set; }
        /// <summary>
        /// 内容段前
        /// </summary>
        public double ContentSpaceBefore { get; set; }
        /// <summary>
        /// 内容段后
        /// </summary>
        public double ContentSpaceAfter { get; set; }
        /// <summary>
        /// 内容左缩进
        /// </summary>
        public double ContentLeftIndent { get; set; }
        /// <summary>
        /// 内容右缩进
        /// </summary>
        public double ContentRightIndent { get; set; }
        /// <summary>
        /// 内容位置
        /// </summary>
        public ParagraphAlignment AlignmentContent { get; set; }
        #endregion

        #region 编号
        /// <summary>
        /// 编号/符号1字体
        /// </summary>
        public System.Drawing.Font FlagOneFont { get; set; }
        /// <summary>
        /// 编号/符号 1 字体颜色
        /// </summary>
        public Color FlagOneColor { get; set; }
        /// <summary>
        /// 编号/符号 2 字体
        /// </summary>
        public System.Drawing.Font FlagTwoFont { get; set; }
        /// <summary>
        /// 编号/符号 2 字体颜色
        /// </summary>
        public Color FlagTwoColor { get; set; }
        /// <summary>
        /// 编号/符号 3 字体
        /// </summary>
        public System.Drawing.Font FlagThreeFont { get; set; }
        /// <summary>
        /// 编号/符号 3 字体颜色
        /// </summary>
        public Color FlagThreeColor { get; set; }
        /// <summary>
        /// 编号/符号 4 字体
        /// </summary>
        public System.Drawing.Font FlagFourFont { get; set; }
        /// <summary>
        /// 编号/符号 4 字体颜色
        /// </summary>
        public Color FlagFourColor { get; set; }
        /// <summary>
        /// 编号/符号 5 字体
        /// </summary>
        public System.Drawing.Font FlagFiveFont { get; set; }
        /// <summary>
        /// 编号/符号 5 字体颜色
        /// </summary>
        public Color FlagFiveColor { get; set; } 
        #endregion
    }
}
