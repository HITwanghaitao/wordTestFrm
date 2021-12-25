using Aspose.Words;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace wordTestFrm
{
    public partial class FrmWordStruct : Form
    {
        public Document doc = null;
        public Paragraph globalP = null;
        public FrmWordStruct()
        {
            InitializeComponent();
        }

        public void InitTreeView(bool isHeading)
        {
            
            NodeCollection nodes = doc.GetChildNodes(NodeType.Paragraph, true);
            TreeNode trNode = new TreeNode();
            trNode.Text = "word文档";

            TreeNode trPrev = new TreeNode();
            for (int i = 0; i < nodes.Count; i++)
            {
                Paragraph p = (Paragraph)nodes[i];
                
                TreeNode trNode_Current = new TreeNode();
                trNode_Current.Text = p.GetText();
                trNode_Current.Tag = p;
                switch (p.ParagraphFormat.OutlineLevel)
                {
                    case OutlineLevel.Level1:
                        trNode.Nodes.Add(trNode_Current);
                        trPrev = trNode_Current;
                        break;
                    case OutlineLevel.BodyText:
                        if (isHeading)
                        {
                            break ;
                        }
                        TreeNode treeNodeParent = this.GetParentNode(trNode_Current, trPrev);
                        treeNodeParent.Nodes.Add(trNode_Current);
                        trPrev = trNode_Current;
                        break;
                    default:
                        treeNodeParent = this.GetParentNode(trNode_Current, trPrev);
                        treeNodeParent.Nodes.Add(trNode_Current);
                        trPrev = trNode_Current;
                        break;
                }
                
            }
            tvStruct.Nodes.Add(trNode);

        }

        /// <summary>
        /// 获取节点的上级目录
        /// </summary>
        /// <param name="current"></param>
        /// <param name="prev"></param>
        /// <returns></returns>
        public TreeNode GetParentNode(TreeNode current,TreeNode prev)
        {
            Paragraph p = (Paragraph)current.Tag;
            Paragraph pPrev = (Paragraph)prev.Tag;
            if (pPrev == null) return prev;
            else if(p.ParagraphFormat.OutlineLevel>pPrev.ParagraphFormat.OutlineLevel)
            {
                return prev;
            }
            else
            {
               return GetParentNode(current, prev.Parent);
            }
        }

        private void tvStruct_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {

           
        }

        private void tvStruct_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            TreeNode node = e.Node;
            if (node.Tag != null)
            {
                Paragraph p = (Paragraph)node.Tag;
                globalP = p;
                txtContent.Text = p.GetText();
                string content = string.Empty;
                content += p.ParagraphFormat.OutlineLevel.ToString() + "\r\n";
                for (int i = 0; i < p.Runs.Count; i++)
                {
                    Run run = p.Runs[i];
                    content += run.Font.Name + " " + run.Font.Size + " " + run.Font.Color.ToString() + "\r\n";
                }
                //lblStyle.Text = content;
            }
        }

        private void chkOnlyHeading_CheckedChanged(object sender, EventArgs e)
        {
            tvStruct.Nodes.Clear();
            if(chkOnlyHeading.Checked)
            {
                InitTreeView(true);
            }
            else
            {
                InitTreeView(false);
            }
        }

        private void lblColor_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                Label lbl = ((Label)sender);
                lbl.Tag = colorDialog1.Color;
                lbl.BackColor = colorDialog1.Color;
            }
        }

        private void lblFont_Click(object sender, EventArgs e)
        {
            if (fontDialog1.ShowDialog() == DialogResult.OK)
            {
                Label lbl = ((Label)sender);
                lbl.Tag = fontDialog1.Font;
                lbl.Text = string.Format("字体:{0}  字号: {1}", fontDialog1.Font.Name, fontDialog1.Font.Size);
            }
        }

        private void btn_Apply_Click(object sender, EventArgs e)
        {
            System.Drawing.Font f = (System.Drawing.Font)lblFont.Tag;
            Color cFont = (Color)lblColor.Tag;
            foreach (Run item in globalP.Runs)
            {
                if (item == null) continue;
                item.Font.Size = f.Size;
                item.Font.Color = cFont;
                item.Font.Bold = f.Bold;
                item.Font.Italic = f.Italic;
                item.Font.Name = f.Name;
            }
        }

        private void FrmWordStruct_Load(object sender, EventArgs e)
        {
            Color color = Color.Black;
            System.Drawing.Font font_def = new System.Drawing.Font("微软雅黑", 14.25f);
            lblFont.Tag = font_def;
            lblColor.Tag = color;
        }
    }
}
