using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Aspose.Words;
using wordTestFrm.Model;

namespace wordTestFrm
{
    public partial class FrmWordCreateTest : Form
    {
        ProjectInfo project = new ProjectInfo();
        public FrmWordCreateTest()
        {
            InitializeComponent();
            cbxCompanySize.SelectedIndex = 1;
        }

        private void btnCreateWord_Click(object sender, EventArgs e)
        {
            SetProjectInfo();

            FolderBrowserDialog of = new FolderBrowserDialog();
            of.SelectedPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            if (of.ShowDialog() == DialogResult.OK)
            {
                string dirPath = of.SelectedPath;
                string fileName = txtWordName.Text + ".docx";
                string fullPath = Path.Combine(dirPath, fileName);
                //File.Create(fullPath);
                Document docMain = new Document();
         
                DirectoryInfo directoryInfo = new DirectoryInfo(dirPath);
                DirectoryInfo[] dirItems = directoryInfo.GetDirectories().OrderBy(item=>int.Parse(item.Name.Substring(0,item.Name.IndexOf('-')))).ToArray();
                
                Dictionary<string, Document> dict = new Dictionary<string, Document>();
                for (int i = 0; i < dirItems.Length; i++)
                {
                    DirectoryInfo info = dirItems[i];
                    FileInfo[] items = info.GetFiles();
                    for (int a = 0; a < items.Length; a++)
                    {
                        try
                        {
                            Document tmpDoc = new Document(items[a].FullName);
                           
                            SetBookMarkVal(tmpDoc);
                            
                            docMain.AppendDocument(tmpDoc, ImportFormatMode.UseDestinationStyles);
                        }
                        catch (Aspose.Words.UnsupportedFileFormatException ex)
                        {

                            continue;
                        }
                        
                    }
                }

                NodeCollection nodes = docMain.GetChildNodes(NodeType.Paragraph, true);
                for (int h = 0; h < nodes.Count; h++)
                {
                    Node node = nodes[h];
                    if (node.NodeType == NodeType.Paragraph)
                    {
                        Paragraph p = (Paragraph)node;
                        if (p.ParagraphFormat.IsHeading)
                        {
                            if (p.ParagraphFormat.OutlineLevel == OutlineLevel.Level1)
                            {
                                for (int k = 0; k < p.Runs.Count; k++)
                                {
                                    Run run = p.Runs[k];
                                    run.Font.Name = "微软雅黑";
                                    run.Font.Bold = true;
                                    run.Font.Size = 20;
                                }
                                p.ParagraphFormat.Style.Font.Name = "微软雅黑";
                                p.ParagraphFormat.Style.Font.Bold = true;
                                p.ParagraphFormat.Style.Font.Size = 20;
                            }
                            else if (p.ParagraphFormat.OutlineLevel == OutlineLevel.Level2)
                            {
                                for (int k = 0; k < p.Runs.Count; k++)
                                {
                                    Run run = p.Runs[k];
                                    run.Font.Name = "微软雅黑";
                                    run.Font.Bold = true;
                                    run.Font.Size = 16;
                                }
                              
                                p.ParagraphFormat.Style.Font.Name = "微软雅黑";
                                p.ParagraphFormat.Style.Font.Bold = true;
                                p.ParagraphFormat.Style.Font.Size = 16;
                            }
                        }
                        else
                        {
                            for (int k = 0; k < p.Runs.Count; k++)
                            {
                                Run run = p.Runs[k];
                                run.Font.Name = "微软雅黑";
                                run.Font.Bold = false;
                                run.Font.Size = 12;
                              
                            }
                            p.ParagraphFormat.Style.Font.Name = "微软雅黑";

                            p.ParagraphFormat.Style.Font.Bold = false;
                            p.ParagraphFormat.Style.Font.Size = 12;
                        }
                    }
                }
                docMain.RemoveChild(docMain.FirstChild);
                //word 版式
                Aspose.Words.Settings.CompatibilityOptions compatibility = docMain.CompatibilityOptions;
                compatibility.UlTrailSpace = true;//设置尾部空格下划线 重要
                fullPath = Path.Combine(dirPath, txtWordName.Text + "AsposeWord.docx");
                docMain.Save(fullPath, Aspose.Words.Saving.SaveOptions.CreateSaveOptions(SaveFormat.Docx));
                Process.Start(fullPath);


            }
        }

        public string retSavePath(string filePath)
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

        private void txtSetBookMark_Click(object sender, EventArgs e)
        {
            SetProjectInfo();


            OpenFileDialog of = new OpenFileDialog();
            of.InitialDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            if (of.ShowDialog() == DialogResult.OK)
            {
                string realName = Path.GetFileName(of.FileName);
                string CopyPath = this.retSavePath(of.FileName);
                Document doc = new Aspose.Words.Document(CopyPath);
               
                SetBookMarkVal(doc);
                string savePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Path.GetFileNameWithoutExtension(realName) + "_AsposeWord.docx");
                doc.Save(savePath, SaveFormat.Docx);
                Process.Start(savePath);
            }
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



        public void SetProjectInfo()
        {
            project.ProjectName = txtProjectName.Text;
            project.ProjectNum = txtPJNum.Text;
            project.GoodsName = txtGoodsName.Text;
            project.manager = ProjectInfo.SetPersonInfo(txtManager.Text, txtMangerJob.Text);
            project.donor = ProjectInfo.SetPersonInfo(txtDonor.Text, txtDonorName.Text);
            project.company = ProjectInfo.SetCompanyInfo(txtCompany.Text, txtCPAdr.Text, txtCPPhone.Text,
                txtCPFax.Text, txtMailNum.Text);
            project.PJAmount = float.Parse(txtPjAmount.Text);
            project.Bank = txtBank.Text;
            project.EarnestMoney = float.Parse(txtEarnestMoney.Text);
            project.Tenderee = txtTenderee.Text;
            project.TendereeCompany = txtTenderCompany.Text;
            project.ServiceCharge = float.Parse(txtServiceCharge.Text);


            project.DateLimit = int.Parse(txtDateLimit.Text);
            project.company.buildDate = txtBuildDate.Text;
            project.company.CreditReferenceBank = txtCreditReferenceBank.Text;
            project.company.BankAddress = txtBankAdr.Text;
            project.company.netValue = float.Parse(txtNetValue.Text);
            project.company.PaidinCapital = float.Parse(txtPaidinCapital.Text);
            project.company.fixedAssets = float.Parse(txtFixedAssets.Text);
            project.company.accruedAssets = float.Parse(txtAccruedAssets.Text);
            project.company.longtermDebt = float.Parse(txtLongtermDebt.Text);
            project.company.floatingDebt = float.Parse(txtFloatingDebt.Text);
            project.company.CompanySize = cbxCompanySize.SelectedItem.ToString();
            project.company.AccountBank = txtAccountBank.Text;
            project.company.AccountNum = txtAccountNum.Text;
            project.company.RefundBank = txtRefundBank.Text;
            project.company.refundAccountNum = txtRefundAccount.Text;

            project.PJAmount = float.Parse(txtPjAmount.Text);
            project.Bank = txtBank.Text;

            project.EarnestMoney = float.Parse(txtEarnestMoney.Text);
            project.DateLimit = int.Parse(txtDateLimit.Text);
        }

        /// <summary>
        /// 设置标签值
        /// </summary>
        /// <param name="doc"></param>
        public void SetBookMarkVal(Document doc)
        {
            BookmarkCollection bookmarks = doc.Range.Bookmarks;
            for (int i = 0; i < bookmarks.Count; i++)
            {
                Bookmark bk = bookmarks[i];
                string name = bk.Name;

                string bkVal = this.SetVal(name);
                if (!string.IsNullOrEmpty(bkVal))
                {
                    bk.Text = bkVal;

                }
            }
        }



        /// <summary>
        /// 金额转为大写金额
        /// </summary>
        /// <param name="LowerMoney"></param>
        /// <returns></returns>
        public string MoneyToChinese(string LowerMoney)
        {
            string functionReturnValue = null;
            bool IsNegative = false; // 是否是负数
            if (LowerMoney.Trim().Substring(0, 1) == "-")
            {
                // 是负数则先转为正数
                LowerMoney = LowerMoney.Trim().Remove(0, 1);
                IsNegative = true;
            }
            string strLower = null;
            string strUpart = null;
            string strUpper = null;
            int iTemp = 0;
            // 保留两位小数 123.489→123.49　　123.4→123.4
            LowerMoney = Math.Round(double.Parse(LowerMoney), 2).ToString();
            if (LowerMoney.IndexOf(".") > 0)
            {
                if (LowerMoney.IndexOf(".") == LowerMoney.Length - 2)
                {
                    LowerMoney = LowerMoney + "0";
                }
            }
            else
            {
                LowerMoney = LowerMoney + ".00";
            }
            strLower = LowerMoney;
            iTemp = 1;
            strUpper = "";
            while (iTemp <= strLower.Length)
            {
                switch (strLower.Substring(strLower.Length - iTemp, 1))
                {
                    case ".":
                        strUpart = "圆";
                        break;
                    case "0":
                        strUpart = "零";
                        break;
                    case "1":
                        strUpart = "壹";
                        break;
                    case "2":
                        strUpart = "贰";
                        break;
                    case "3":
                        strUpart = "叁";
                        break;
                    case "4":
                        strUpart = "肆";
                        break;
                    case "5":
                        strUpart = "伍";
                        break;
                    case "6":
                        strUpart = "陆";
                        break;
                    case "7":
                        strUpart = "柒";
                        break;
                    case "8":
                        strUpart = "捌";
                        break;
                    case "9":
                        strUpart = "玖";
                        break;
                }

                switch (iTemp)
                {
                    case 1:
                        strUpart = strUpart + "分";
                        break;
                    case 2:
                        strUpart = strUpart + "角";
                        break;
                    case 3:
                        strUpart = strUpart + "";
                        break;
                    case 4:
                        strUpart = strUpart + "";
                        break;
                    case 5:
                        strUpart = strUpart + "拾";
                        break;
                    case 6:
                        strUpart = strUpart + "佰";
                        break;
                    case 7:
                        strUpart = strUpart + "仟";
                        break;
                    case 8:
                        strUpart = strUpart + "万";
                        break;
                    case 9:
                        strUpart = strUpart + "拾";
                        break;
                    case 10:
                        strUpart = strUpart + "佰";
                        break;
                    case 11:
                        strUpart = strUpart + "仟";
                        break;
                    case 12:
                        strUpart = strUpart + "亿";
                        break;
                    case 13:
                        strUpart = strUpart + "拾";
                        break;
                    case 14:
                        strUpart = strUpart + "佰";
                        break;
                    case 15:
                        strUpart = strUpart + "仟";
                        break;
                    case 16:
                        strUpart = strUpart + "万";
                        break;
                    default:
                        strUpart = strUpart + "";
                        break;
                }

                strUpper = strUpart + strUpper;
                iTemp = iTemp + 1;
            }

            strUpper = strUpper.Replace("零拾", "零");
            strUpper = strUpper.Replace("零佰", "零");
            strUpper = strUpper.Replace("零仟", "零");
            strUpper = strUpper.Replace("零零零", "零");
            strUpper = strUpper.Replace("零零", "零");
            strUpper = strUpper.Replace("零角零分", "整");
            strUpper = strUpper.Replace("零分", "整");
            strUpper = strUpper.Replace("零角", "零");
            strUpper = strUpper.Replace("零亿零万零圆", "亿圆");
            strUpper = strUpper.Replace("亿零万零圆", "亿圆");
            strUpper = strUpper.Replace("零亿零万", "亿");
            strUpper = strUpper.Replace("零万零圆", "万圆");
            strUpper = strUpper.Replace("零亿", "亿");
            strUpper = strUpper.Replace("零万", "万");
            strUpper = strUpper.Replace("零圆", "圆");
            strUpper = strUpper.Replace("零零", "零");

            // 对壹圆以下的金额的处理
            if (strUpper.Substring(0, 1) == "圆")
            {
                strUpper = strUpper.Substring(1, strUpper.Length - 1);
            }
            if (strUpper.Substring(0, 1) == "零")
            {
                strUpper = strUpper.Substring(1, strUpper.Length - 1);
            }
            if (strUpper.Substring(0, 1) == "角")
            {
                strUpper = strUpper.Substring(1, strUpper.Length - 1);
            }
            if (strUpper.Substring(0, 1) == "分")
            {
                strUpper = strUpper.Substring(1, strUpper.Length - 1);
            }
            if (strUpper.Substring(0, 1) == "整")
            {
                strUpper = "零圆整";
            }
            functionReturnValue = strUpper;

            if (IsNegative == true)
            {
                return "负" + functionReturnValue;
            }
            else
            {
                return functionReturnValue;
            }
        }



        public string SetVal(string name)
        {
            string val = string.Empty;
            string[] nameItems = name.Split('_');
            if(nameItems.Length>=2)
            {
                switch(nameItems[1].ToUpper())
                {
                    case "COMPANY":
                        val = this.project.company.companyName;
                        break;
                    case "CPADDRESS":
                        val = this.project.company.address;
                        break;
                    case "CPFAX":
                        val = this.project.company.fax;
                        break;
                    case "CPMAILNUM":
                        val = this.project.company.mailNum;
                        break;
                    case "CPPHONE":
                        val = this.project.company.phone;
                        break;
                    case "PJNAME":
                        val = this.project.ProjectName;
                        break;
                    case "PJNUM":
                        val = this.project.ProjectNum;
                        break;
                    case "DONOR":
                        val = this.project.donor.perName;
                        break;
                    case "DONORJOB":
                        val = this.project.donor.job;
                        break;
                    case "MANAGER":
                        val = this.project.manager.perName;
                        break;
                    case "MANAGERJOB":
                        val = this.project.manager.job;
                        break;
                    case "BANK":
                        val = this.project.Bank;
                        break;
                    case "EARNESTMONEY":
                        val = string.Format("{0:N2}", this.project.EarnestMoney); 
                        break;
                    case "PJAMOUNT":
                        val = string.Format("{0:N2}", this.project.PJAmount );  //;
                        break;
                    case "PJAMOUNTCHINESE":
                        val = string.Format("{0:N2}", this.project.PJAmount);
                        val = this.MoneyToChinese(val);
                        break;
                    case "DATELIMIT":
                        val = this.project.DateLimit.ToString();
                        break;
                    case "BUILDDATE":
                        val = this.project.company.buildDate;
                        break;

                    case "PAIDINCAPITAL":
                        val = string.Format("{0:N2}", this.project.company.PaidinCapital);
                        break;
                    case "FIXEDASSETS":
                        val = string.Format("{0:N2}", this.project.company.fixedAssets);  //;
                        break;
                    case "ACCRUEDASSETS":
                        val = string.Format("{0:N2}", this.project.company.accruedAssets);
                        break;
                    case "LONGTERMDEBT":
                        val = string.Format("{0:N2}", this.project.company.longtermDebt);
                        break;
                    case "FLOATINGDEBT":
                        val = string.Format("{0:N2}", this.project.company.floatingDebt);  //;
                        break;
                    case "NETVALUE":
                        val = string.Format("{0:N2}", this.project.company.netValue);
                        break;
                    case "CREDITREFERENCEBANK":
                        val = this.project.company.CreditReferenceBank;
                        break;
                    case "BANKADDRESS":
                        val = this.project.company.BankAddress;
                        break;
                    case "TENDEREE":
                        val = this.project.Tenderee;
                        break;
                    case "COMPANYSIZE":
                        val = this.project.company.CompanySize;
                        break;
                    case "TENDERCOMPANY":
                        val = this.project.TendereeCompany;
                        break;
                    case "SERVICECHARGE":
                        val = string.Format("{0:N2}", this.project.ServiceCharge); 
                        break;
                    case "GOODSNAME":
                        val = this.project.GoodsName;
                        break;

                    case "ACCOUNTBANK":
                        val = this.project.company.AccountBank;
                        break;

                    case "ACCOUNTNUM":
                        val = this.project.company.AccountNum;
                        break;

                    case "REFUNDACCOUNT":
                        val = this.project.company.refundAccountNum;
                        break;

                    case "REFUNDBANK":
                        val = this.project.company.RefundBank;
                        break;


                }
            }
            return val;
        }

        private void eventLog1_EntryWritten(object sender, EntryWrittenEventArgs e)
        {

        }
    }
}
