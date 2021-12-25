using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace wordTestFrm.Model
{
    /// <summary>
    /// 投标类
    /// </summary>
    public class ProjectInfo
    {
        /// <summary>
        /// 投标项目名称
        /// </summary>
        public string ProjectName = string.Empty;

        /// <summary>
        /// 货品名称
        /// </summary>
        public string GoodsName = string.Empty;

        /// <summary>
        /// 投标编号
        /// </summary>
        public string ProjectNum = string.Empty;

        /// <summary>
        /// 银行
        /// </summary>
        public string Bank = string.Empty;

        /// <summary>
        /// 保证金额
        /// </summary>
        public float EarnestMoney = 0.00f;

        /// <summary>
        /// 总金额
        /// </summary>
        public float PJAmount = 0.00f;


        /// <summary>
        /// 有效期
        /// </summary>
        public int DateLimit = 0;

        /// <summary>
        /// 招标单位
        /// </summary>
        public string Tenderee = string.Empty;


        /// <summary>
        /// 招标公司
        /// </summary>
        public string TendereeCompany = string.Empty;

        /// <summary>
        /// 招标招标服务费用
        /// </summary>
        public float ServiceCharge = 0.00f;

        /// <summary>
        /// 公司负责人
        /// </summary>
        public PerInfo manager = null;

        /// <summary>
        /// 投标授权人
        /// </summary>
        public PerInfo donor = null;

        /// <summary>
        /// 投标公司
        /// </summary>
        public Company company = null;

        /// <summary>
        /// 投标人员信息
        /// </summary>
        /// <param name="name"></param>
        /// <param name="job"></param>
        /// <param name="idNumber"></param>
        /// <param name="pic"></param>
        public static PerInfo SetPersonInfo(string name, string job, string idNumber = "", byte[] pic=null)
        {
            PerInfo per = new PerInfo() {
                perName=name,
                job=job,
                idNumber=idNumber,
                idCardPic=pic
            };
            return per;
        }

        /// <summary>
        /// 公司信息
        /// </summary>
        /// <param name="name"></param>
        /// <param name="adr"></param>
        /// <param name="phone"></param>
        /// <param name="fax"></param>
        /// <param name="mailNum"></param>
        /// <returns></returns>
        public static Company SetCompanyInfo(string name,string adr,string phone,string fax,string mailNum)
        {
            Company company = new Company()
            {
                companyName=name,
                address=adr,
                phone=phone,
                fax=fax,
                mailNum=mailNum
            };
            return company;
        }
    }
}
