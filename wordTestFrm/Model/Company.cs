using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace wordTestFrm.Model
{
    /// <summary>
    /// 投标公司
    /// </summary>
    public class Company
    {
        /// <summary>
        /// 公司id
        /// </summary>
        public string companyId;
        /// <summary>
        /// 公司名称
        /// </summary>
        public string companyName = string.Empty;
        /// <summary>
        /// 地址
        /// </summary>
        public string address = string.Empty;

        /// <summary>
        /// 电话
        /// </summary>
        public string phone = string.Empty;

        /// <summary>
        /// 传真
        /// </summary>
        public string fax = string.Empty;

        /// <summary>
        /// 邮编
        /// </summary>
        public string mailNum = string.Empty;

        public string buildDate = string.Empty;

        /// <summary>
        /// 实收资本
        /// </summary>
        public float PaidinCapital = 5000000.00f;

        /// <summary>
        /// 固定资产
        /// </summary>
        public float fixedAssets = 14783826.66f;

        /// <summary>
        /// 流动资产
        /// </summary>
        public float accruedAssets = 59077694.15f;

        /// <summary>
        /// 长期债务
        /// </summary>
        public float longtermDebt = 474183.60f;

        /// <summary>
        /// 流动债务
        /// </summary>
        public float floatingDebt= 22904910.14f;

        /// <summary>
        /// 净值
        /// </summary>
        public float netValue = 73861520.81f;

        /// <summary>
        /// 征信银行
        /// </summary>
        public string CreditReferenceBank = string.Empty;

        /// <summary>
        /// 银行地址
        /// </summary>
        public string BankAddress = string.Empty;

        /// <summary>
        /// 公司规模
        /// </summary>
        public string CompanySize = string.Empty;

        /// <summary>
        /// 开户银行
        /// </summary>
        public string AccountBank = string.Empty;

        public string AccountNum = string.Empty;

        /// <summary>
        /// 退款银行
        /// </summary>
        public string RefundBank = string.Empty;

        public string refundAccountNum = string.Empty;
    }
}
