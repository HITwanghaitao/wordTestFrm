using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace wordTestFrm.Model
{
    /// <summary>
    /// 投标相关人员
    /// </summary>
    public class PerInfo
    {
        /// <summary>
        /// 人员id
        /// </summary>
        public string perId;

        /// <summary>
        /// 姓名
        /// </summary>
        public string perName = string.Empty;

        /// <summary>
        /// 性别
        /// </summary>
        public string perSex = string.Empty;

        /// <summary>
        /// 职务
        /// </summary>
        public string job = string.Empty;

        /// <summary>
        /// 身份证
        /// </summary>
        public string idNumber = string.Empty;

        /// <summary>
        /// 身份证照片
        /// </summary>
        public byte[] idCardPic = null;

        /// <summary>
        /// 个人简介
        /// </summary>
        public string introduce = string.Empty;

        /// <summary>
        /// 所属公司
        /// </summary>
        public string companyId = string.Empty;

        /// <summary>
        /// 学历
        /// </summary>
        public string qualification = string.Empty;

        /// <summary>
        /// 学历证书
        /// </summary>
        public string qualificationUrl = string.Empty;

        /// <summary>
        /// 上岗证
        /// </summary>
        public string workLicenseUrl = string.Empty;

        /// <summary>
        /// 学分证书
        /// </summary>
        public string creditUrl = string.Empty;

        /// <summary>
        /// 系统管理证书
        /// </summary>
        public string systemManagementCertificateUrl = string.Empty;

        /// <summary>
        /// 售后管理证书
        /// </summary>
        public string afterSalesUrl = string.Empty;

        /// <summary>
        /// 其他证书
        /// </summary>
        public string othersUrl = string.Empty;


    }
}
