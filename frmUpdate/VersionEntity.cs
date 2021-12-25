using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace frmUpdate
{
    public class VersionEntity
    {
        /// <summary>
        /// ID
        /// </summary>
        public int Id { get; set; }
        /// <summary>
        /// 文件名
        /// </summary>
        public string FileName { get; set; }
        /// <summary>
        /// 版本号 / MD5
        /// </summary>
        public string Version { get; set; }
        /// <summary>
        /// 导入路径 
        /// </summary>
        public string FilePath { get; set; }
        /// <summary>
        ///  服务器存储路径
        /// </summary>
        public string SavePath { get; set; }
        /// <summary>
        /// 更新文件类型 0：主程序(.exe)、1：文档文件、2：更新程序(其他dll文件)、5：更新程序、6：更新程序配置文件
        /// </summary>
        public int Type { get; set; }
        /// <summary>
        /// 是否删除
        /// </summary>
        public bool Deleted { get; set; }

        /// <summary>
        /// 所属区域
        /// </summary>
        public int Area { get; set; }

        /// <summary>
        /// 更新描述
        /// </summary>
        public string UpdateDescription { get; set; }
    }
}
