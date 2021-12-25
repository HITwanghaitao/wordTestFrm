using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using wordTestFrm.ControlTool;
using System.IO;
using System.Configuration;
using System.Security.Cryptography;
using Newtonsoft.Json;
using RestSharp;
using System.Reflection;

namespace wordTestFrm
{
    static class Program
    {
        public static List<Image> loadImgItems = new List<Image>();
        public static List<Image> guideItems = new List<Image>();
        public static int width = 490;
        public static int height = 250;
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Thread th = new Thread(() =>
            {
                try
                {
                    string dirGif = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "gif");
                    if (Directory.Exists(dirGif))
                    {
                        Directory.Delete(dirGif, true);
                    }
                }
                catch (IOException ex)
                {

                }
                guideItems.Add(Properties.Resources.word格式化1);
                loadImgItems.Add(ucLoading.reSizeForImg(Properties.Resources.loading0, width, height, nameof(Properties.Resources.loading0)));

                guideItems.Add(ucLoading.reSizeForImg(Properties.Resources.生成word1, 150, 150, nameof(Properties.Resources.生成word1)));
                guideItems.Add(ucLoading.reSizeForImg(Properties.Resources.生成word_2, 150, 150, nameof(Properties.Resources.生成word_2)));
                guideItems.Add(ucLoading.reSizeForImg(Properties.Resources.生成word3, 150, 150, nameof(Properties.Resources.生成word3)));
                guideItems.Add(ucLoading.reSizeForImg(Properties.Resources.生成word4, 150, 150, nameof(Properties.Resources.生成word4)));
                guideItems.Add(ucLoading.reSizeForImg(Properties.Resources.生成word5, 150, 150, nameof(Properties.Resources.生成word5)));


                loadImgItems.Add(ucLoading.reSizeForImg(Properties.Resources.loadingCat, width, height, nameof(Properties.Resources.loadingCat)));
                loadImgItems.Add(ucLoading.reSizeForImg(Properties.Resources.loadingcat2, width, height, nameof(Properties.Resources.loadingcat2)));
                loadImgItems.Add(ucLoading.reSizeForImg(Properties.Resources.cloud, width, height, nameof(Properties.Resources.cloud)));
                loadImgItems.Add(ucLoading.reSizeForImg(Properties.Resources.loading1, width, height, nameof(Properties.Resources.loading1)));
                loadImgItems.Add(ucLoading.reSizeForImg(Properties.Resources.loading2, width, height, nameof(Properties.Resources.loading2)));
                loadImgItems.Add(ucLoading.reSizeForImg(Properties.Resources.loadingcar, width, height, nameof(Properties.Resources.loadingcar)));
                loadImgItems.Add(ucLoading.reSizeForImg(Properties.Resources.loadingOctopus, width, height, nameof(Properties.Resources.loadingOctopus)));
                loadImgItems.Add(ucLoading.reSizeForImg(Properties.Resources.man, width, height, nameof(Properties.Resources.man)));
                loadImgItems.Add(ucLoading.reSizeForImg(Properties.Resources.orangejuice, width, height, nameof(Properties.Resources.orangejuice)));
                loadImgItems.Add(ucLoading.reSizeForImg(Properties.Resources.panda, width, height, nameof(Properties.Resources.panda)));
                loadImgItems.Add(ucLoading.reSizeForImg(Properties.Resources.polygon, width, height, nameof(Properties.Resources.polygon)));
                loadImgItems.Add(ucLoading.reSizeForImg(Properties.Resources.polygon2, width, height, nameof(Properties.Resources.polygon2)));
                loadImgItems.Add(ucLoading.reSizeForImg(Properties.Resources.ship, width, height, nameof(Properties.Resources.ship)));

               

            });
            th.Start();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            string version = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            UpdateVersion("1");

            if (Environment.UserDomainName== "DESKTOP-HHETJO2")
               Application.Run(new FrmMain());
            else
            Application.Run(new FrmMain());
        }

        static void UpdateVersion(string area)
        {
            try
            {
                string url = "http://" + ConfigurationManager.AppSettings["Ip"] + "/api/System/new";
                //检查远程端是否有响应
                var sysStatus = GetServerInformation(url);
                if (string.IsNullOrEmpty(sysStatus) || sysStatus.Equals("Too many connections"))
                {
                    //MessageBox.Show("未连接到远程服务端无法更新，请联系管理员", "确认", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //Environment.Exit(0);
                    return;
                }
                if (sysStatus.Equals("0")) return;
                //获取版本数据
                url = "http://" + ConfigurationManager.AppSettings["Ip"] + "/api/System/version/" + area;
                var res = GetServerInformation(url);
                //转换转义符
                res = System.Text.RegularExpressions.Regex.Unescape(res);
                if (res.Equals("not found"))
                    return;
                var dt = JsonConvert.DeserializeObject<List<VersionEntity>>(res);
                var client = new System.Net.WebClient();

                //最终需要更新的数据ID
                if (dt != null && dt.Count > 0)
                {
                    foreach (var item in dt)
                    {
                        string downloadPath = "http://" + ConfigurationManager.AppSettings["Ip"] + item.SavePath + item.FileName;

                        if (item.Type == 5)
                        {
                            if (!File.Exists(Application.StartupPath + @"\Update.exe"))
                            {
                                string filePath = "";
                                filePath = Path.Combine(Application.StartupPath, item.FileName);
                                client.DownloadFile(downloadPath, filePath);
                                continue;
                            }
                            System.Diagnostics.FileVersionInfo fileVerInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.StartupPath + @"\Update.exe");
                            if (fileVerInfo.FileVersion != null && !fileVerInfo.FileVersion.Equals(item.Version))
                            {
                                string filePath = "";
                                filePath = Path.Combine(Application.StartupPath, item.FileName);
                                client.DownloadFile(downloadPath, filePath);
                            }
                        }
                        else if (item.Type == 6)
                        {
                            string path = "";
                            path = Path.Combine(Application.StartupPath, item.FileName);
                            string md5local = GetMD5HashFromFile(path);
                            if (!item.Version.Equals(md5local))
                            {
                                client.DownloadFile(downloadPath, path);
                            }
                        }
                    }
                }

                if (GetIsUpdate() > 0)
                {
                    var result = MessageBox.Show("发现有版本，是否更新？", "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.No)
                    {
                        return;
                    }
                    string startFile= Path.Combine(Application.StartupPath, ConfigurationManager.AppSettings["name"] + ".exe");
                    if (File.Exists(startFile))
                    {
                        System.Diagnostics.Process proce = new System.Diagnostics.Process();
                        proce.StartInfo.FileName = startFile;
                        proce.StartInfo.Arguments = "11";
                        proce.Start();
                        return;
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("未连接到远程服务端");
                //Application.Exit();
            }

        }

        public static int GetIsUpdate()
        {
            int count = 0;
            string httpurl = "http://" + ConfigurationManager.AppSettings["Ip"] + "/api/System/version/1";
            string objStr = Program.GetServerInformation(httpurl);
            //objStr = System.Text.RegularExpressions.Regex.Unescape(objStr);
            if (objStr.Equals("not found"))
            {
                return 0;
            }
            var listVersionEntity = new System.Web.Script.Serialization.JavaScriptSerializer().Deserialize<List<VersionEntity>>(objStr);
            string url = "http://" + ConfigurationManager.AppSettings["Ip"];
            if (listVersionEntity != null)
            {
                //文件下载至文件夹路径
                string downloadDirectoty = Path.Combine(Application.StartupPath, "download");

                foreach (var item in listVersionEntity)
                {
                    string downloadPath = url + item.SavePath + item.FileName;
                    if (item.Type == 0)
                    {
                        string localPath = Path.Combine(Application.StartupPath , item.FileName);
                        if (File.Exists(localPath))
                        {
                            //更新主程序 exe 获取文件版本号是否一致
                            System.Diagnostics.FileVersionInfo fileVerInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.StartupPath + @"\" + item.FileName);
                            if (!fileVerInfo.FileVersion.Equals(item.Version))
                            {
                                count++;
                            }
                        }
                        else {
                            count++;
                        }
                    }
                    else if (item.Type == 1)
                    {
                        //更新文本文件，查看MD5是否一致
                        string path = "";
                        if (string.IsNullOrEmpty(item.FilePath))
                            path = Path.Combine(Application.StartupPath, item.FileName);
                        else
                            path = Path.Combine(new string[] { Application.StartupPath, item.FilePath, item.FileName });
                        string md5local = Program.GetMD5HashFromFile(path);
                        if (!item.Version.Equals(md5local))
                        {
                            count++;
                        }
                    }
                    else if (item.Type == 2)
                    {
                        //更新其他dll文件，获取版本号是否一致
                        string path = "";
                        if (string.IsNullOrEmpty(item.FilePath))
                            path = Path.Combine(Application.StartupPath, item.FileName);
                        else
                            path = Path.Combine(new string[] { Application.StartupPath, item.FilePath, item.FileName });
                        if (!File.Exists(path))
                        {
                            count++;
                            continue;
                        }
                        System.Diagnostics.FileVersionInfo fvi = System.Diagnostics.FileVersionInfo.GetVersionInfo(path);
                        if (!fvi.FileVersion.Equals(item.Version))
                        {
                            count++;
                        }
                    }
                }
            }
            return count;
        }
        /// <summary>
        /// 启动程序调用此方法，用来更新 更新程序
        /// </summary>
        /// <param name="area"></param>

        public static string GetMD5HashFromFile(string fileName)
        {
            try
            {
                if (!File.Exists(fileName))
                    return "-1";
                string val = "";
                FileStream file = new FileStream(fileName, FileMode.Open);
                MD5 md5 = MD5.Create();
                byte[] b = md5.ComputeHash(file);
                file.Close();
                for (int i = 0; i < b.Length; i++)
                {
                    val += b[i].ToString("x");
                }
                return val;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public static string GetServerInformation(string url)
        {
            var client = new RestClient(url);
            client.Timeout = 5000;
            client.MaxRedirects = int.MaxValue;
            var request = new RestRequest("", Method.GET);
            IRestResponse res = client.Execute(request);
            return res.Content;
        }

    }

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
