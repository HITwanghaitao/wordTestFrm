using frmUpdate;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace Update
{
    public partial class frmUpdate : Form
    {
        private const int Alpha = 111;
        public int Percentage = 0;
        Pen srcPen;
        SolidBrush srcBrush;

        private const uint WS_EX_LAYERED = 0x80000;
        private const int WS_EX_TRANSPARENT = 0x20;
        private const int GWL_EXSTYLE = (-20);
        private string Var_genre = "";//记录当前操作的类型

        #region 在窗口结构中为指定的窗口设置信息
        /// <summary>
        /// 在窗口结构中为指定的窗口设置信息
        /// </summary>
        /// <param name="hwnd">欲为其取得信息的窗口的句柄</param>
        /// <param name="nIndex">欲取回的信息</param>
        /// <param name="dwNewLong">由nIndex指定的窗口信息的新值</param>
        /// <returns></returns>
        [DllImport("user32", EntryPoint = "SetWindowLong")]
        private static extern uint SetWindowLong(IntPtr hwnd, int nIndex, uint dwNewLong);
        #endregion

        #region 从指定窗口的结构中取得信息
        /// <summary>
        /// 从指定窗口的结构中取得信息
        /// </summary>
        /// <param name="hwnd">欲为其获取信息的窗口的句柄</param>
        /// <param name="nIndex">欲取回的信息</param>
        /// <returns></returns>
        [DllImport("user32", EntryPoint = "GetWindowLong")]
        private static extern uint GetWindowLong(IntPtr hwnd, int nIndex);
        #endregion

        #region 使窗口有鼠标穿透功能
        /// <summary>
        /// 使窗口有鼠标穿透功能
        /// </summary>
        private void CanPenetrate()
        {
            uint intExTemp = GetWindowLong(this.Handle, GWL_EXSTYLE);
            uint oldGWLEx = SetWindowLong(this.Handle, GWL_EXSTYLE, WS_EX_TRANSPARENT | WS_EX_LAYERED);
        }
        #endregion


        public frmUpdate(string[] info)
        {
            InitializeComponent();
            this.SetStyle(ControlStyles.OptimizedDoubleBuffer |
                    ControlStyles.ResizeRedraw |
                    ControlStyles.AllPaintingInWmPaint, true);
            this.TransparencyKey = this.BackColor;

            if (info.Length > 0)
            {
                str = System.Web.HttpUtility.UrlDecode(info[0]);
            }
            else
            {
                MessageBox.Show("参数异常", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(0);
            }
        }
        WebClient client = null;
        /// <summary>
        /// 启动程序名称
        /// </summary>
        string ExeName = "";
        /// <summary>
        /// 启动程序需要的参数
        /// </summary>
        string str = "";


        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                //开启 WS_EX_TRANSPARENT, 使控件支持透明
                cp.ExStyle |= 0x00000020;
                return cp;
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            float width;
            float height;

            Color srcColor = Color.FromArgb(Alpha, this.BackColor);

            srcPen = new Pen(srcColor, 0);

            srcBrush = new SolidBrush(srcColor);

            base.OnPaint(e);

            width = this.Size.Width;

            height = this.Size.Height;

            e.Graphics.DrawRectangle(srcPen, 0, 0, width, height);
            e.Graphics.FillRectangle(srcBrush, 0, 0, width, height);

            e.Graphics.DrawString(this.Text, this.Font, new SolidBrush(Color.Black), new Rectangle(0, 0, this.Width, this.Height), new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center });
            this.srcPen.Dispose();
            this.srcBrush.Dispose();
            base.OnPaint(e);
        }


        private void frmUpdate_Load(object sender, EventArgs e)
        {
            try
            {
                ExeName = ConfigurationManager.AppSettings["name"];
                //杀掉客户端程序，确保其关闭未被占用
                Process[] ps = Process.GetProcesses();
                foreach (Process p in ps)
                {
                    if (p.ProcessName.Equals(ExeName, StringComparison.InvariantCultureIgnoreCase))
                    {
                        p.Kill();
                        break;
                    }
                }
                string url = "http://" + ConfigurationManager.AppSettings["Ip"] + "/api/System/new";
                string result = Program.GetServerInformation(url);
                if (string.IsNullOrEmpty(result)||result.Equals("0") || result.Equals("-1"))
                {
                    ProcessPro();
                    Environment.Exit(0);
                }
                client = new WebClient();
                DownloadUpdate d = DownloadUpdates;
                d.BeginInvoke(CompleteDownloadUpdate, d);
            }
            catch (Exception ex)
            {
                ProcessPro();
                Environment.Exit(0);
            }
        }

        private delegate void DownloadUpdate();
        private void CompleteDownloadUpdate(IAsyncResult iar)
        {
            DownloadUpdate d = (DownloadUpdate)iar.AsyncState;
            d.EndInvoke(iar);
        }

        /// <summary>
        /// 下载更新
        /// </summary>
        private void DownloadUpdates()
        {
            try
            {
                string httpurl = "http://" + ConfigurationManager.AppSettings["Ip"] + "/api/System/version/"+ ConfigurationManager.AppSettings["area"];
                string objStr = Program.GetServerInformation(httpurl);
                //objStr = System.Text.RegularExpressions.Regex.Unescape(objStr);
                if (objStr.Equals("not found"))
                {
                    ProcessPro();
                    Environment.Exit(0);
                }
                var listVersionEntity = new JavaScriptSerializer().Deserialize<List<VersionEntity>>(objStr);
                string url = "http://" + ConfigurationManager.AppSettings["Ip"];
                if (listVersionEntity != null)
                {
                    //文件下载至文件夹路径
                    string downloadDirectoty = Path.Combine(Application.StartupPath, "download");
                    if (Directory.Exists(downloadDirectoty))
                    {
                        Directory.Delete(downloadDirectoty, true);
                    }
                    Directory.CreateDirectory(downloadDirectoty);

                    foreach (var item in listVersionEntity)
                    {
                        string downloadPath = url + item.SavePath + item.FileName;
                        if (item.Type == 0)
                        {
                            //更新主程序 exe 获取文件版本号是否一致
                            FileVersionInfo fileVerInfo = FileVersionInfo.GetVersionInfo(Application.StartupPath + @"\" + ConfigurationManager.AppSettings["name"] + ".exe");
                            if (!fileVerInfo.FileVersion.Equals(item.Version))
                            {
                                client.DownloadFile(downloadPath, Path.Combine(downloadDirectoty, item.FileName));
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
                                client.DownloadFile(downloadPath, Path.Combine(downloadDirectoty, item.FileName));
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
                                client.DownloadFile(downloadPath, Path.Combine(downloadDirectoty, item.FileName));
                                continue;
                            }
                            FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(path);
                            if (!fvi.FileVersion.Equals(item.Version))
                            {
                                client.DownloadFile(downloadPath, Path.Combine(downloadDirectoty, item.FileName));
                            }
                        }
                        //client.DownloadFile(downloadPath, filePath);
                    }
                    foreach (var file in Directory.GetFiles(downloadDirectoty))
                    {
                        //全部下载完毕，一起更新程序
                        var item = listVersionEntity.FirstOrDefault(s => s.FileName.Equals(Path.GetFileName(file)));
                        if (!Directory.Exists(Path.Combine(Application.StartupPath, string.IsNullOrEmpty(item.FilePath) ? "" : item.FilePath)))
                        {
                            Directory.CreateDirectory(Path.Combine(Application.StartupPath, string.IsNullOrEmpty(item.FilePath) ? "" : item.FilePath));
                        }
                        string filePath = "";
                        if (string.IsNullOrEmpty(item.FilePath))
                            filePath = Path.Combine(Application.StartupPath, item.FileName);
                        else
                            filePath = Path.Combine(new string[] { Application.StartupPath, item.FilePath, item.FileName });
                        File.Copy(file, filePath, true);
                    }
                }
                ProcessPro();
                Environment.Exit(0);
            }
            catch (Exception ex)
            {
                MessageBox.Show("更新异常请联系管理员", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProcessPro();
                Environment.Exit(0);

            }
        }

        /// <summary>
        /// 启动程序
        /// </summary>
        private void ProcessPro()
        {
            Process proce = new Process();
            proce.StartInfo.FileName = Path.Combine(Application.StartupPath, ConfigurationManager.AppSettings["name"] + ".exe");
            proce.StartInfo.Arguments = str;
            proce.Start();
        }

        /// <summary>
        /// 获取版本号
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        private string GetServerInformation(string url)
        {
            var client = new RestClient(url);
            client.Timeout = 5000;
            client.MaxRedirects = int.MaxValue;
            var request = new RestRequest("", Method.GET);
            IRestResponse res = client.Execute(request);
            return res.Content;
        }

        /// <summary>
        /// 获取文件的MD5值
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private string GetMD5HashFromFile(string fileName)
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

    }
    
}
