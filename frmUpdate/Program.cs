using frmUpdate;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Windows.Forms;

namespace Update
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new frmUpdate(args));
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
}
