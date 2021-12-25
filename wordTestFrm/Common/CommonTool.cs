using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;

namespace wordTestFrm.Common
{
    public class CommonTool
    {
        #region IO操作
        /// <summary>
        /// 生成存储目录
        /// </summary>
        /// <param name="SaveDir">存储目录</param>
        /// <param name="fileName">文件名</param>
        /// <param name="flagStr">生成的特有后缀</param>
        /// <returns>反馈唯一的存储文件路径</returns>
        public static string retSaveFilePath(string SaveDir, string fileName, string flagStr = "AsposeWord")
        {
            string filePath = string.Empty;
            if (!Directory.Exists(SaveDir)) Directory.CreateDirectory(SaveDir);
            string tmpPath = Path.GetFileNameWithoutExtension(fileName) + "_" + flagStr + "_";
            List<string> items = Directory.GetFiles(SaveDir).ToList().FindAll(item => item.Contains(tmpPath));
            string file = items.OrderByDescending(item => {
                string tmpVal = item.Substring(item.LastIndexOf('_') + 1);
                try
                {
                    int result = int.Parse(tmpVal.Substring(0, tmpVal.LastIndexOf('.')));
                    return result;
                }
                catch (FormatException ex)
                {
                    return 0;
                }
                catch (Exception ex)
                {
                    return 0;
                }

            }).Max();
            int maxIndex = 1;
            if (!string.IsNullOrEmpty(file))
            {
                string val = file.Substring(file.LastIndexOf('_') + 1);
                maxIndex = int.Parse(val.Substring(0, val.LastIndexOf('.'))) + 1;
            }
            string fileType = fileName.Substring(fileName.LastIndexOf('.'));
            string saveFileName = Path.GetFileNameWithoutExtension(fileName) + "_" + flagStr + "_" + maxIndex.ToString("00") + fileType;
            filePath = Path.Combine(SaveDir, saveFileName);


            /*
            int index = 1;
            while(true)
            {
                string saveFileName = Path.GetFileNameWithoutExtension(fileName) + "_" + flagStr + "_" + index.ToString("00") + fileType;
                filePath = Path.Combine(SaveDir, saveFileName);
                if(!File.Exists(filePath))
                {
                    break;
                }
                index++;
            }*/
            return filePath;
        }

        /// <summary>
        /// 复制并反馈备份文件
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static string retSavePath(string filePath)
        {
            string fileName = Path.GetFileName(filePath);

            string saveDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Work", "CopyFile", Path.GetFileNameWithoutExtension(fileName));
            if (!Directory.Exists(saveDir))
            {
                Directory.CreateDirectory(saveDir);
            }

            string savePath = Path.Combine(saveDir, fileName);  //CommonTool.retSaveFilePath(saveDir, fileName);
            
            File.Copy(filePath, savePath, true);
            return savePath;
        }

        #endregion

        #region http请求
        public static string httpApi(string url, string jsonStr = "",
            string type = "POST")
        {
            if (string.IsNullOrEmpty(url))
            {
                //LogHelper.WriteLog("API链接地址为空，配置INI文件缺失");
                return null;
            }

            string result = "";//返回结果
            try
            {
                Encoding encoding = Encoding.UTF8;
                HttpWebResponse response;
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);//webrequest请求api地址  
                //if (Globas.token != null)//是否加上Token令牌
                //{
                //    //var headers = request.Headers;
                //    //headers["Authorization"] = "Bearer " + Globas.token.token;
                //    //request.Headers = headers;
                //    request.Headers.Add("Authorization", "Bearer " + Globas.token.token);
                //    request.Credentials = CredentialCache.DefaultCredentials;
                //}
                request.Accept = "text/html,application/xhtml+xml,*/*";
                request.ContentType = "application/json";
                request.Method = type.ToUpper().ToString();//get或者post
                if (!string.IsNullOrEmpty(jsonStr))//Get请求无需拼接此参数
                {
                    byte[] buffer = encoding.GetBytes(jsonStr);
                    request.ContentLength = buffer.Length;
                    request.GetRequestStream().Write(buffer, 0, buffer.Length);
                }

                try
                {
                    response = (HttpWebResponse)request.GetResponse();
                }
                catch (WebException ex)
                {
                    return "-1";
                    //response = (HttpWebResponse)ex.Response;
                }
                using (StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                {
                    result = reader.ReadToEnd();
                    reader.Close();
                }
                if (response.StatusCode != HttpStatusCode.OK)//返回响应码非成功格式化数据后返回
                {
                    return "-1";
                    //result = "Exception:" + JsonConvert.DeserializeObject<string>(result);
                }
                return result;
            }
            catch (WebException ex)
            {
                return "-1";
                //return "Exception:" + ex.Message;
            }
        }

        #endregion
    }
}
