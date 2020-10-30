using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Web;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Xml;
using System.Threading;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;

namespace 号码采集系统
{
    class StringHelp
    {
        public static string pathError = "D:\\TelWH_Error" + DateTime.Now.ToString("yyyyMMdd") + ".txt";
        public static string filePathOut;

        public static string GetHtml0(string url, string keyword)
        {

            HttpWebRequest req;
            HttpWebResponse response;
            Stream stream;
            req = (HttpWebRequest)WebRequest.Create(url + keyword);


            StringBuilder sb = new StringBuilder();
            try
            {
                response = (HttpWebResponse)req.GetResponse();
                stream = response.GetResponseStream();
                int count = 0;
                byte[] buf = new byte[8192];
                string decodedString = null;
                do
                {
                    count = stream.Read(buf, 0, buf.Length);
                    if (count > 0)
                    {
                        decodedString = Encoding.GetEncoding("UTF-8").GetString(buf, 0, count);
                        sb.Append(decodedString);
                    }
                } while (count > 0);
            }
            catch
            {

                //Write(path, keyword+"\r\n");
                //try
                //{
                //    Write(path, keyword + "\r\n");
                //}
                //catch (Exception)
                //{

                //}
                //MessageBox.Show("网络连接错误", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            //stream.Close();
            //if (response != null)
            //{
            //    response.Close();
            //}
            //if (req != null)
            //{
            //    req.Abort();
            //}
            return sb.ToString();
        }

        ///// <summary> 
        ///// ASPX页面生成静态Html页面
        ///// </summary> 
        public static string GetHtml(string url, string keyword)
        {
            //string url = @"http://www." + domain + ".com/";
            //string encodedKeyword = HttpUtility.UrlEncode(keyword, Encoding.GetEncoding("UTF-8"));
            //string query = "s?wd=" + encodedKeyword;

            StreamReader sr = null;
            string str = "";
            try
            {
                //读取远程路径 
                WebRequest request = WebRequest.Create(url + keyword);
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                sr = new StreamReader(response.GetResponseStream(), Encoding.GetEncoding(response.CharacterSet));
                //Thread.Sleep(800);
                str = sr.ReadToEnd();
                sr.Close();
            }
            catch (Exception ex)
            {
                try
                {
                    Write(pathError, ex.Message + keyword + "\r\n");
                }
                catch (Exception)
                {

                }
            }

            return str;
        }


        public static string GetHtml1(string url)
        {
            string pageHtml = "";
            try
            {
                WebClient webClient = new WebClient();
                webClient.Credentials = CredentialCache.DefaultCredentials;//获取或设置用于向Internet资源的请求进行身份验证的网络凭据  
                Byte[] pageData = webClient.DownloadData(url);
                //pageHtml = Encoding.Default.GetString(pageData);  //如果获取网站页面采用的是GB2312，则使用这句         
                pageHtml = Encoding.UTF8.GetString(pageData); //如果获取网站页面采用的是UTF-8，则使用这句  

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return pageHtml;
        }

        public static string GetHtml2(string url, string keyword)
        {
            //string url = @"http://www."+ domain + ".com/";
            //string encodedKeyword = HttpUtility.UrlEncode(keyword, Encoding.GetEncoding("UTF-8"));
            //string query = "s?q=" + encodedKeyword;

            string pageHtml = "";
            try
            {
                //System.Threading.Thread.Sleep(3 * 1000);//vps切换ip，断开再连接的时候会报异常，休眠两秒

                WebClient webClient = new WebClient();
                webClient.Credentials = CredentialCache.DefaultCredentials;//获取或设置用于向Internet资源的请求进行身份验证的网络凭据  
                                                                           //getState();
                Object locker = new Object();
                lock (locker)
                {
                    Byte[] pageData = webClient.DownloadData(url + keyword);
                    //pageHtml = Encoding.Default.GetString(pageData);  //如果获取网站页面采用的是GB2312，则使用这句         
                    pageHtml = Encoding.UTF8.GetString(pageData); //如果获取网站页面采用的是UTF-8，则使用这句  
                }

                //using (StreamWriter sw = new StreamWriter("C:\\ouput.txt"))//将获取的内容写入文本  
                //{
                //    htm = sw.ToString();//测试StreamWriter流的输出状态，非必须  
                //    sw.Write(pageHtml);
                //}
            }
            catch (Exception ex)
            {
                try
                {
                    Write(pathError, DateTime.Now.ToString() + ex.Message + keyword + "\r\n");
                    System.Threading.Thread.Sleep(3 * 1000);
                    getState();//操作同一个vpn，其他线程引起的搜索引擎屏蔽切换IP，断网都会引起其他线程的断网，所以要循环vpn状态，等待vpn处于连接状态，才可以继续获取源代码，单线程不必加此句
                    Write(pathError, DateTime.Now.ToString() + "重新执行：" + keyword + "\r\n");
                    GetHtml2(url, keyword);
                }
                catch (Exception e)
                {

                }
            }

            return pageHtml;
        }

        private static readonly string DefaultUserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; SV1; .NET CLR 1.1.4322; .NET CLR 2.0.50727)";

        private static bool CheckValidationResult(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
        {
            return true; //总是接受   
        }

        public static HttpWebResponse CreatePostHttpResponse(string url, IDictionary<string, string> parameters, Encoding charset,string session)
        {
            HttpWebRequest request = null;
            //HTTPSQ请求
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(CheckValidationResult);
            request = WebRequest.Create(url) as HttpWebRequest;
            request.ProtocolVersion = HttpVersion.Version10;
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            request.Headers["Cookie"] = session;
            request.UserAgent = DefaultUserAgent;
            //如果需要POST数据   
            if (!(parameters == null || parameters.Count == 0))
            {
                StringBuilder buffer = new StringBuilder();
                int i = 0;
                foreach (string key in parameters.Keys)
                {
                    if (i > 0)
                    {
                        buffer.AppendFormat("&{0}={1}", key, parameters[key]);
                    }
                    else
                    {
                        buffer.AppendFormat("{0}={1}", key, parameters[key]);
                    }
                    i++;
                }
                byte[] data = charset.GetBytes(buffer.ToString());
                using (Stream stream = request.GetRequestStream())
                {
                    stream.Write(data, 0, data.Length);
                }
            }
            return request.GetResponse() as HttpWebResponse;
        }

        public static void RequestReferer(string url,int ahpvno,int v_no)
        {
            Random r = new Random();
            Thread.Sleep(1500);
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Timeout = 30 * 1000;//设置30s的超时
            //request.ContentType = "text/html; charset=utf-8";// "text/html;charset=gbk";// 
            request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9";
            //request.Headers["Accept-Encoding"] = "gzip, deflate, br";
            request.Headers["Accept-Language"] = "zh-CN,zh;q=0.9";
            request.Headers["Cache-Control"] = "no-cache";
            request.Headers["Cookie"] = "fvlid=16026567427291ew27x4OpA; sessionid=611274BC-A672-463E-A3EE-B633A4290E6C%7C%7C2020-10-14+14%3A25%3A43.535%7C%7C0; autoid=fb8ca6f98f5ece5f3d8e0dc33ad69588; ahpau=1; sessionuid=611274BC-A672-463E-A3EE-B633A4290E6C%7C%7C2020-10-14+14%3A25%3A43.535%7C%7C0; __ah_uuid_ng=u_203308454; sessionip=58.57.32.162; area=371302; __utma=1.1485844862.1602656756.1602663668.1603769203.3; __utmz=1.1603769203.3.3.utmcsr=i.autohome.com.cn|utmccn=(referral)|utmcmd=referral|utmcct=/; pcpopclub=04bdd95580cb4dce9e1249a4c29655740c1e3da6; clubUserShow=203308454|0|2|%e4%b8%b4%e6%b2%82%e8%bd%a6%e5%8f%8b6dhcxn|0|0|0|/g29/M08/86/87/120X120_0_q87_autohomecar__ChcCSF3eTKmAKfHhAABjOIth1jg298.jpg|2020-10-27 11:33:48|0; autouserid=203308454; sessionuserid=203308454; sessionlogin=fa001e60ad9c47e2a9f2cd4ee3e1e47c0c1e3da6; historybbsName4=c-5566%7C%E7%BA%A2%E6%97%97H9%2Cc-65%7C%E5%AE%9D%E9%A9%AC5%E7%B3%BB; sessionvid=950BBCE6-57E1-4984-A637-DBFB4C70C3BB; ahpvno=152; pvidchain=6835678,6835678; v_no=3; visit_info_ad=611274BC-A672-463E-A3EE-B633A4290E6C||950BBCE6-57E1-4984-A637-DBFB4C70C3BB||-1||-1||3; ref=localhost%7C0%7C0%7C0%7C2020-10-29+14%3A19%3A27.570%7C2020-10-27+18%3A08%3A56.990";
            //request.Headers["Host"] = "club.autohome.com.cn";
            //request.Headers["Connection"] = "keep-alive";
            request.Headers["Pragma"] = "no-cache";
            request.Headers["Sec-Fetch-Dest"] = "document";
            request.Headers["Sec-Fetch-Mode"] = "navigate";
            request.Headers["Sec-Fetch-Site"] = "none";
            request.Headers["Sec-Fetch-User"] = "?1";
            request.Headers["Upgrade-Insecure-Requests"] = "1";
            request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.106 Safari/537.36";

            HttpWebResponse response = null;
            Stream stream = null;
            StreamReader reader = null;
            try
            {
                response = (HttpWebResponse)request.GetResponse();
                stream = response.GetResponseStream();
                reader = new StreamReader(stream, Encoding.UTF8);
                string html = reader.ReadToEnd();//.Replace("\r\n", ""); 
            }
            catch (Exception excpt)
            {
            }
        }

        public static string DownloadHtml(string url, string keyword)
        {
            string html = string.Empty;
            try
            {
                HttpWebRequest request = HttpWebRequest.Create(url) as HttpWebRequest;//模拟请求
                request.Timeout = 30 * 1000;//设置30s的超时
                request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.106 Safari/537.36";
                //request.UserAgent = "User - Agent:Mozilla / 5.0(iPhone; CPU iPhone OS 7_1_2 like Mac OS X) App leWebKit/ 537.51.2(KHTML, like Gecko) Version / 7.0 Mobile / 11D257 Safari / 9537.53";

                request.ContentType = "text/html; charset=utf-8";// "text/html;charset=gbk";// 
                request.CookieContainer = new CookieContainer();
                request.Headers["sessionlogin"] = "b62d03f0617a4b9db1e23fef22428c8f0c1e3da6";
                //Cookie c1 = new Cookie("sessionlogin", "b62d03f0617a4b9db1e23fef22428c8f0c1e3da6");
                using (HttpWebResponse response = request.GetResponse() as HttpWebResponse)//发起请求
                {
                    if (response.StatusCode != HttpStatusCode.OK)
                    {
                        //logger.Warn(string.Format("抓取{0}地址返回失败,response.StatusCode为{1}", url, response.StatusCode));
                    }
                    else
                    {
                        try
                        {
                            Object locker = new Object();
                            lock (locker)
                            {
                                StreamReader sr = new StreamReader(response.GetResponseStream(), Encoding.GetEncoding("UTF-8"));
                                html = sr.ReadToEnd();//读取数据
                                sr.Close();
                            }
                        }
                        catch (Exception ex)
                        {
                            //logger.Error(string.Format($"DownloadHtml抓取{url}失败"), ex);
                            try
                            {
                                Write(pathError, DateTime.Now.ToString() + ex.Message + keyword + "\r\n");
                                System.Threading.Thread.Sleep(3 * 1000);
                                getState();//操作同一个vpn，其他线程引起的搜索引擎屏蔽切换IP，断网都会引起其他线程的断网，所以要循环vpn状态，等待vpn处于连接状态，才可以继续获取源代码，单线程不必加此句
                                Write(pathError, DateTime.Now.ToString() + "重新执行：" + keyword + "\r\n");
                                DownloadHtml(url, keyword);
                            }
                            catch (Exception e)
                            {

                            }
                        }
                    }
                }
            }
            catch (System.Net.WebException ex)
            {
                if (ex.Message.Equals("远程服务器返回错误: (306)。"))
                {
                    //logger.Error("远程服务器返回错误: (306)。", ex);
                    html = null;
                }
            }
            catch (Exception ex)
            {
                //logger.Error(string.Format("DownloadHtml抓取{0}出现异常", url), ex);
                html = null;
            }
            return html;
        }

        public static string getState()
        {
            Object locker = new Object();
            lock (locker)
            {
                string stateHtml = StringHelp.GetHtml1(@"http://127.0.0.1:8222/getstate/");

                while (stateHtml.IndexOf("<state>11</state>") < 0)
                {
                    Stconnect();
                    stateHtml = StringHelp.GetHtml1(@"http://127.0.0.1:8222/getstate/");
                }
                return stateHtml;
            }
        }

        public static void getChangeIp()
        {
            System.Threading.Thread.Sleep(3 * 1000);//等待屏蔽之前正在运行的线程结束
            Object locker = new Object();
            lock (locker)
            {
                Disconnect();
                Stconnect();
                //连接状态
                string stateHtml = getState();

                //连接成功之后，获取连接的ip和信息
                System.Xml.XmlDocument xml = new System.Xml.XmlDocument();
                xml.LoadXml(stateHtml);
                XmlNode ipNode = xml.SelectSingleNode("root/ip");
                string ip = ipNode.InnerText.Trim();
                XmlNode lineinfoNode = xml.SelectSingleNode("root/lineinfo");
                string lineinfo = lineinfoNode.InnerText.Trim();

                Write(pathError, DateTime.Now.ToString() + "当前ip：" + ip + "当前lineinfo：" + lineinfo + "\r\n");
            }
        }

        public static void Disconnect()
        {
            Object locker = new Object();
            lock (locker)
            {
                string disconHtml = "";
                //断开链接
                do
                {
                    disconHtml = GetHtml1(@"http://127.0.0.1:8222/disconnect/");
                    System.Threading.Thread.Sleep(1 * 1000);
                } while (disconHtml.IndexOf("<code>0</code>") < 0);

                Write(pathError, DateTime.Now.ToString() + "断开连接\r\n");
            }
        }

        public static void Stconnect()
        {
            Object locker = new Object();
            lock (locker)
            {
                //随机连接某一条静态线路
                string conHtml = "";
                int xm = 3;
                do
                {
                    conHtml = GetHtml1(@"http://127.0.0.1:8222/stconnect/");
                    System.Threading.Thread.Sleep(xm * 1000);
                    xm++;
                } while (conHtml.IndexOf("<code>0</code>") < 0 && conHtml.IndexOf("<code>106</code>") < 0);
                Write(pathError, DateTime.Now.ToString() + "随机连接某一条静态线路……\r\n");
            }
        }

        public static string ReplaceHtmlTag(string html, int length = 0)
        {
            string strText = System.Text.RegularExpressions.Regex.Replace(html, "<[^>]+>", "");
            strText = System.Text.RegularExpressions.Regex.Replace(strText, "&[^;]+;", "");

            if (length > 0 && strText.Length > length)
                return strText.Substring(0, length);

            return strText;
        }


        public static void Write(string path, string content)
        {
            Object locker = new Object();
            lock (locker)
            {
                try
                {
                    FileStream fs = new FileStream(path, FileMode.Append);
                    StreamWriter sw = new StreamWriter(fs, Encoding.Default);
                    //开始写入//使用与系统一致的编码方式
                    sw.Write(content);

                    //清空缓冲区
                    sw.Flush();
                    //关闭流
                    sw.Close();
                    fs.Close();
                }
                catch (Exception)
                {
                    System.Threading.Thread.Sleep(1 * 1000);
                    Write(path, content);
                }
            }
        }

        public static HSSFWorkbook newhssfworkbook;//创建新表格
        public static ISheet newsheet;//创建新工作簿
        public static IFont font;//创建字体样式  
        public static ICellStyle style;//创建单元格样式  

        public static void CreateExcel()
        {
            newhssfworkbook = new HSSFWorkbook();//创建新表格
            newsheet = newhssfworkbook.CreateSheet();//创建新工作簿
            font = newhssfworkbook.CreateFont();//创建字体样式  
            style = newhssfworkbook.CreateCellStyle();//创建单元格样式  

            string[] titles = { "Search", "Telphone", "Title", "Abstract" };
            newsheet.SetColumnWidth(0, (int)((10 + 0.72) * 256));
            newsheet.SetColumnWidth(1, (int)((15 + 0.72) * 256));
            newsheet.SetColumnWidth(2, (int)((60 + 0.72) * 256));
            newsheet.SetColumnWidth(3, (int)((150 + 0.72) * 256));

            IRow newrow0 = newsheet.CreateRow(0);//创建新行
            for (int i = 0; i < titles.Length; i++)
            {
                var cell = newrow0.CreateCell(i);
                cell.SetCellValue(titles[i]);
            }
        }

        //public static IFont font = newhssfworkbook.CreateFont();//创建字体样式  
        //public static ICellStyle style = newhssfworkbook.CreateCellStyle();//创建单元格样式  

        public static void WriteExcel(TelCollection tel)
        {
            #region 设置字体  
            font.Color = NPOI.HSSF.Util.HSSFColor.Blue.Index;//设置字体颜色  
            style.SetFont(font);//设置单元格样式中的字体样式  
            #endregion


            //默认数据行
            var newrow = newsheet.CreateRow(newsheet.LastRowNum + 1);//创建新行
            //搜索引擎
            var newcell0 = newrow.CreateCell(0);
            newcell0.SetCellValue(tel.Search);

            #region 号码
            var newcell1 = newrow.CreateCell(1);
            newcell1.SetCellValue(tel.Telphone);//手机号
            //设置号码超链接  
            HSSFHyperlink link = new HSSFHyperlink(HyperlinkType.Url);//建一个HSSFHyperlink实体，指明链接类型为URL（这里是枚举，可以根据需求自行更改） 
            if (tel.Search == "360")
            {
                link.Address = Reg360.url + tel.Telphone;//给HSSFHyperlink的地址赋值  
            }
            else
            {
                link.Address = RegBaidu.url + tel.Telphone;
            }
            newcell1.Hyperlink = link;//将链接方式赋值给单元格的Hyperlink即可将链接附加到单元格上 
            //超链接字体  
            newcell1.CellStyle = style;//为单元格设置显示样式   
            #endregion


            #region 标题
            var newcell2 = newrow.CreateCell(2);
            newcell2.SetCellValue(tel.Title);//标题
            //设置标题超链接  
            HSSFHyperlink titleLink = new HSSFHyperlink(HyperlinkType.Url);//建一个HSSFHyperlink实体，指明链接类型为URL（这里是枚举，可以根据需求自行更改） 
            titleLink.Address = tel.TitleLink;//给HSSFHyperlink的地址赋值  
            newcell2.Hyperlink = titleLink;//将链接方式赋值给单元格的Hyperlink即可将链接附加到单元格上 
            //设置字体  
            newcell2.CellStyle = style;//为单元格设置显示样式  
            #endregion

            //简介
            var newcell3 = newrow.CreateCell(3);
            newcell3.SetCellValue(tel.Abstract);

            Object locker = new Object();
            lock (locker)
            {
                try
                {
                    FileStream fs = File.OpenWrite(filePathOut);
                    newhssfworkbook.Write(fs);   //向打开的这个xls文件中写入表并保存。  
                    fs.Close();
                }
                catch (Exception ex)
                {
                    StringHelp.Write(StringHelp.pathError, ex.Message + tel.Telphone + "\r\n");
                    System.Threading.Thread.Sleep(1 * 1000);
                    WriteExcel(tel);
                    StringHelp.Write(StringHelp.pathError, tel.Telphone + "重新执行\r\n");
                }
            }
        }


    }
}
