using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using HtmlAgilityPack;

namespace 号码采集系统
{
    public partial class ShuaLiang : Form
    {
        public ShuaLiang()
        {
            InitializeComponent();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DateTime hdStart = DateTime.Now;
            for (int x = 0; x < 1000; x++)
            {
                DateTime hdStart0 = DateTime.Now;
                for (int k = 0; k < 20; k++)
                {
                    if (textBox1.Text == "")
                    {
                        MessageBox.Show("网址为空！");
                    }

                    Process p = new Process();//引用using System.Diagnostics
                    p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;//将启动IEXPLORE的窗体设为隐藏
                    p.StartInfo.FileName = "IEXPLORE.EXE";//打开IEXPLORE
                    p.StartInfo.Arguments = textBox1.Text.Replace("/bbs/", "/o/bbs/");//输入要打开的网址
                                                                                      //替换旧版本连接/bbs/--》/o/bbs/
                    p.Start();
                    //DateTime mytime = p.StartTime;
                    //string mytime = p.StartTime.ToString();//定义一个变量记录刚才打开的网页的启动时间（为以后关闭它使用）

                    //建议：启动和关闭之间间隔几秒时间，让网页充分打开

                    //停顿时间
                    Thread.Sleep(2000);//3000毫秒

                    //关闭指定网页，则需要根据刚才记录的标示（mytime）来关闭网页（防止关闭其他已打开的网页）
                }
                Thread.Sleep(2000);//3000毫秒
                Process[] pp = Process.GetProcessesByName("iexplore");//iexplore
                for (int i = 0; i < pp.Length; i++)
                {
                    if (pp[i].HasExited == false)
                    {
                        pp[i].Kill();//关闭网页（进程）
                    }
                }
                DateTime hdEnd0 = DateTime.Now;
                //MessageBox.Show("执行完成25！" + "用时：" + DateHelp.DateDiff(hdEnd0, hdStart0));
            }

            DateTime hdEnd = DateTime.Now;
            MessageBox.Show("执行完成！" + "用时：" + DateHelp.DateDiff(hdEnd, hdStart));

        }

        //string[] repalys = { "加油奥利给！", "期盼降价", "看好这款车!" };

        private void button6_Click(object sender, EventArgs e)
        {
            //文件路径
            string filePath = System.AppDomain.CurrentDomain.BaseDirectory + "回复.txt";
            try
            {
                if (File.Exists(filePath))
                {
                    string rps = File.ReadAllText(filePath);
                    string ids = textBox8.Text;
                    if (rps == "")
                    {
                        MessageBox.Show("回复文件不存在！");
                    }


                    string[] urlarr = textBox1.Text.Split(new string[] { "\r\n" }, StringSplitOptions.None);//url
                    string[] rparr = rps.Split(new string[] { "\r\n" }, StringSplitOptions.None);//例句
                    string[] idarr = ids.Split(new string[] { "\r\n" }, StringSplitOptions.None);

                    //生成随机数句子
                    ArrayList myAL = new ArrayList();
                    Random rd = new Random();
                    for (int k = 0; k < 10; k++)
                    {
                        while (true)
                        {
                            int temp = rd.Next(1, rparr.Length);
                            if (!myAL.Contains(temp))
                            {
                                myAL.Add(temp);
                                break;
                            }
                        }
                    }

                    for (int u = 0; u < urlarr.Length; u++)
                    {
                        string url = urlarr[u];
                        string topicid = url.Split('/', '-')[6];//https://club.autohome.com.cn/bbs/thread/d846d5111aef1fdc/90932303-1.html
                        //HtmlAgilityPack 添加Nuget包1.9,兼容2.0fw
                        HtmlAgilityPack.HtmlDocument htmlDocument = new HtmlAgilityPack.HtmlDocument();

                        string htmlStr = StringHelp.GetHtml1(url);
                        htmlDocument.LoadHtml(htmlStr);
                        string daohang = "//*[@id='js-sticky-toolbar']/div/div/div[1]/div[1]/a";//浏览器-检查-复制xpath
                        var nods = htmlDocument.DocumentNode.SelectNodes(daohang);
                        string href = nods[0].Attributes["href"].Value;
                        string bbsid = href.Split('-')[2];


                        for (int i = 0; i < 10;)
                        {
                            for (int n = 0; n < idarr.Length; n++)
                            {
                                //取数据
                                int d = int.Parse(myAL[i].ToString());

                                string replyUrl = "https://club.autohome.com.cn/frontapi/reply/add";
                                Encoding encoding = Encoding.GetEncoding("utf-8");
                                IDictionary<string, string> parameters = new Dictionary<string, string>();
                                parameters.Add("bbs", "c");
                                parameters.Add("bbsid", bbsid);
                                parameters.Add("topicid", topicid);
                                parameters.Add("content", rparr[d]);
                                HttpWebResponse response = StringHelp.CreatePostHttpResponse(replyUrl, parameters, encoding, idarr[n]);
                                //打印返回值
                                Stream stream = response.GetResponseStream();   //获取响应的字符串流
                                StreamReader sr = new StreamReader(stream); //创建一个stream读取流
                                string jsonText = sr.ReadToEnd();   //从头读到尾，放到字符串html


                                JObject jo = (JObject)JsonConvert.DeserializeObject(jsonText);//或者JObject jo = JObject.Parse(jsonText);
                                string returncode = jo["returncode"].ToString();//输出 "深圳"
                                if (returncode == "0")
                                {
                                    textBox7.Text += "成功\r\n";
                                }
                                else
                                {
                                    //textBox7.Text += jo["message"].ToString() + "\r\n";
                                    MessageBox.Show(jo["message"].ToString());
                                }

                                Thread.Sleep(3000);//3s
                                i++;
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("回复文件不存在");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }






            MessageBox.Show("执行完成！");
        }
    }
}
