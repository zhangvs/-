using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
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

namespace 号码采集系统
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string x = textBox1.Text;
            if (x == "")
            {
                MessageBox.Show("亲，网址不能为空！");
                return;
            }
            xx();
            timer1.Enabled = true;
            button3.Enabled = false;
            button4.Enabled = true;

        }

        public void xx()
        {
            string x = textBox1.Text;
            Regex reg = new Regex(@",");//网址之间用逗号隔开
            string[] r = reg.Split(x);

            Process ps = new Process();
            ps.StartInfo.FileName = "iexplore.exe";//调用IE
            for (int i = 0; i < r.Length; i++)
            {
                System.Diagnostics.Process.Start("" + r[i] + "");

            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            xx();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            button4.Enabled = false;
            button3.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox2.Text == "")
            {
                MessageBox.Show("设定时间不能为空");
                return;
            }
            int x = Convert.ToInt32(textBox2.Text) * 1000;
            timer1.Interval = x;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            button4.Enabled = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox3.Text == "")
            {
                MessageBox.Show("设定时间不能为空");
                return;
            }
            timer2.Interval = Convert.ToInt32(textBox3.Text) * 60000;

        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            Process[] myProcess;
            myProcess = Process.GetProcessesByName("iexplore");//360chrome.exe
            foreach (Process instance in myProcess)
            {
                instance.WaitForExit(3);
                instance.CloseMainWindow();
            }
        }

        //private void timer2_Tick(object sender, EventArgs e)
        //{//指定时间间隔发生事件
        //    Process[] CurrentPro = Process.GetProcesses();
        //    for (int i = 0; i < CurrentPro.Length; i++)
        //    {
        //        if (CurrentPro[i].MainWindowTitle.Contains("三国杀"))
        //        {
        //            //自动关闭
        //            DateTime now = DateTime.Now;
        //            DateTime end = DateTime.Parse("2013-12-2 16:15:50");
        //            TimeSpan ts = end.Subtract(now);
        //            if (ts.Seconds < 10)
        //            {
        //                this.Close();
        //                break;
        //            }
        //            MessageBox.Show("^_^就不让你玩^_^即将关闭浏览器！");
        //            time();
        //            try
        //            {
        //                CurrentPro[i].Kill();
        //            }
        //            catch
        //            {
        //                MessageBox.Show("好像出错了~");
        //            }
        //        }
        //    }
        //}


        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("iexplore.exe", "http://www.gliii.com");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //timer2.Enabled = true;
            for (int k = 0; k < 10; k++)
            {
                Process p = new Process();//引用using System.Diagnostics
                p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;//将启动IEXPLORE的窗体设为隐藏
                p.StartInfo.FileName = "IEXPLORE.EXE";//打开IEXPLORE
                p.StartInfo.Arguments = textBox1.Text;//输入要打开的网址
                if (textBox1.Text=="")
                {
                    MessageBox.Show("网址为空！");
                }
                p.Start();
                //DateTime mytime = p.StartTime;
                //string mytime = p.StartTime.ToString();//定义一个变量记录刚才打开的网页的启动时间（为以后关闭它使用）

                //建议：启动和关闭之间间隔几秒时间，让网页充分打开

                //停顿时间
                Thread.Sleep(3000);//3000毫秒

                //关闭指定网页，则需要根据刚才记录的标示（mytime）来关闭网页（防止关闭其他已打开的网页）
                Process[] pp = Process.GetProcessesByName("iexplore");//iexplore
                for (int i = 0; i < pp.Length; i++)
                {
                    if (pp[i].MainWindowTitle!="")
                    {
                        DateTime now = DateTime.Now;
                        TimeSpan ts = now.Subtract(pp[i].StartTime);//如果打开时间超过10s钟则关闭
                        if (ts.Seconds > 10)
                        {
                            pp[i].Kill();//关闭网页（进程）
                        }
                    }

                    //if (pp[i].StartTime.ToString() == mytime)//判断已打开的网页启动时间
                    //{
                    //    pp[i].Kill();//关闭网页（进程）
                    //}
                }
            }

        }

        string[] repalys = { "加油奥利给！", "期盼降价", "看好这款车!" };

        private void button6_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < repalys.Length; i++)
            {
                string url = "https://club.autohome.com.cn/frontapi/reply/add";
                Encoding encoding = Encoding.GetEncoding("utf-8");
                IDictionary<string, string> parameters = new Dictionary<string, string>();
                parameters.Add("bbs", textBox4.Text);
                parameters.Add("bbsid", textBox5.Text);
                parameters.Add("topicid", textBox6.Text);
                parameters.Add("content", repalys[i]);
                HttpWebResponse response = StringHelp.CreatePostHttpResponse(url, parameters, encoding, textBox8.Text);
                //打印返回值
                Stream stream = response.GetResponseStream();   //获取响应的字符串流
                StreamReader sr = new StreamReader(stream); //创建一个stream读取流
                string jsonText = sr.ReadToEnd();   //从头读到尾，放到字符串html


                JObject jo = (JObject)JsonConvert.DeserializeObject(jsonText);//或者JObject jo = JObject.Parse(jsonText);
                string returncode = jo["returncode"].ToString();//输出 "深圳"
                if (returncode=="0")
                {
                    textBox7.Text += "成功\r\n";
                }
                else
                {
                    textBox7.Text += jo["message"].ToString()+ "\r\n";
                }


                Thread.Sleep(3000);//3000毫秒
            }
        }
    }
}
