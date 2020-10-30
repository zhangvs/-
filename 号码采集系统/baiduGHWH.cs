using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web;
using System.Windows.Forms;

namespace 号码采集系统
{
    public partial class baiduGHWH : Form
    {
        public baiduGHWH()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;
        }
        RegisteredWaitHandle rhw;
        DateTime dt;
        string dateStr = DateTime.Now.ToString("yyyyMMdd");
        int num = 0;
        string search = "baidu";

        private void button1_Click(object sender, EventArgs e)
        {
            //C:\Users\Administrator\Documents\Visual Studio 2015\Projects\号码采集系统\号码采集系统\bin\x86\Debug\20170914
            string spath = Environment.CurrentDirectory + "\\" + dateStr;
            if (!Directory.Exists(spath))
            {
                DirectoryInfo directoryInfo = new DirectoryInfo(spath);
                directoryInfo.Create();
            }
            StringHelp.filePathOut = spath + "\\baidu固话万号采集_" + textBox2.Text.Trim() + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xls";
            StringHelp.pathError = spath + "\\baidu固话万号采集_" + textBox2.Text.Trim() + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".txt";
            //创建excel文件
            StringHelp.CreateExcel();

            if (textBox1.Text.Trim() == "" || textBox2.Text.Trim() == "")
            {
                MessageBox.Show("号码不能为空！", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            dt = DateTime.Now;
            label3.Text = dt + "正在执行……";
            mainWhile(textBox1.Text.Trim(), textBox2.Text.Trim());
        }


        void mainWhile(string qh, string q3)
        {
            string txt = q3 + "0000";
            long hd = Convert.ToInt64(txt);

            for (int i = 0; i < 10000; i++)
            {
                ThreadPool.QueueUserWorkItem(new WaitCallback(SearchResult), qh + (hd + i).ToString());
                ThreadPool.QueueUserWorkItem(new WaitCallback(SearchResult), qh + "-" + (hd + i).ToString());

                #region MyRegion
                //不带 - 的，05398930001
                //HtmlResult hr = new HtmlResult();
                //hr.index = i;
                //hr.keyword = qh + (hd + i).ToString();
                //hr.htmlStr = StringHelp.GetHtml2(url, hr.keyword);
                //hr.input = r.Match(hr.htmlStr).Value;
                //if (hr.input!="")
                //{
                //    ThreadPool.QueueUserWorkItem(new WaitCallback(SearchResult), hr);
                //}

                ////带-的，0539-8930001
                //HtmlResult hr2 = new HtmlResult();
                //hr2.keyword = qh + "-" + (hd + i).ToString();
                //hr2.htmlStr = StringHelp.GetHtml2(url, hr2.keyword);
                //hr2.input = r.Match(hr2.htmlStr).Value;
                //if (hr2.input!="")
                //{
                //    ThreadPool.QueueUserWorkItem(new WaitCallback(SearchResult), hr2);
                //} 
                #endregion

            }

            rhw = ThreadPool.RegisterWaitForSingleObject(new System.Threading.AutoResetEvent(false), CheckThreadPool, null, 3000, false);
        }

        private void CheckThreadPool(object state, bool timeout)
        {
            Object locker = new Object();
            lock (locker)
            {
                int maxWorkerThreads, workerThreads;
                int portThreads;
                ThreadPool.GetMaxThreads(out maxWorkerThreads, out portThreads);
                ThreadPool.GetAvailableThreads(out workerThreads, out portThreads);
                if (maxWorkerThreads - workerThreads == 0)
                {
                    // 走到这里，所有线程都结束了
                    rhw.Unregister(null);
                    DateTime end = DateTime.Now;
                    StringHelp.Write(StringHelp.pathError, end.ToString() + "执行结束。" + "用时：" + DateHelp.DateDiff(end, dt) + "获取个数：" + num + "\r\n");
                    label3.Text = end.ToString() + "执行完毕,已将检索结果导出到与导入文件所在的相同目录。";
                    MessageBox.Show(end.ToString() + "执行结束。" + "用时：" + DateHelp.DateDiff(end, dt) + "。"+
                        "\r\n获取个数：" + num +", 已将检索结果导出到与导入文件所在的相同目录。\r\n");
                }
            }

        }

        void SearchResult(object qh_hd)
        {
            string keyword = qh_hd.ToString();
            int bl = CollectRule.MainWhile(search, keyword);
            if (bl == 1)
            {
                num++;
            }
            else if (bl == -1)
            {
                MessageBox.Show(DateTime.Now.ToString() + "执行到：" + keyword + "时，当前IP开始被" + search + "屏蔽,未启动VPN！\r\n");
            }

            #region MyRegion
            //Regex r = new Regex("<div id=\"content_left\"[\\s\\S]*<div style=\"clear:both;height:0;\"></div>", RegexOptions.IgnoreCase);
            //string htmlStr = StringHelp.GetHtml2(url, keyword);
            //string input = r.Match(htmlStr).Value;

            ////HtmlResult hr = (HtmlResult)ht;
            //if (htmlStr.IndexOf("您的电脑或所在的局域网络有异常的访问") > 0)
            //{
            //    try
            //    {
            //        StringHelp.Write(StringHelp.pathError, DateTime.Now.ToString() + "——————当前IP开始被baidu屏蔽————————————" + keyword + "\r\n");
            //        do
            //        {
            //            StringHelp.getChangeIp();
            //            htmlStr = StringHelp.GetHtml2(url, keyword);
            //        } while (hr.htmlStr.IndexOf("您的电脑或所在的局域网络有异常的访问") > 0);

            //        if (hr.input != "")
            //        {
            //            MainProc(hr.keyword, hr.input);
            //        }
            //    }
            //    catch (Exception ex)
            //    {
            //        StringHelp.Write(StringHelp.pathError, DateTime.Now.ToString() + ex + hr.keyword + "\r\n");
            //        MessageBox.Show("baidu开始屏蔽！当前执行到" + hr.keyword + "用时:" + DateHelp.DateDiff(DateTime.Now, dt) + "，未启动IP精灵", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    }
            //}
            //else if (hr.input != "")
            //{
            //    if (MainProc(hr.keyword, hr.input))
            //    {
            //        num++;
            //    }
            //} 
            #endregion
        }

        
    }
}
