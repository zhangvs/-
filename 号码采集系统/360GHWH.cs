using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace 号码采集系统
{
    public partial class _360GHWH : Form
    {
        public _360GHWH()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;
        }

        DateTime dt;
        string dateStr = DateTime.Now.ToString("yyyyMMdd");
        string search = "360";
        int num = 0;

        private void button1_Click(object sender, EventArgs e)
        {
            //C:\Users\Administrator\Documents\Visual Studio 2015\Projects\号码采集系统\号码采集系统\bin\x86\Debug\20170914
            string spath = Environment.CurrentDirectory + "\\" + dateStr;
            if (!Directory.Exists(spath))
            {
                DirectoryInfo directoryInfo = new DirectoryInfo(spath);
                directoryInfo.Create();
            }
            StringHelp.filePathOut = spath + "\\_360固话万号采集_" + textBox2.Text.Trim() + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xls";
            StringHelp.pathError = spath + "\\_360固话万号采集_" + textBox2.Text.Trim() + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".txt";
            StringHelp.CreateExcel();

            if (textBox1.Text.Trim() == "" || textBox2.Text.Trim() == "")
            {
                MessageBox.Show("号码不能为空！", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            dt = DateTime.Now;
            label3.Text = dt + "正在执行……";
            //执行主循环万号
            mainWhile(textBox1.Text.Trim(), textBox2.Text.Trim());
            DateTime end = DateTime.Now;
            StringHelp.Write(StringHelp.pathError, end.ToString() + "执行结束。" + "用时：" + DateHelp.DateDiff(end, dt) + "获取个数：" + num + "\r\n");
            label3.Text = end.ToString() + "执行完毕,已将检索结果导出到与导入文件所在的相同目录。";
            MessageBox.Show(end.ToString() + "执行结束。" + "用时：" + DateHelp.DateDiff(end, dt) + "。" +
                "\r\n获取个数：" + num + ", 已将检索结果导出到与导入文件所在的相同目录。\r\n");
        }

        void mainWhile(string qh, string q3)
        {
            
            string txt = q3 + "0000";
            long hd = Convert.ToInt64(txt);

            for (int i = 0; i < 10000; i++)
            {
                string keyword = qh + (hd + i).ToString();
                int bl = CollectRule.MainWhile(search, keyword);
                if (bl == 1)
                {
                    num++;
                }
                else if (bl == -1)
                {
                    MessageBox.Show(DateTime.Now.ToString() + "执行到：" + keyword + "时，当前IP开始被" + search + "屏蔽,未启动VPN！\r\n");
                }

            }
        }

        
    }
}
