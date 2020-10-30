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
    public partial class _360SJWH : Form
    {
        public _360SJWH()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;
        }

        RegisteredWaitHandle rhw;
        DateTime dt;
        string dateStr = DateTime.Now.ToString("yyyyMMdd");
        string z4Str = "";
        string search = "360";
        int num = 0;

        private void _360SJWH_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //C:\\Users\\Administrator\\Documents\\Visual Studio 2015\\Projects\\号码采集系统\\号码采集系统\\bin\\Debug
            string spath = Environment.CurrentDirectory + "\\" + dateStr;
            if (!Directory.Exists(spath))
            {
                DirectoryInfo directoryInfo = new DirectoryInfo(spath);
                directoryInfo.Create();
            }

            StringHelp.filePathOut= spath + "\\_360手机万号采集_"+ textBox2.Text.Trim() +"_"+ DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xls";
            StringHelp.pathError = spath + "\\_360手机万号采集_"+ textBox2.Text.Trim() + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".txt";
            //创建excel文件
            StringHelp.CreateExcel();

            if (textBox1.Text.Trim() == "" || textBox2.Text.Trim() == "")
            {
                MessageBox.Show("城市或号段不能为空！", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            dt = DateTime.Now;
            label3.Text = dt + "正在执行……";
            z4Str = textBox2.Text;

            ThreadPool.SetMaxThreads(5, 5); //设置最大线程数
            StringHelp.Write(StringHelp.pathError, DateTime.Now.ToString() + "\r\n--------------------------------"+ textBox2.Text+ "号段开始-------------------------------------" + "\r\n");

            //string strSql = "select SUBSTRING(Number7,1,3) q3w from TelphoneData where SUBSTRING(Number7,4,4)=" + textBox2.Text.Trim() + " and city='" + textBox1.Text.Trim() + "' AND SUBSTRING(Number7,1,3) NOT IN ('170','171')";
            //DataTable returnDt = SqlHelp.bangding(strSql);
            //for (int i = 0; i < returnDt.Rows.Count; i++)
            //{
            //    ThreadPool.QueueUserWorkItem(new WaitCallback(mainWhile), returnDt.Rows[i][0].ToString());
            //}
            #region 测试代码
            //SELECT max(手机号) '最大尾号',count(1) '个数' FROM _360手机万号采集_6549_20170913110445155 GROUP BY SUBSTRING(手机号, 1, 3)
            //SELECT max(Telphone) '最大尾号',count(1) '个数' FROM TelCollection_0539 GROUP BY begin3
            //DELETE FROM TelCollection_0539 WHERE begin3 NOT IN (135,137)
            //DELETE FROM TelCollection_0539 WHERE SUBSTRING(Telphone, 8, 4) > 204
            //delete FROM TelCollection_0539 WHERE SUBSTRING(Telphone, 8, 4) > 7100 AND begin3 NOT IN (135,137)
            //6539号段
            //int[] q = { 131,134,133,136,150,151,152,158,159,178,182,187,188 };
            //0539号段
            //int[] q = { 131, 132, 133, 134, 135, 136, 137, 138, 139, 147, 150, 151, 152, 153, 155, 156, 157, 158, 159, 173, 175, 176, 177, 178, 180, 181, 182, 183, 184, 185, 186, 187, 188, 189 };
            int[] q = {  155, 156, 157, 158, 159, 173, 175, 176, 177, 178, 180, 181, 182, 183, 184, 185, 186, 187, 188, 189 };//147, 150, 151, 152, 153,
            for (int i = 0; i < q.Length; i++)
            {
                ThreadPool.QueueUserWorkItem(new WaitCallback(mainWhile), q[i]);
            }
            #endregion

            //是否结束所有线程
            rhw = ThreadPool.RegisterWaitForSingleObject(new System.Threading.AutoResetEvent(false), CheckThreadPool, null, 5000, false);
        }

        private void CheckThreadPool(object state, bool timeout)
        {
            //锁，只能有一个线程执行此方法
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
                    label4.Text = end.ToString() + "执行完毕,已将检索结果导出到与导入文件所在的相同目录。";
                    MessageBox.Show(end.ToString() + "执行结束。" + "用时：" + DateHelp.DateDiff(end, dt) + "。" +
                        "\r\n获取个数：" + num + ", 已将检索结果导出到与导入文件所在的相同目录。\r\n");
                }
            }
        }

        void mainWhile(object q3Str)
        {
            DateTime hdStart = DateTime.Now;
            int hdnum = 0;
            string txt = q3Str + z4Str + "0000";
            long hd = Convert.ToInt64(txt);
            for (int i = 0; i < 10000; i++)
            {
                string keyword = (hd + i).ToString();
                int bl = CollectRule.MainWhile(search, keyword);
                if (bl == 1)
                {
                    num++;
                    hdnum++;
                }
                else if (bl == -1)
                {
                    MessageBox.Show(DateTime.Now.ToString() + "执行到：" + keyword + "时，当前IP开始被" + search + "屏蔽,未启动VPN！\r\n");
                }
            }

            DateTime hdEnd = DateTime.Now;
            StringHelp.Write(StringHelp.pathError, hdEnd.ToString() + " 号段：" + q3Str + "结束。" + "用时：" + DateHelp.DateDiff(hdEnd, hdStart) + "获取个数：" + hdnum + "\r\n");
        }

        private void _360SJWH_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Dispose(true);
            this.Close();
            Application.Exit();
        }
    }
}
