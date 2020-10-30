using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
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
    public partial class baiduGHPL : Form
    {
        public baiduGHPL()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;
        }

        string search = "baidu";
        DateTime dt;
        RegisteredWaitHandle rhw;
        int num = 0;

        private void btn_Upload_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            var dlg = sender as OpenFileDialog;
            txtFileName.Text = dlg.FileName;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var filePath = "";
            if (!File.Exists(txtFileName.Text))
            {
                MessageBox.Show("文件不存在。");
                return;
            }
            else if (textBox2.Text == "")
            {
                MessageBox.Show("城市不能为空。");
                return;
            }
            else
            {
                filePath = openFileDialog1.FileName;//txtFileName.Text;
                //读取excel
                IWorkbook ssfworkbook;
                using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    var fileExtension = Path.GetExtension(filePath);
                    if (fileExtension == ".xls")
                    {
                        ssfworkbook = new HSSFWorkbook(file);
                        StringHelp.filePathOut = filePath.Substring(0, filePath.Length - 4) + "_baidu固话批量检索结果_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xls";
                        StringHelp.pathError = filePath.Substring(0, filePath.Length - 5) + "_baidu固话批量执行日志_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".txt";
                    }
                    else if (fileExtension == ".xlsx")
                    {
                        ssfworkbook = new XSSFWorkbook(file);
                        StringHelp.filePathOut = filePath.Substring(0, filePath.Length - 5) + "_baidu固话批量检索结果_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xls";
                        StringHelp.pathError = filePath.Substring(0, filePath.Length - 5) + "_baidu固话批量执行日志_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".txt";
                    }
                    else
                    {
                        MessageBox.Show("文件类型不支持");
                        return;
                    }
                }
                StringHelp.CreateExcel();
                //for (int i = 0; i < ssfworkbook.NumberOfSheets; i++)//遍历薄
                //{
                var sheet = ssfworkbook.GetSheetAt(0);
                dt = DateTime.Now;
                int daoNum = sheet.LastRowNum + 1;
                label3.Text = "导入" + daoNum + "个号码。";
                label4.Text = dt + "正在采集……";
                
                ThreadPool.SetMaxThreads(4, 4); //设置最大线程数
                List<Thread> arr = new List<Thread>();
                for (int j = 0; j < daoNum; j++)  //LastRowNum 获取的是最后一行的编号（编号从0开始）。getPhysicalNumberOfRows()获取的是物理行数，也就是不包括那些空行（隔行）的情况。
                {
                    var row = sheet.GetRow(j);  //读取当前行数据
                    if (row.GetCell(0).ToString() != "")
                    {
                        ThreadPool.QueueUserWorkItem(new WaitCallback(mainWhile), textBox2.Text + row.GetCell(0).ToString());
                        ThreadPool.QueueUserWorkItem(new WaitCallback(mainWhile), textBox2.Text +"-"+ row.GetCell(0).ToString());
                    }
                }
                //}
            }

            rhw = ThreadPool.RegisterWaitForSingleObject(new System.Threading.AutoResetEvent(false), CheckThreadPool, null, 10000, false);
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
                    label4.Text = end.ToString() + "执行完毕,已将检索结果导出到与导入文件所在的相同目录。";
                    MessageBox.Show(end.ToString() + "执行结束。" + "用时：" + DateHelp.DateDiff(end, dt) + "。" +
                        "\r\n获取个数：" + num + ", 已将检索结果导出到与导入文件所在的相同目录。\r\n");
                }
            }
        }

        void mainWhile(object tel)
        {
            string keyword = tel.ToString();
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
