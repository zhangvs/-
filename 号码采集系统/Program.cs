using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;

namespace 号码采集系统
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            //int ahpvno = 73;
            //int v_no = 18;
            //for (int i = 0; i < 100; i++)
            //{
            //    StringHelp.RequestReferer("https://club.autohome.com.cn/o/bbs/thread/d846d5111aef1fdc/90932303-1.html", ahpvno++, v_no++);

            //    //Process.Start("iexplore.exe", "https://club.autohome.com.cn/bbs/thread/d846d5111aef1fdc/90932303-1.html");
            //}

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new ShuaLiang());
        }
    }
}
