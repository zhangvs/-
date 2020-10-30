using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Net;
using System.Text;
using System.Windows.Forms;
using System.Net.Sockets;

namespace 号码采集系统
{
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            //System.Net.IPAddress[] addressList = Dns.GetHostByName(Dns.GetHostName()).AddressList;
            string strHostName = Dns.GetHostName(); //得到本机的主机名
            IPHostEntry IpEntry = Dns.GetHostEntry(strHostName); //取得本机IP
            string strAddr = "";
            for (int i = 0; i < IpEntry.AddressList.Length; i++)
            {
                //从IP地址列表中筛选出IPv4类型的IP地址
                //AddressFamily.InterNetwork表示此IP为IPv4,
                //AddressFamily.InterNetworkV6表示此地址为IPv6类型
                if (IpEntry.AddressList[i].AddressFamily == AddressFamily.InterNetwork)
                {
                    strAddr= IpEntry.AddressList[i].ToString();
                }
            }

            string strSql = "select * from LoginUser where UserName='" + txtUserName.Text + "' and Password='" + txtPassword.Text + "' and Computer='" + strHostName + "'";//' and ipAdress='" + strAddr + "
            DataTable returnDt = SqlHelp.bangding(strSql);
            if (returnDt.Rows.Count>=1)
            {
                Main main = new Main();
                main.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("登录失败！");
            }
        }
    }
}
