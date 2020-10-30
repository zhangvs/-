using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace 号码采集系统
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            baiduSJWH form = new baiduSJWH();
            form.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            _360SJWH form = new _360SJWH();
            form.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            _360SJPL form = new _360SJPL();
            form.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            baiduSJPL form = new baiduSJPL();
            form.Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            _360GHWH form = new _360GHWH();
            form.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            baiduGHWH form = new baiduGHWH();
            form.Show();
        }


        private void button7_Click(object sender, EventArgs e)
        {
            _360GHPL form = new _360GHPL();
            form.Show();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            baiduGHPL form = new baiduGHPL();
            form.Show();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            string connString = "Data Source=47.93.253.194;Initial Catalog=HZSoftFramework_Base_2016;User ID=sa;Password=system@123";
            System.Windows.Forms.OpenFileDialog fd = new OpenFileDialog();
            if (fd.ShowDialog() == DialogResult.OK)
            {
                TransferData(fd.FileName, "Sheet0", connString);
            }
        }

        public void TransferData(string excelFile, string sheetName, string connectionString)
        {
            DataSet ds = new DataSet();
            try
            {
                string fileName=Path.GetFileNameWithoutExtension(excelFile);
                string fileExt = Path.GetExtension(excelFile);
                string connStr = "";
                if (fileExt == ".xls")
                {
                    connStr = "Provider = Microsoft.Jet.OLEDB.4.0 ; Data Source =" + excelFile + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'";
                }
                else
                {
                    connStr = "Provider = Microsoft.ACE.OLEDB.12.0 ; Data Source =" + excelFile + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1'";
                }
                //获取全部数据
                //string strConn = "Provider=Microsoft.Jet.OLEDB.12.0;" + "Data Source=" + excelFile + ";" + "Extended Properties=Excel 12.0;";
                OleDbConnection conn = new OleDbConnection(connStr);
                conn.Open();
                string strExcel = "";
                OleDbDataAdapter myCommand = null;
                strExcel = string.Format("select * from [{0}$]", sheetName);
                myCommand = new OleDbDataAdapter(strExcel, connStr);
                myCommand.Fill(ds, sheetName);
                //如果目标表不存在则创建
                string strSql = string.Format("if object_id('{0}') is null create table {0}(", fileName);
                foreach (System.Data.DataColumn c in ds.Tables[0].Columns)
                {
                    strSql += string.Format("[{0}] nvarchar(1000),", c.ColumnName);
                }
                strSql = strSql.Trim(',') + ")";
                using (System.Data.SqlClient.SqlConnection sqlconn = new System.Data.SqlClient.SqlConnection(connectionString))
                {
                    sqlconn.Open();
                    System.Data.SqlClient.SqlCommand command = sqlconn.CreateCommand();
                    command.CommandText = strSql;
                    command.ExecuteNonQuery();
                    sqlconn.Close();
                }
                //用bcp导入数据
                using (System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(connectionString))
                {
                    bcp.SqlRowsCopied += new System.Data.SqlClient.SqlRowsCopiedEventHandler(bcp_SqlRowsCopied);
                    bcp.BatchSize = 100;//每次传输的行数
                    bcp.NotifyAfter = 100;//进度提示的行数
                    bcp.DestinationTableName = fileName;//目标表
                    bcp.WriteToServer(ds.Tables[0]);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
        //进度显示
        void bcp_SqlRowsCopied(object sender, System.Data.SqlClient.SqlRowsCopiedEventArgs e)
        {
            this.Text = e.RowsCopied.ToString();
            this.Update();
        }

    }
}
