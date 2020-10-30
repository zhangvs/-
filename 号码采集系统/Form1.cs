using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string dateStr = DateTime.Now.ToString("yyyyMMdd");

        private void Form1_Load(object sender, EventArgs e)
        {


            //string keyword = "05398827666";
            //if (keyword.Substring(0,1)=="0")
            //{
            //    int k = keyword.IndexOf('-');
            //    if (k > 0)
            //    {
            //        keyword = keyword.Substring(k + 1, 7);
            //    }
            //    else
            //    {
            //        keyword = keyword.Substring(4, 7);
            //    }
            //}


            //bool dd = Regex.IsMatch(" 小区", @"[\u4e00-\u9fa5]+(广场|大街|大厦|小区|(软件|工业|软件)园|交汇|有限公司|有限责任公司|地址：|地址:)");

            string rem = "使用说明 手机号码: 五行数理说明:号码 13305390000 (山东省临沂市 山东电信CDMA卡) 的数理为: 80 ,凶星入度,清本缩小,一生皆苦之孤独空虚数 (大凶) ,其暗示...";
            bool bl=CollectRule.ruleJudge(rem, "13305390000");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //测试，将excel中的sheet1导入到sqlserver中
            string connString = "Data Source=47.93.253.194;Initial Catalog=HZSoftFramework_Base_2016;User ID=sa;Password=system@123";
            System.Windows.Forms.OpenFileDialog fd = new OpenFileDialog();
            if (fd.ShowDialog() == DialogResult.OK)
            {
                TransferData(fd.FileName, "Sheet0", connString);
            }

            //string filePath = fd.FileName;

            //IWorkbook ssfworkbook;
            //using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            //{
            //    var fileExtension = Path.GetExtension(filePath);
            //    if (fileExtension == ".xls")
            //    {
            //        ssfworkbook = new HSSFWorkbook(file);
            //    }
            //    else if (fileExtension == ".xlsx")
            //    {
            //        ssfworkbook = new XSSFWorkbook(file);
            //    }
            //    else
            //    {
            //        MessageBox.Show("文件类型不支持");
            //        return;
            //    }
            //}
            //StringHelp.CreateExcel();
            //var sheet = ssfworkbook.GetSheetAt(0);

            //DataTable dt = new DataTable();

            ////默认，第一行是字段
            //IRow headRow = sheet.GetRow(0);
            //dt.Columns.Add("Telphone");

            ////DataRow titleRow = dt.NewRow();
            ////titleRow[0] = "Telphone";
            ////遍历列
            //for (int i = headRow.FirstCellNum, len = headRow.LastCellNum; i < len; i++)
            //{
            //    //遍历行
            //    for (int j = (sheet.FirstRowNum), r = sheet.LastRowNum + 1; j < r; j++)
            //    {
            //        ICell cell = sheet.GetRow(j).GetCell(i);

            //        if (cell != null)
            //        {
            //            DataRow dataRow = dt.NewRow();
            //            switch (cell.CellType)
            //            {
            //                case CellType.String:
            //                    dataRow[0] = cell.StringCellValue;
            //                    dt.Rows.Add(dataRow);
            //                    break;
            //                case CellType.Numeric:
            //                    dataRow[0] = cell.NumericCellValue;
            //                    dt.Rows.Add(dataRow);
            //                    break;
            //                case CellType.Boolean:
            //                    break;
            //                default:
            //                    break;
            //            }
            //        }
            //    }
            //}
            //using (System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(connString))
            //{
            //    bcp.SqlRowsCopied += new System.Data.SqlClient.SqlRowsCopiedEventHandler(bcp_SqlRowsCopied);
            //    bcp.BatchSize = 100;//每次传输的行数
            //    bcp.NotifyAfter = 100;//进度提示的行数
            //    bcp.DestinationTableName = "TelphoneExcel0539";//目标表
            //    bcp.WriteToServer(dt);
            //}
            MessageBox.Show("完成！");
        }
        public void TransferData(string excelFile, string sheetName, string connectionString)
        {
            DataSet ds = new DataSet();
            try
            {
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
                string strSql = string.Format("if object_id('{0}') is null create table {0}(", sheetName);
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
                    bcp.DestinationTableName = sheetName;//目标表
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

        private void button2_Click(object sender, EventArgs e)
        {
            string sql = "";
            for (int i = 0; i < 100; i++)
            {
                sql += ad("0539", i * 100, (i + 1) * 100);
            }
            string ss = sql.Substring(0,sql.Length-10);

            SqlHelp.connStr = "Data Source=47.93.253.194;Initial Catalog=HZSoftFramework_Base_2016;User ID=sa;Password=system@123";
            DataTable returnDt = SqlHelp.bangding(ss);

            dataGridView1.DataSource = returnDt;

        }
        
        private string ad(string hd,int q,int e)
        {
            string dd = "SELECT '"+hd+"' hd, '"+q+"-"+e+"' qj, count(1) count FROM TelphoneSource WHERE SUBSTRING(Telphone, 8, 4) > "+q+" AND SUBSTRING(Telphone, 8, 4)<= "+e+" AND SUBSTRING(Telphone,4, 4)= "+hd+ "\r\n UNION ALL ";
            return dd;
        }
        

        private void button3_Click(object sender, EventArgs e)
        {
            string strSql = "";
            for (int i = 0; i < 100; i++)
            {

                strSql = "SELECT SUBSTRING(Telphone, 9, 3)%100 '序号',Telphone '电话',CASE sellMark WHEN 0 THEN '未售' ELSE '已售' end '是否已售' FROM TelphoneSource WHERE SUBSTRING(Telphone, 8, 4) > " + i * 100 + " AND SUBSTRING(Telphone, 8, 4)<= " + (i + 1) * 100 + " AND SUBSTRING(Telphone,4, 4)= 0539";
                DataTable returnDt = SqlHelp.bangding(strSql);
                for (int r = 0; r < returnDt.Rows.Count; r++)
                {

                }
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                System.Diagnostics.Process.Start("iexplore.exe", dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString());
            }
            else if (e.ColumnIndex == 2)
            {
                System.Diagnostics.Process.Start("iexplore.exe", dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString());
            }
        }
    }
}
