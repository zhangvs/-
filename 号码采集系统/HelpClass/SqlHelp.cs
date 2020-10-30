using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Text;

namespace 号码采集系统
{
    class SqlHelp
    {

       public static string connStr = "Data Source=47.93.253.194;Initial Catalog=lywenkaiData;User ID=sa;Password=system@123";// pooling=false

        public static int ExecuteSql(string strSql)
        {
            using (SqlConnection conn = new SqlConnection(connStr))
            {
                using (SqlCommand cmd = new SqlCommand(strSql, conn))
                {
                    try
                    {
                        int rows = 0;
                        if (conn.State != ConnectionState.Open)
                        {
                             conn.Open();
                        }
                        Object locker = new Object();
                        lock (locker)
                        {
                            rows = cmd.ExecuteNonQuery();
                        }
                        return rows;
                    }
                    catch (SqlException ex)
                    {
                        int rows = 0;
                        do
                        {
                            StringHelp.Write(StringHelp.pathError, DateTime.Now.ToString() + strSql + ex.Message + "\r\n");
                            System.Threading.Thread.Sleep(3 * 1000);
                            StringHelp.getState();
                            rows = ExecuteSql(strSql);
                        } while (rows<=0);
                        StringHelp.Write(StringHelp.pathError, DateTime.Now.ToString() +"重新执行\r\n");
                        return rows;
                    }
                }
            }
        }

        public static DataTable bangding(string sqlsel)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connStr))
                {
                    DataTable dt = new DataTable();
                    SqlDataAdapter da = new SqlDataAdapter(sqlsel, conn);
                    da.Fill(dt);
                    return dt;
                }
            }
            catch (SqlException ex)
            {
                DataTable dt;
                do
                {
                    StringHelp.Write(StringHelp.pathError, DateTime.Now.ToString() + sqlsel + ex.Message + "\r\n");
                    System.Threading.Thread.Sleep(3 * 1000);
                    StringHelp.getState();
                    dt = bangding(sqlsel);
                } while (dt == null);
                StringHelp.Write(StringHelp.pathError, DateTime.Now.ToString() + "重新执行\r\n");
                return dt;
            }

        }

        public static string getCreateTable(string tableName)
        {
            string createSql = @"
IF OBJECT_ID ('dbo.TelCollection') IS NOT NULL
	DROP TABLE dbo.TelCollection 

CREATE TABLE dbo.TelCollection
	(
	ID         BIGINT IDENTITY NOT NULL,
	Belong     INT,
	Search     NVARCHAR (20),
	Middle4    NCHAR (10),
	Begin3     NCHAR (10),
	Telphone   NCHAR (20),
	Title      NVARCHAR (200),
	TitleLink  NVARCHAR (500),
	Abstract   NVARCHAR (1000),
	Address    NVARCHAR (500),
	Coordinate NVARCHAR (50),
	Track      NVARCHAR (500),
	Satate     INT,
    UserID     NVARCHAR (50),
    UserName   NVARCHAR (50),
	CreateDate NCHAR (10),
	CreateTime DATETIME CONSTRAINT DF_TelCollection_CreateTime DEFAULT (getdate()),
	CONSTRAINT PK_TelCollection PRIMARY KEY (ID)
	)";
            createSql = createSql.Replace("TelCollection", tableName);
            return createSql;

        }

        /// </summary>
        /// <param name="connectionString">目标连接字符</param>
        /// <param name="TableName">目标表</param>
        /// <param name="dt">源数据</param>
        private void SqlBulkCopyByDatatable( DataTable dt)
        {
            using (SqlConnection conn = new SqlConnection(connStr))
            {
                using (SqlBulkCopy sqlbulkcopy = new SqlBulkCopy(connStr, SqlBulkCopyOptions.UseInternalTransaction))
                {
                    try
                    {
                        sqlbulkcopy.DestinationTableName = "CollectionData";
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            sqlbulkcopy.ColumnMappings.Add(dt.Columns[i].ColumnName, dt.Columns[i].ColumnName);
                        }
                        sqlbulkcopy.WriteToServer(dt);
                    }
                    catch (System.Exception ex)
                    {
                        throw ex;
                    }
                }
            }
        }
    }
}
