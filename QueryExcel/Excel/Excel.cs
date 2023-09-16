using System.Data;
using System.Data.OleDb;

namespace QueryExcel
{
    internal class Excel
    {
        public DataSet data;

        public Excel(string filePath)
        {
            data = LoadExcelFile(filePath);
        }

        private DataSet LoadExcelFile(string filePath)
        {
            // 连接字符串
            string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1'";

            // 创建连接
            OleDbConnection conn = new OleDbConnection(connString);

            // 打开连接
            conn.Open();

            // 获取Excel文件中所有工作表的名字
            DataTable dtSheetName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            // 声明DataSet1对象
            DataSet ds = new DataSet();

            // 遍历所有工作表
            foreach (DataRow dr in dtSheetName.Rows)
            {
                string sheetName = dr["TABLE_NAME"].ToString();

                // 查询语句
                string sql = "select * from [" + sheetName + "]";

                // 执行查询语句
                OleDbDataAdapter adp = new OleDbDataAdapter(sql, conn);

                // 创建一个新的DataTable来存储工作表数据
                DataTable dt = new DataTable();

                // 将查询结果填充到DataTable中
                adp.Fill(dt);

                // 将DataTable添加到DataSet中，并以工作表名命名
                ds.Tables.Add(dt);
                ds.Tables[ds.Tables.Count - 1].TableName = sheetName.Replace("$", "");
            }

            // 关闭连接
            conn.Close();

            return ds;
        }
    }
}
