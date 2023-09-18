using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using System.Diagnostics;
using System.Text;
using System;
using System.Runtime.InteropServices.ComTypes;

namespace QueryExcel
{
    internal class Excel
    {
        public DataSet data;

        public Excel(){}

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

        public void ExportDataToExcel(DataGridView dgv)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "导出Excel文件";
            saveFileDialog.Filter = "Microsoft Office Excel 工作簿 (*.xls)|*.xls";
            saveFileDialog.FilterIndex = 1;
            saveFileDialog.AddExtension = true;
            saveFileDialog.RestoreDirectory = true;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string localFilePath = saveFileDialog.FileName.ToString();
                int TotalCount;
                int RowRead = 0;
                int Percent = 0;
                TotalCount = dgv.Rows.Count;

                Stream myStream = saveFileDialog.OpenFile();
                StreamWriter sw = new StreamWriter(myStream, Encoding.Unicode);//Unicode编码可避免显示乱码问题
                string strHeader = "";
                Stopwatch timer = new Stopwatch();
                timer.Start();

                try
                {
                    for (int i = 0; i < dgv.Columns.Count; i++)
                    {
                        if (i > 0)
                        {
                            strHeader += "\t";
                        }
                        if (dgv.Columns[i].HeaderText.ToString() != null)
                        {
                            strHeader += dgv.Columns[i].HeaderText.ToString();
                        }
                        else
                        {
                            strHeader += "";
                        }
                    }
                    sw.WriteLine(strHeader);

                    for (int i = 0; i < dgv.Rows.Count; i++)
                    {
                        RowRead++;
                        Percent = 100 * RowRead / TotalCount;
                        string strData = "";

                        for (int j = 0; j < dgv.Columns.Count; j++)
                        {
                            if (j > 0)
                            {
                                strData += "\t";
                            }

                            if (dgv.Rows[i].Cells[j].Value != null)
                            {
                                strData += dgv.Rows[i].Cells[j].Value.ToString();
                            }
                            else
                            {
                                strData += "";
                            }
                        }
                        sw.WriteLine(strData);
                    }

                    sw.Close();
                    myStream.Close();
                    timer.Reset();
                    timer.Stop();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                finally
                {
                    sw.Close();
                    myStream.Close();
                    timer.Stop();
                }

                if (MessageBox.Show("导出成功，是否立即打开？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(localFilePath);
                }
            }
        }
    }
}
