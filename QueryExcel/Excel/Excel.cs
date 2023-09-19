using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Windows.Forms;
using MiniExcelLibs;
using System.Linq;

public class ExcelHandler
{
    /// <summary>
    /// 导入函数
    /// </summary>
    /// <param name="filePath"></param>
    /// <returns>DataSet对象</returns>
    public DataSet ImportExcelToDataSet(string filePath)
    {
        // 创建一个新的DataSet
        DataSet dataSet = new DataSet();

        bool readState = false;

        try
        {
            // 获取Excel文件中所有的工作表名
            var sheetNames = MiniExcel.GetSheetNames(filePath);

            readState = true;

            // 遍历每一个工作表
            foreach (var sheetName in sheetNames)
            {
                // 使用MiniExcel的Query方法获取工作表的数据
                var data = MiniExcel.Query(filePath, sheetName: sheetName, useHeaderRow: false).ToList();

                // 创建一个新的DataTable，并将工作表的名字设为DataTable的TableName
                DataTable dataTable = new DataTable(sheetName);

                // 将数据添加到DataTable中
                bool columnsDefined = false;
                foreach (var row in data)
                {
                    var rowDict = (IDictionary<string, object>)row;
                    if (!columnsDefined)
                    {
                        foreach (var column in rowDict.Keys)
                        {
                            string columnName = rowDict[column] != null ? rowDict[column].ToString() : column;
                            dataTable.Columns.Add(columnName);
                        }
                        columnsDefined = true;
                        continue;
                    }

                    DataRow dataRow = dataTable.NewRow();
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        if (rowDict.ContainsKey(column.ColumnName) && rowDict[column.ColumnName] != null)
                        {
                            dataRow[column.ColumnName] = rowDict[column.ColumnName];
                        }
                    }

                    // 添加行到DataTable
                    dataTable.Rows.Add(dataRow);
                }

                // 将DataTable添加到DataSet
                dataSet.Tables.Add(dataTable);
            }
        }
        catch (Exception ex)
        {
            // 如果导出过程中发生错误，显示错误消息框
            MessageBox.Show($"读取失败: {ex.Message}", "错误信息", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (!readState)
            {
                MessageBox.Show($"请确认文件格式是否兼容表中的所有内容，\n" +
                    $"不兼容的文件格式会影响文件结构的完整性，本程序将无法正常读取。\n" +
                    $"建议：使用Office Excel应用程序完成修复：文件->信息->转换", "提示信息",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        return dataSet;
    }

    /// <summary>
    /// 导出函数
    /// </summary>
    /// <param name="dataGridView"></param>
    public void ExportExcel(object dataSource)
    {
        // 创建一个保存文件对话框
        SaveFileDialog saveFileDialog = new SaveFileDialog();
        saveFileDialog.Filter = "Excel Files|*.xlsx";
        saveFileDialog.Title = "Save an Excel File";
        saveFileDialog.FileName = "default.xlsx";  // 默认文件名

        // 显示保存文件对话框
        if (saveFileDialog.ShowDialog() == DialogResult.OK)
        {
            // 获取用户选择的文件路径
            string filePath = saveFileDialog.FileName;

            try
            {
                // 如果文件已存在，删除它
                if (System.IO.File.Exists(filePath))
                {
                    System.IO.File.Delete(filePath);
                }

                // 检查数据源类型并执行相应操作
                if (dataSource is DataGridView)
                {
                    DataTable dataTable = (DataTable)((DataGridView)dataSource).DataSource;
                    MiniExcel.SaveAs(filePath, dataTable, printHeader: true, sheetName:dataTable.TableName);
                }
                else if (dataSource is DataSet)
                {
                    MiniExcel.SaveAs(filePath, dataSource, printHeader: true);
                }
                else
                {
                    throw new ArgumentException("无效的数据源类型。数据源必须是 DataGridView 或 DataSet。");
                }

                // 显示导出成功的消息框
                MessageBox.Show("导出成功!", "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                // 如果导出过程中发生错误，显示错误消息框
                MessageBox.Show($"导出失败: {ex.Message}", "错误信息", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
