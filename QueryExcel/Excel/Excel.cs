using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;
using System.Windows.Forms;
using MiniExcelLibs;

public class ExcelHandler
{
    /// <summary>
    /// 将Excel文件导入到DataSet中
    /// </summary>
    /// <param name="filePath">文件路径</param>
    /// <returns>DataSet对象</returns>
    public DataSet ImportExcelToDataSet(string filePath)
    {
        var dataSet = new DataSet();
        bool isReadSuccessful = false;

        try
        {
            // 获取所有工作表的名称
            var sheetNames = MiniExcel.GetSheetNames(filePath);

            foreach (var sheetName in sheetNames)
            {
                // 为每个工作表获取行的集合
                // 使用MiniExcel的Query方法获取工作表的数据
                var rows = MiniExcel.Query(filePath, sheetName: sheetName, useHeaderRow: false).ToList();
                isReadSuccessful = true;

                // 为每个工作表创建一个新的DataTable，并使用工作表名称命名
                var dataTable = new DataTable(sheetName);

                // 标记是否为第一行（列名行）
                bool isFirstRow = true;

                // 遍历每一行
                foreach (var row in rows)
                {
                    var rowDictionary = (IDictionary<string, object>)row;

                    // 如果是第一行（列名行）
                    if (isFirstRow)
                    {
                        // 遍历每一列
                        foreach (var column in rowDictionary.Keys)
                        {
                            // 如果单元格为空，则使用单元格索引作为列名；否则，使用单元格值作为列名
                            string columnName = rowDictionary[column] != null ? rowDictionary[column].ToString() : column;
                            // 将列添加到DataTable中
                            dataTable.Columns.Add(columnName);
                        }

                        // 标记已处理完第一行（列名行）
                        isFirstRow = false;
                    }
                    else  // 如果不是第一行（数据行）
                    {
                        // 创建一个新的DataRow
                        DataRow dataRow = dataTable.NewRow();

                        // 遍历每一列
                        int columnIndex = 0;
                        foreach (var column in rowDictionary.Keys)
                        {
                            // 将单元格的值添加到DataRow中
                            dataRow[columnIndex] = rowDictionary[column];
                            columnIndex++;
                        }

                        // 将DataRow添加到DataTable中
                        dataTable.Rows.Add(dataRow);
                    }
                }

                // 将DataTable添加到DataSet中
                dataSet.Tables.Add(dataTable);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"读取失败, {ex.Message}", "错误信息", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (!isReadSuccessful)
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
