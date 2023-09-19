using System.IO;
using System.Data;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;

namespace QueryExcel
{
    internal class Excel
    {
        /// <summary>
        /// 静态构造函数
        /// </summary>
        public Excel()
        {
        
        }

        /// <summary>
        /// 函数，读取Excel文件数据到Dataset数据集
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns>DataSet数据集</returns>
        public DataSet LoadExcelToDataSet(string filePath)
        {
            var dataSet = new DataSet();

            // 打开Excel文件
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = Path.GetExtension(filePath) != ".xls" ? new XSSFWorkbook(stream) : new HSSFWorkbook(stream);

                for (int k = 0; k < workbook.NumberOfSheets; k++)
                {
                    var sheet = workbook.GetSheetAt(k);
                    var dataTable = new DataTable(sheet.SheetName);

                    // 获取所有已使用的单元格最后一列的列数
                    int cellCount = 0;
                    for (int i = sheet.LastRowNum; i >= 0; i--)
                    {
                        var row = sheet.GetRow(i);
                        if (row != null)
                        {
                            cellCount = row.LastCellNum;
                            break;
                        }
                    }

                    // 读取表头
                    var headerRow = sheet.GetRow(0);
                    if (headerRow == null)
                    {
                        headerRow = sheet.CreateRow(0);
                    }
                    for (int i = 0; i < cellCount; i++)
                    {
                        var cell = headerRow.GetCell(i);
                        if (cell == null || string.IsNullOrEmpty(cell.ToString())) // 如果表头单元格为空，则创建默认列名
                        {
                            cell = headerRow.CreateCell(i);
                            cell.SetCellValue($"Column{GetColumnName(i + 1)}");
                        }
                        dataTable.Columns.Add(cell.ToString());
                    }

                    // 读取数据行
                    for (int i = 1; i <= sheet.LastRowNum; i++)
                    {
                        var row = sheet.GetRow(i);
                        if (row == null) continue;

                        var dataRow = dataTable.NewRow();
                        for (int j = 0; j < dataTable.Columns.Count; j++) // 只处理包含数据的单元格
                        {
                            var cell = row.GetCell(j);
                            if (cell == null) continue;

                            // 根据单元格的数据类型进行转换
                            dataRow[j] = GetCellValue(cell);
                        }

                        dataTable.Rows.Add(dataRow);
                    }

                    dataSet.Tables.Add(dataTable);
                }
            }

            return dataSet;
        }

        /// <summary>
        /// 函数，导出DataGridView数据集到Excel文件
        /// </summary>
        /// <param name="dataGridView"></param>
        public void SaveDataGridViewToExcel(DataGridView dataGridView)
        {
            // 创建一个文件选择器
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx",
                FilterIndex = 2,
                RestoreDirectory = true,
                FileName = "ExportedData.xlsx" // 默认的文件名
            };

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                var filePath = saveFileDialog.FileName;

                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("Sheet1");

                // 写入表头
                var headerRow = sheet.CreateRow(0);
                for (int i = 0; i < dataGridView.Columns.Count; i++)
                {
                    var cell = headerRow.CreateCell(i);
                    cell.SetCellValue(dataGridView.Columns[i].HeaderText);
                }

                // 写入数据行
                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    var row = sheet.CreateRow(i + 1);
                    for (int j = 0; j < dataGridView.Columns.Count; j++)
                    {
                        var cell = row.CreateCell(j);

                        // 根据单元格的数据类型进行转换
                        if (dataGridView[j, i].Value != null)
                        {
                            switch (dataGridView[j, i].ValueType.Name)
                            {
                                case "String":
                                    cell.SetCellValue(dataGridView[j, i].Value.ToString());
                                    break;
                                case "DateTime":
                                    cell.SetCellValue((DateTime)dataGridView[j, i].Value);
                                    break;
                                case "Boolean":
                                    cell.SetCellValue((bool)dataGridView[j, i].Value);
                                    break;
                                case "Int32":
                                case "Double":
                                case "Decimal":
                                    cell.SetCellType(CellType.Numeric); // 设置单元格类型为数字
                                    cell.SetCellValue(Convert.ToDouble(dataGridView[j, i].Value));
                                    break;
                                default:
                                    cell.SetCellValue(dataGridView[j, i].Value.ToString());
                                    break;
                            }
                        }
                    }
                }

                // 保存到文件
                using (var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(stream);
                }

                MessageBox.Show("文件已成功保存！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); // 显示消息框提示
            }
        }

        /// <summary>
        /// 用于获取单元格的值 || 根据单元格的类型来返回相应的值
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private object GetCellValue(ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.String:
                case CellType.Formula when cell.CachedFormulaResultType == CellType.String:
                    // 如果单元格是字符串类型，返回字符串值
                    return cell.StringCellValue;
                case CellType.Numeric:
                case CellType.Formula when cell.CachedFormulaResultType == CellType.Numeric:
                    // 如果单元格是数字类型，根据是否是日期格式返回日期值或数字值
                    return DateUtil.IsCellDateFormatted(cell) ? cell.DateCellValue : cell.NumericCellValue;
                case CellType.Boolean:
                case CellType.Formula when cell.CachedFormulaResultType == CellType.Boolean:
                    // 如果单元格是布尔类型，返回布尔值
                    return cell.BooleanCellValue;
                default:
                    // 对于其他类型，返回"未知类型"
                    return null;
            }
        }

        /// <summary>
        /// 用于根据列的索引生成列名 || 它使用了Excel列名的命名规则，即A-Z之后是AA-AZ，然后是BA-BZ，以此类推
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        private string GetColumnName(int index)
        {
            int dividend = index;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
    }
}
