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

                    // 读取表头
                    var headerRow = sheet.GetRow(0);
                    for (int i = 0; i < headerRow.Cells.Count; i++)
                    {
                        dataTable.Columns.Add(headerRow.Cells[i].ToString());
                    }

                    // 读取数据行
                    for (int i = 1; i <= sheet.LastRowNum; i++)
                    {
                        var row = sheet.GetRow(i);
                        if (row == null) continue;

                        var dataRow = dataTable.NewRow();
                        for (int j = 0; j < row.Cells.Count; j++)
                        {
                            var cell = row.GetCell(j);
                            if (cell == null) continue;

                            // 根据单元格的数据类型进行转换
                            switch (cell.CellType)
                            {
                                case CellType.String:
                                    dataRow[j] = cell.StringCellValue;
                                    break;
                                case CellType.Numeric:
                                    dataRow[j] = DateUtil.IsCellDateFormatted(cell) ? cell.DateCellValue : cell.NumericCellValue;
                                    break;
                                case CellType.Boolean:
                                    dataRow[j] = cell.BooleanCellValue;
                                    break;
                                case CellType.Formula: // 新增对公式单元格的处理
                                    switch (cell.CachedFormulaResultType)
                                    {
                                        case CellType.String:
                                            dataRow[j] = cell.StringCellValue;
                                            break;
                                        case CellType.Numeric:
                                            dataRow[j] = DateUtil.IsCellDateFormatted(cell) ? cell.DateCellValue : cell.NumericCellValue;
                                            break;
                                        case CellType.Boolean:
                                            dataRow[j] = cell.BooleanCellValue;
                                            break;
                                        default:
                                            dataRow[j] = "未知类型";
                                            break;
                                    }
                                    break;
                                default:
                                    dataRow[j] = cell.ToString();
                                    break;
                            }
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

    }
}
