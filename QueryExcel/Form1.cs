using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace QueryExcel
{
    public partial class Form1 : Form
    {
        private string previousText = "";

        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 描述：当对象被拖动到窗体边界时，触发事件； || 实现：获取拖拽到窗体中的文件路径。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_DragEnter(object sender, DragEventArgs e) //获得“信息”
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All; //重要代码：表明是所有类型的数据，比如文件路径
            else
                e.Effect = DragDropEffects.None;
        }

        /// <summary>
        /// 描述：当你在窗体上释放拖动的对象时，触发事件； || 实现：获取拖拽到窗体中的文件路径。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_DragDrop(object sender, DragEventArgs e) //解析信息
        {
            string path = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString(); //获得路径
            textBox1.Text = path; //由一个textBox显示路径
        }

        /// <summary>
        /// 描述：当文本框内容发生更改时，触发事件； || 实现：校验文件路径。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
            TextBox textBox = textBox1;
            string filePath = textBox.Text;
            if (string.IsNullOrEmpty(filePath) == false)
            {
                string fileType = Path.GetExtension(filePath);
                if (fileType == ".xls" || fileType == ".xlsx")
                {
                    // 如果是Excel文件路径，执行相应的操作
                    previousText = filePath; // 保存当前的值，以便下次使用
                }
                else
                {
                    // 如果不是Excel文件路径，执行相应的操作
                    textBox.Text = previousText; // 将文本框的值重置为上一次的值
                }
            }
        }

        /// <summary>
        /// 描述：点击按钮时，触发事件； || 实现：读取指定文件路径的Excel表格数据。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button1_Click(object sender, EventArgs e)
        {
            //try
            //{
                // 在这里放置可能会抛出异常的代码

                //获取文件路径

                string filePath = textBox1.Text.ToString(); // 声明ListBox对象

                if (string.IsNullOrEmpty(filePath))
                {
                    MessageBox.Show("请先输入文件路径，或将文件拖拽到本窗口获取文件路径！");
                }
                else
                {
                    dataSet1.Clear(); // 清空dataSetd对象
                    listBox1.Items.Clear(); // 清空listBox1内容

                    ExcelHandler excelHandler = new();
                    dataSet1 = excelHandler.ImportExcelToDataSet(textBox1.Text);  // 读取Excel内容到dataSetd对象

                    for (int i = 0; i <= dataSet1.Tables.Count - 1; i++) // 读取tabcontrol除第一页页面外的所有TabPage标题到listBox1中
                    {
                        listBox1.Items.Add(dataSet1.Tables[i].TableName);
                    }
                }
            //}
            //catch (Exception ex)
            //{
            //    // 在这里处理异常
            //    MessageBox.Show(ex.Message);
            //}
        }

        /// <summary>
        /// 描述：当列表框中选择的项发生更改时，触发事件； || 实现：读取列表框选项对应的表数据，达到数据对象的切换。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ListBox lb = sender as ListBox; // 声明ListBox控件事件触发对象

            string tableName = lb.SelectedItem.ToString();// 获取listbox的选中值

            if (dataSet1.Tables.Contains(tableName)) // 检查DataSet中是否存在这个DataTable
            {
                this.dataGridView1.DataSource = null;
                this.dataGridView1.DataSource = dataSet1.Tables[tableName]; // 将DataTable绑定到DataGridView

                // 将Datable中所有列名绑定到bindingSource
                List<string> columnNames = new();
                foreach (DataColumn column in dataSet1.Tables[tableName].Columns)
                {
                    columnNames.Add(column.ColumnName);
                }

                this.bindingSource1.DataSource = null;
                this.bindingSource1.DataSource = columnNames;
            }
            else
            {
                MessageBox.Show("未找到名为 " + tableName + " 的DataTable");
            }
        }

        /// <summary>
        /// 描述：点击按钮时，触发事件； || 实现：导出DataGridView数据到新的Excel文件。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button4_Click(object sender, EventArgs e)
        {
            ExcelHandler excelHandler = new();
            excelHandler.ExportExcel(dataGridView3);
        }

        /// <summary>
        /// 描述：点击按钮时，触发事件； || 实现：导出DataSet数据到新的Excel文件。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button5_Click(object sender, EventArgs e)
        {
            ExcelHandler excelHandler = new();
            excelHandler.ExportExcel(dataSet1);
        }

        /// <summary>
        /// 描述：点击按钮时，触发事件； || 实现：查询DataTable数据到DataGridView。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button2_Click(object sender, EventArgs e)
        {
            /*
             代码调试中......
             */

            string filter = ""; //过滤条件字符串

            //遍历每一行,但不包括最后新建行 || 自定义筛选条件
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                if (!row.IsNewRow)
                {
                    // 获取列的值
                    string Value1 = (string)row.Cells["Column1"].Value;
                    string Value2 = (string)row.Cells["Column2"].Value;
                    string Value3 = (string)row.Cells["Column3"].Value;
                    string Value4 = (string)row.Cells["Column4"].Value;
                    string Value5 = (string)row.Cells["Column5"].Value;

                    //MessageBox.Show($"{Value1} - {Value2} - {Value3} - {Value4} - {Value5}");

                    // 处理字符串,根据比较运算符定义查询字符串
                    if (!string.IsNullOrEmpty(Value2))
                    {
                        // 使用switch语句
                        switch (Value2)
                        {
                            case "等于":
                                filter += string.Format("{0} = '{1}'", Value1, Value3);
                                break;
                            case "不等于":
                                filter += string.Format("{0} <> '{1}'", Value1, Value3);
                                break;
                            case "包含":
                                filter += string.Format("({0} LIKE '%{1}%')", Value1, Value3);
                                break;
                            case "不包含":
                                filter += string.Format("({0} NOT LIKE '%{1}%')", Value1, Value3);
                                break;
                            case "大于等于":
                                filter += string.Format("({0} >= '{1}')", Value1, Value3);
                                break;
                            case "小于等于":
                                filter += string.Format("({0} <= '{1}')", Value1, Value3);
                                break;
                            case "大于":
                                filter += string.Format("({0} > '{1}')", Value1, Value3);
                                break;
                            case "小于":
                                filter += string.Format("({0} < '{1}')", Value1, Value3);
                                break;
                            case "在...之内":
                                filter += string.Format("({0} >= '{1}' AND {0} <= '{2}')", Value1, Value3, Value4);
                                break;
                            case "在...之外":
                                filter += string.Format("({0} < '{1}' OR {0} > '{2}')", Value1, Value3, Value4);
                                break;
                            // 如果columnName不等于上述任何值
                            default:
                                break;
                        }
                    }
                    //根据行数判断是否添加多条件连接符
                    if (!string.IsNullOrEmpty(Value5))
                    {
                        // 使用switch语句
                        switch (Value5)
                        {
                            case "且":
                                filter = $"{filter} AND ";
                                break;
                            case "或":
                                filter = $"{filter} OR ";
                                break;
                            // 如果columnName不等于上述任何值
                            default:
                                break;
                        }
                    }
                }
            }

            //获取DataTable数据
            string tableName = listBox1.SelectedItem.ToString();

            Query query = new();
            DataTable dataTable = query.Select(dataSet1.Tables[tableName], filter);

            //开始过滤数据
            dataGridView3.DataSource = null;
            dataGridView3.DataSource = dataTable;
        }

        /// <summary>
        /// 清空自定义数据筛选条件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView2.Rows.Clear();
        }
    }
}
