using NPOI.POIFS.FileSystem;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq.Expressions;
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
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            TextBox textBox = sender as TextBox;
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
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
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

                    Excel excel = new Excel(); //初始化Excel类
                    dataSet1 = excel.LoadExcelToDataSet(textBox1.Text);  // 读取Excel内容到dataSetd对象

                    for (int i = 0; i <= dataSet1.Tables.Count - 1; i++) // 读取tabcontrol除第一页页面外的所有TabPage标题到listBox1中
                    {
                        listBox1.Items.Add(dataSet1.Tables[i].TableName);
                    }
                }
            }
            catch (Exception ex)
            {
                // 在这里处理异常
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 描述：当列表框中选择的项发生更改时，触发事件； || 实现：读取列表框选项对应的表数据，达到数据对象的切换。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ListBox lb = sender as ListBox; // 声明ListBox控件事件触发对象

            string tableName = lb.SelectedItem.ToString();// 获取listbox的选中值

            if (dataSet1.Tables.Contains(tableName)) // 检查DataSet中是否存在这个DataTable
            {
                this.dataGridView1.DataSource = dataSet1.Tables[tableName]; // 将DataTable绑定到DataGridView

                // 将Datable中所有列名绑定到bindingSource
                List<string> columnNames = new List<string>();
                foreach (DataColumn column in dataSet1.Tables[tableName].Columns)
                {
                    columnNames.Add(column.ColumnName);
                }
                this.bindingSource1.DataSource = columnNames;
            }
            else
            {
                MessageBox.Show("未找到名为 " + tableName + " 的DataTable");
            }
        }

        /// <summary>
        /// 描述：点击按钮时，触发事件； || 实现：导出数据到新的Excel文件。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            Excel excel = new Excel();
            excel.SaveDataGridViewToExcel(dataGridView1);
        }
    }
}
