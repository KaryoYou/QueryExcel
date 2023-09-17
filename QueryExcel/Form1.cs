using System;
using System.Data;
using System.IO;
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

        private void Form1_DragEnter(object sender, DragEventArgs e) //获得“信息”
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All; //重要代码：表明是所有类型的数据，比如文件路径
            else
                e.Effect = DragDropEffects.None;
        }

        private void Form1_DragDrop(object sender, DragEventArgs e) //解析信息
        {
            string path = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString(); //获得路径
            textBox1.Text = path; //由一个textBox显示路径
        }

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

        private void button1_Click(object sender, EventArgs e)
        {
            ListBox lb = listBox1; // 声明ListBox对象

            dataSet1.Clear(); // 清空dataSetd对象
            lb.Items.Clear(); // 清空listBox1内容

            
            Excel excel = new Excel(textBox1.Text); //初始化Excel类
            dataSet1 = excel.data;  // 读取Excel内容到dataSetd对象

            for (int i = 0; i <= dataSet1.Tables.Count - 1; i++) // 读取tabcontrol除第一页页面外的所有TabPage标题到listBox1中
            {
                lb.Items.Add(dataSet1.Tables[i].TableName);
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ListBox lb = sender as ListBox; // 声明ListBox控件事件触发对象

            string tableName = lb.SelectedItem.ToString();// 获取listbox的选中值

            if (dataSet1.Tables.Contains(tableName)) // 检查DataSet中是否存在这个DataTable
            {
                this.dataGridView1.DataSource = dataSet1.Tables[tableName]; // 将DataTable绑定到DataGridView
            }
            else
            {
                MessageBox.Show("未找到名为 " + tableName + " 的DataTable");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Excel excel = new Excel();
            excel.ExportDataToExcel(dataGridView1);
        }
    }
}
