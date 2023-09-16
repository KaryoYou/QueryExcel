using System;
using System.IO;
using System.Data;
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
            if (textBox != null)
            {
                string filePath = textBox.Text;
                Excel excel = new Excel(filePath);
                if (excel.isExcelFile)
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

        private void comboBox1_DropDown(object sender, EventArgs e)
        {
            string filePath = textBox1.Text; // Excel文件路径
            Excel excel = new Excel(filePath);
            if (excel.isExcelFile)
            {
                comboBox1.DataSource = excel.excelSheets; // 将数据绑定到comboBox1
                comboBox1.DisplayMember = "TABLE_NAME";
            }
        }
    }
}
