'Imports EmisWio2.MsSql

Imports System.IO

Public Class Form1
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' 打开文件选择对话框
        OpenFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            ' 将选中的文件路径添加到ComboBox中
            ComboBox1.Items.Add(OpenFileDialog1.FileName)
            ' 将选中的文件路径设置为ComboBox的当前选项
            ComboBox1.SelectedItem = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub ComboBox1_DropDown(sender As Object, e As EventArgs) Handles ComboBox1.DropDown
        ' 如果ComboBox为空，则打开文件选择对话框
        ' 如果ComboBox非空，则获取当前文件路径下的所有Excel文件
        If Not String.IsNullOrEmpty(ComboBox1.SelectedItem) Then
            Dim directoryPath As String = Path.GetDirectoryName(ComboBox1.SelectedItem.ToString())
            Dim excelFiles As String() = Directory.GetFiles(directoryPath, "*.xl*")
            ComboBox1.Items.Clear()
            For Each file In excelFiles
                ComboBox1.Items.Add(file)
            Next
            If ComboBox1.Items.Count > 0 Then
                ComboBox1.SelectedIndex = 0
            End If
        End If
    End Sub

    Private Sub ComboBox1_MouseDown(sender As Object, e As MouseEventArgs) Handles ComboBox1.MouseDown
        ' 如果ComboBox非空，则不显示下拉选项框,转到获取当前文件路径下的所有Excel文件的按键点击事件
        If ComboBox1.Items.Count = 0 Then
            'MessageBox.Show("未指定Excel文件路径", "文件路径选项", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ComboBox1.AllowDrop = False
            Button3.PerformClick()
        Else
            ComboBox1.AllowDrop = True
        End If
    End Sub

    Private Sub ComboBox2_MouseDown(sender As Object, e As MouseEventArgs) Handles ComboBox2.MouseDown
        If String.IsNullOrEmpty(ComboBox1.SelectedItem) Then
            MessageBox.Show("未指定Excel文件路径", "工作簿选项", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ComboBox2.AllowDrop = False
        Else
            ComboBox2.AllowDrop = True
        End If
    End Sub
End Class
