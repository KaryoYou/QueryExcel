<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        OpenFileDialog1 = New OpenFileDialog()
        TabPage5 = New TabPage()
        Button3 = New Button()
        ListBox1 = New ListBox()
        DataGridView1 = New DataGridView()
        Button2 = New Button()
        Button1 = New Button()
        ComboBox7 = New ComboBox()
        Label7 = New Label()
        ComboBox6 = New ComboBox()
        Label6 = New Label()
        ComboBox5 = New ComboBox()
        Label5 = New Label()
        ComboBox4 = New ComboBox()
        Label4 = New Label()
        ComboBox3 = New ComboBox()
        Label3 = New Label()
        ComboBox2 = New ComboBox()
        Label2 = New Label()
        ComboBox1 = New ComboBox()
        Label1 = New Label()
        TabControl1 = New TabControl()
        TabPage5.SuspendLayout()
        CType(DataGridView1, ComponentModel.ISupportInitialize).BeginInit()
        TabControl1.SuspendLayout()
        SuspendLayout()
        ' 
        ' OpenFileDialog1
        ' 
        OpenFileDialog1.FileName = "OpenFileDialog1"
        ' 
        ' TabPage5
        ' 
        TabPage5.Controls.Add(Button3)
        TabPage5.Controls.Add(ListBox1)
        TabPage5.Controls.Add(DataGridView1)
        TabPage5.Controls.Add(Button2)
        TabPage5.Controls.Add(Button1)
        TabPage5.Controls.Add(ComboBox7)
        TabPage5.Controls.Add(Label7)
        TabPage5.Controls.Add(ComboBox6)
        TabPage5.Controls.Add(Label6)
        TabPage5.Controls.Add(ComboBox5)
        TabPage5.Controls.Add(Label5)
        TabPage5.Controls.Add(ComboBox4)
        TabPage5.Controls.Add(Label4)
        TabPage5.Controls.Add(ComboBox3)
        TabPage5.Controls.Add(Label3)
        TabPage5.Controls.Add(ComboBox2)
        TabPage5.Controls.Add(Label2)
        TabPage5.Controls.Add(ComboBox1)
        TabPage5.Controls.Add(Label1)
        TabPage5.Location = New Point(4, 26)
        TabPage5.Name = "TabPage5"
        TabPage5.Size = New Size(852, 497)
        TabPage5.TabIndex = 4
        TabPage5.Text = "Excel数据查询"
        TabPage5.UseVisualStyleBackColor = True
        ' 
        ' Button3
        ' 
        Button3.Location = New Point(73, 9)
        Button3.Name = "Button3"
        Button3.Size = New Size(28, 17)
        Button3.TabIndex = 36
        Button3.Text = "```"
        Button3.TextAlign = ContentAlignment.TopCenter
        Button3.UseVisualStyleBackColor = True
        ' 
        ' ListBox1
        ' 
        ListBox1.FormattingEnabled = True
        ListBox1.ItemHeight = 17
        ListBox1.Location = New Point(11, 108)
        ListBox1.Name = "ListBox1"
        ListBox1.Size = New Size(827, 55)
        ListBox1.TabIndex = 35
        ' 
        ' DataGridView1
        ' 
        DataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridView1.Location = New Point(11, 169)
        DataGridView1.Name = "DataGridView1"
        DataGridView1.RowTemplate.Height = 25
        DataGridView1.Size = New Size(827, 325)
        DataGridView1.TabIndex = 34
        ' 
        ' Button2
        ' 
        Button2.Location = New Point(646, 77)
        Button2.Name = "Button2"
        Button2.Size = New Size(85, 25)
        Button2.TabIndex = 32
        Button2.Text = "增加条件"
        Button2.UseVisualStyleBackColor = True
        ' 
        ' Button1
        ' 
        Button1.Location = New Point(737, 77)
        Button1.Name = "Button1"
        Button1.Size = New Size(85, 25)
        Button1.TabIndex = 31
        Button1.Text = "清空条件"
        Button1.UseVisualStyleBackColor = True
        ' 
        ' ComboBox7
        ' 
        ComboBox7.FormattingEnabled = True
        ComboBox7.Location = New Point(519, 77)
        ComboBox7.Name = "ComboBox7"
        ComboBox7.Size = New Size(121, 25)
        ComboBox7.TabIndex = 30
        ' 
        ' Label7
        ' 
        Label7.AutoSize = True
        Label7.Location = New Point(519, 57)
        Label7.Name = "Label7"
        Label7.Size = New Size(56, 17)
        Label7.TabIndex = 29
        Label7.Text = "多查条件"
        ' 
        ' ComboBox6
        ' 
        ComboBox6.FormattingEnabled = True
        ComboBox6.Location = New Point(392, 77)
        ComboBox6.Name = "ComboBox6"
        ComboBox6.Size = New Size(121, 25)
        ComboBox6.TabIndex = 28
        ' 
        ' Label6
        ' 
        Label6.AutoSize = True
        Label6.Location = New Point(392, 57)
        Label6.Name = "Label6"
        Label6.Size = New Size(27, 17)
        Label6.TabIndex = 27
        Label6.Text = "值2"
        ' 
        ' ComboBox5
        ' 
        ComboBox5.FormattingEnabled = True
        ComboBox5.Location = New Point(265, 77)
        ComboBox5.Name = "ComboBox5"
        ComboBox5.Size = New Size(121, 25)
        ComboBox5.TabIndex = 26
        ' 
        ' Label5
        ' 
        Label5.AutoSize = True
        Label5.Location = New Point(265, 57)
        Label5.Name = "Label5"
        Label5.Size = New Size(27, 17)
        Label5.TabIndex = 25
        Label5.Text = "值1"
        ' 
        ' ComboBox4
        ' 
        ComboBox4.FormattingEnabled = True
        ComboBox4.Location = New Point(138, 77)
        ComboBox4.Name = "ComboBox4"
        ComboBox4.Size = New Size(121, 25)
        ComboBox4.TabIndex = 24
        ' 
        ' Label4
        ' 
        Label4.AutoSize = True
        Label4.Location = New Point(138, 57)
        Label4.Name = "Label4"
        Label4.Size = New Size(56, 17)
        Label4.TabIndex = 23
        Label4.Text = "比较条件"
        ' 
        ' ComboBox3
        ' 
        ComboBox3.FormattingEnabled = True
        ComboBox3.Location = New Point(11, 77)
        ComboBox3.Name = "ComboBox3"
        ComboBox3.Size = New Size(121, 25)
        ComboBox3.TabIndex = 22
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(11, 57)
        Label3.Name = "Label3"
        Label3.Size = New Size(32, 17)
        Label3.TabIndex = 21
        Label3.Text = "列名"
        ' 
        ' ComboBox2
        ' 
        ComboBox2.FormattingEnabled = True
        ComboBox2.Location = New Point(265, 29)
        ComboBox2.Name = "ComboBox2"
        ComboBox2.Size = New Size(248, 25)
        ComboBox2.TabIndex = 20
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(265, 9)
        Label2.Name = "Label2"
        Label2.Size = New Size(44, 17)
        Label2.TabIndex = 19
        Label2.Text = "工作簿"
        ' 
        ' ComboBox1
        ' 
        ComboBox1.FormattingEnabled = True
        ComboBox1.Location = New Point(11, 29)
        ComboBox1.Name = "ComboBox1"
        ComboBox1.RightToLeft = RightToLeft.No
        ComboBox1.Size = New Size(248, 25)
        ComboBox1.TabIndex = 18
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(11, 9)
        Label1.Name = "Label1"
        Label1.Size = New Size(56, 17)
        Label1.TabIndex = 17
        Label1.Text = "文件路径"
        ' 
        ' TabControl1
        ' 
        TabControl1.Controls.Add(TabPage5)
        TabControl1.Location = New Point(12, 12)
        TabControl1.Name = "TabControl1"
        TabControl1.SelectedIndex = 0
        TabControl1.Size = New Size(860, 527)
        TabControl1.TabIndex = 1
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 17F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(884, 561)
        Controls.Add(TabControl1)
        Name = "Form1"
        Text = "Excel数据查询"
        TabPage5.ResumeLayout(False)
        TabPage5.PerformLayout()
        CType(DataGridView1, ComponentModel.ISupportInitialize).EndInit()
        TabControl1.ResumeLayout(False)
        ResumeLayout(False)
    End Sub
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents TabPage5 As TabPage
    Friend WithEvents Button3 As Button
    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Button2 As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents ComboBox7 As ComboBox
    Friend WithEvents Label7 As Label
    Friend WithEvents ComboBox6 As ComboBox
    Friend WithEvents Label6 As Label
    Friend WithEvents ComboBox5 As ComboBox
    Friend WithEvents Label5 As Label
    Friend WithEvents ComboBox4 As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents ComboBox3 As ComboBox
    Friend WithEvents Label3 As Label
    Friend WithEvents ComboBox2 As ComboBox
    Friend WithEvents Label2 As Label
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents TabControl1 As TabControl
End Class
