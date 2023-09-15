'使用Microsoft.ACE.OLEDB.12.0作为数据连接时，必须安装access驱动程序，成功安装安装包后仍需启用IIS发布配置
'由于配置繁琐，不适用小白用户，所以避免使用Microsoft.ACE.OLEDB
'https://download.microsoft.com/download/E/4/2/E4220252-5FAE-4F0A-B1B9-0B48B5FBCCF9/AccessDatabaseEngine_X64.exe
'https://download.microsoft.com/download/E/4/2/E4220252-5FAE-4F0A-B1B9-0B48B5FBCCF9/AccessDatabaseEngine.exe

Imports Microsoft.Office.Interop

Namespace MsExcel

    Public Class ExcelHelper
        Public Function GetAllSheetName(ByVal strFilePath As String) As String()

            Dim AppXls As Excel.Application '声明Excel对象
            Dim AppWokBook As Excel.Workbook '声明工作簿对象
            Dim AppSheet As Excel.Worksheet '声明工作表对象

            AppXls = New Excel.Application '实例化Excel对象
            AppXls.Workbooks.Open(strFilePath) '打开已经存在的EXCEL文件
            AppXls.Visible = False '使Excel不可见

            AppWokBook = AppXls.Workbooks(1) 'AppWokBook对象指向工作簿"C:\\Users\\25042\\Desktop\\ReadExcelTest.xls"
            AppSheet = AppWokBook.Sheets("Sheet1") 'AppSheet对象指向AppWokBook对象中的表“Sheet1”

            '要读取数据表"Sheet1"中的单元格“A1”的值，到变量S1里
            Dim S1 As String
            S1 = AppXls.Workbooks(1).Sheets("Sheet1").Range("A1").Value
            MsgBox(S1)

            '使用完毕必须关闭EXCEL，并退出
            AppXls.ActiveWorkbook.Close(SaveChanges:=True)
            AppXls.Quit()
        End Function
    End Class
End Namespace