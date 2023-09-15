Imports System.Data.OleDb

Namespace Sql

    Public Class ExcelHelper
        Private connectionString As String

        Public Sub New(filePath As String)
            ' Excel 8.0 为 Excel 97 版本
            Me.connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & filePath & ";Extended Properties=""Excel 8.0;HDR=YES;IMEX=1"""
        End Sub

        Public Function ExecuteQuery(sheetName As String) As DataTable
            Using connection As New OleDbConnection(connectionString)
                connection.Open()

                Dim query As String = "SELECT * FROM [" & sheetName & "$]"
                Using command As New OleDbCommand(query, connection)
                    Using reader As OleDbDataReader = command.ExecuteReader()
                        Dim result As New DataTable()
                        result.Load(reader)
                        Return result
                    End Using
                End Using
            End Using
        End Function
    End Class
End Namespace