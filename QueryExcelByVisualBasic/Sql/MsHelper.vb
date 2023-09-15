Imports System.Data.SqlClient

'使用示例
'Dim helper As New Helper()
'Dim result As DataTable = helper.ExecuteQuery("SELECT * FROM YourTable")

Namespace Sql
    Public Class MsHelper
        Private ReadOnly connectionString As String

        Public Sub New()
            Me.connectionString = "Data Source=DATASERVER;Initial Catalog=DATABASE;Integrated Security=True;User ID=YourId;Password=YourPassword;Connect Timeout=30;"
        End Sub

        Public Function ExecuteQuery(sql As String) As DataTable
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                If Not connection.State Then
                    'MessageBox.Show("连接失败", "数据库连接状态", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return Nothing
                    Exit Function
                End If

                Using command As New SqlCommand(sql, connection)
                    Using reader As SqlDataReader = command.ExecuteReader()
                        Dim result As New DataTable()
                        result.Load(reader)
                        Return result
                    End Using
                End Using
            End Using
        End Function

        Public Sub ExecuteNonQuery(sql As String)
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                If Not connection.State Then
                    'MessageBox.Show("连接失败", "数据库连接状态", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                Using command As New SqlCommand(sql, connection)
                    command.ExecuteNonQuery()
                End Using
            End Using
        End Sub
    End Class
End Namespace