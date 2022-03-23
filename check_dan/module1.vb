
' VER 02

Imports System.Data.Odbc
Imports System.Data.Common.DbException
Imports System.IO

Module Module1

    Sub Main(ByVal feed() As String)
        Dim command As String = feed(0)


        'newValue = "108.99"
        'au = "00196P01"

        Try
            updateDev(command) 'Dev
        Catch
        End Try


leave_sub:

    End Sub



    Sub updateDev(ByVal command As String)

        Dim sqlConnection As New OdbcConnection
        sqlConnection = New OdbcConnection("DSN=OMS_Dev;MultipleActiveResultSets=True")
        Dim sqlCommand As New OdbcCommand
        sqlCommand.CommandText = command
        sqlCommand.Connection = sqlConnection

        Try
            If sqlConnection.State = ConnectionState.Closed Then
                sqlConnection.Open()
            End If
        Catch e As System.Exception
            Try
                Dim fail As StreamWriter
                'fail = File.AppendText("D:\storage\IN_HOUSE_SOFTWARE\ALL_SOFTWARE\Visual_Studio_2010\OMS-Download\TEMP\OrdersInfoSQL_Update.txt")
                fail = File.AppendText("T:\IN_HOUSE_SOFTWARE\ALL_SOFTWARE\Visual_Studio_2010\OMS-Download\TEMP\OrdersInfoSQL_Update.txt")
                fail.WriteLine("DatabaseSQL_UpdateCannot Connect")
                fail.WriteLine(sqlCommand.CommandText)
                fail.WriteLine(e.Message)
                fail.Close()
            Catch
            End Try
        End Try

        If sqlConnection.State = ConnectionState.Open Then
            Try
                sqlCommand.ExecuteNonQuery()
            Catch
                sqlConnection.Close()
                Dim fail2 As StreamWriter
                'fail2 = File.AppendText("D:\storage\IN_HOUSE_SOFTWARE\ALL_SOFTWARE\Visual_Studio_2010\OMS-Download\TEMP\OrdersInfoSQL_Update.txt")
                fail2 = File.AppendText("T:\IN_HOUSE_SOFTWARE\ALL_SOFTWARE\Visual_Studio_2010\OMS-Download\TEMP\OrdersInfoSQL_Update.txt")
                fail2.WriteLine("DatabaseSQL_Update Cannot Execute")
                fail2.WriteLine(sqlCommand.CommandText)
                fail2.Close()
            End Try
            sqlConnection.Close()
        End If

    End Sub

End Module



