
' VER 02

Imports System.Data.Odbc
Imports System.Data.Common.DbException
Imports System.IO

Module Module1

    Sub Main(ByVal feed() As String)
        Dim command As String = feed(0)

        '// For BULK update ////////////////////

        '  Dim missing, line As String
        '  missing = "C:\David\WIP\missing.txt"
        '  FileOpen(50, missing, OpenMode.Input, OpenAccess.Read)
        '  Do Until EOF(50)
        '  line = LineInput(50)
        '  If line.Length > 10 Then
        '  Try
        '  updateDev(line)
        '  Catch
        '  End Try
        '  End If
        '  Loop
        '  FileClose(50)
        '
        '/////////////////////////////

        ' command = "INSERT INTO tiff_status (time_stamp, order_num, au_num, del_date, pool, thickness, surface, order_date, work_days, downloaded, del_ori, mask_color, large_order) VALUES (637846769847355258,'david','david-1.mht','16/03/2018','1','1_00mm_material','3','09/03/2018','6',1,'16/03/2018','GREEN','0')"

        'command = "INSERT INTO tiff_status (time_stamp, order_num, au_num, del_date, pool, thickness, surface, work_days, order_date, downloaded, del_ori, rek_reason, mask_color) VALUES (637850187310943822,'12344323','12344323.mht','29/04/2022','1','1_60mm_material','3','2','09/04/2022',.T.,'29/04/2022','','GREEN')"

        command = command.Replace(".T.", "1")
        command = command.Replace(".F.", "0")

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
                fail = File.AppendText("T:\IN_HOUSE_SOFTWARE\ALL_SOFTWARE\Visual_Studio_2010\OMS-Download\TEMP\OMS_Dev_SQL_Update.txt")
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
                fail2 = File.AppendText("T:\IN_HOUSE_SOFTWARE\ALL_SOFTWARE\Visual_Studio_2010\OMS-Download\TEMP\OMS_Dev_SQL_Update.txt")
                fail2.WriteLine("DatabaseSQL_Update Cannot Execute")
                fail2.WriteLine(sqlCommand.CommandText)
                fail2.Close()
            End Try
            sqlConnection.Close()
        End If

    End Sub

End Module



