
Imports System.Data.OleDb



Module Module1

    '  Dim info_file As String


    Sub Main(ByVal feed() As String)
        Dim au_num As String = feed(0)

        'au_num = "'david','david-1.mht','16/03/2018','1','1_00mm_material','3','09/03/2018','6',.T.,'16/03/2018','GREEN','0','34565','34568'"

        Dim stamp As Date
        Dim timeStamp As Long
        stamp = Now
        timeStamp = CLng(stamp.Ticks)

        au_num = timeStamp & "," & au_num

        Try
            update(au_num)
        Catch
        End Try

        Try
            updateSQL(au_num)
        Catch
        End Try

leave_sub:

    End Sub
    Sub updateSQL(ByVal input As String)
        REM ********************* DATABASE ********************************************
        Dim status_table As String = "tiff_status"
        Dim sqlCommand As String
        sqlCommand = "INSERT INTO " & status_table & " (time_stamp,order_num,au_num,del_date,pool,thickness,surface,order_date,work_days,downloaded,del_ori,mask_color,large_order,base_ref_number,ref_number) VALUES (" & input & ")"

        If InStr(sqlCommand, ".T.") Then
            sqlCommand = sqlCommand.Replace(".T.", 1)
        End If
        If InStr(sqlCommand, ".F.") Then
            sqlCommand = sqlCommand.Replace(".F.", 0)
        End If

        Dim id As Integer
        id = Shell("S:\Job\in_house_software\OMS_Dev_SQL_update " & Chr(34) & sqlCommand & Chr(34))
        'id = Shell("D:\Work\Job\in_house_software\OMS_Dev_SQL_update_server " & Chr(34) & sqlCommand & Chr(34))

    End Sub
    Sub update(ByVal input As String)
        REM ********************* DATABASE ********************************************
        Dim status_table As String = "t:\database3\tiff_status.dbf"
        'Dim status_table As String = "D:\Storage\database3\tiff_status.dbf"
        Dim oConnString As String = "Provider=VFPOLEDB.1;Data Source= " + status_table
        Dim oCommandText As String

        oCommandText = "INSERT INTO " & status_table & " (time_stamp,order_num,au_num,del_date,pool,thickness,surface,order_date,work_days,downloaded,del_ori,mask_color,large_order,base_ref_number,ref_number) VALUES (" & input & ")"

        Dim omyConnection As New OleDbConnection(oConnString)

        Try
            omyConnection.Open()
        Catch
            GoTo close_and_leave
        End Try

        Try
            Dim omyCommand As New OleDbCommand(oCommandText, omyConnection)
            omyCommand.ExecuteNonQuery()
        Catch
            omyConnection.Close()
        End Try

close_and_leave:
        omyConnection.Close()


    End Sub

End Module



