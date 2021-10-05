Imports System.Data.SqlClient
Module basNDO
    Public MyNavConnection As SQLConnection
    Public bNavConn As Boolean = False

    Public Sub ConnectNavDB()
        ' Dim MyNavConString As String = "DSN=EverHome;UID=sa"
        Dim MyNavConString As String = My.Settings.NavConnectionString
        'MsgBox(MyNavConString)
        MyNavConnection = New SQLConnection(MyNavConString)
        MyNavConnection.Open()
        bNavConn = True
    End Sub

    Public Sub DisconnectNavDB()
        If bNavConn = True Then
            MyNavConnection.Close()
        End If
        bNavConn = False
    End Sub

    Public Sub ExecuteNavSQL(ByVal mySQL As String)
        '  Try
        Dim MyCommand As New SQLCommand
        If bNavConn = False Then ConnectDB()
        MyCommand.Connection = MyNavConnection
        MyCommand.CommandText = mySQL
        MyCommand.CommandTimeout = 0
        MyCommand.ExecuteNonQuery()
        MyCommand.Dispose()
        'Catch ex As Exception
        'System.IO.File.AppendAllText("C:\ErrorLog.txt", Date.Now & " From BasNDO: " & vbCrLf & ex.Message & vbCrLf)
        'End Try
    End Sub

    Public Function ReadNavRecord(ByVal mySQL As String) As SqlDataReader
        Dim MyCommand As New SqlCommand
        If bNavConn = False Then ConnectDB()
        Try
            MyCommand.Connection = MyNavConnection
            MyCommand.CommandText = mySQL
            MyCommand.CommandTimeout = 0
            ReadNavRecord = MyCommand.ExecuteReader
            MyCommand.Dispose()
        Catch MySQLException As SqlException
            MsgBox(MySQLException.ToString)
            Return Nothing
        Catch MyException As Exception
            MsgBox(MyException.ToString)
            Return Nothing
        End Try
    End Function
End Module
