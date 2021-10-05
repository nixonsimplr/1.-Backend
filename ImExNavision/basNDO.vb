Imports System.Data.SqlClient
Module basNDO
    Public MyNavConnection As SqlConnection
    Public bNavConn As Boolean = False
    Public MyNavConnection1 As SqlConnection
    Public MyNavConnection2 As SqlConnection
    Public bNavConn1 As Boolean = False
    Public bNavConn2 As Boolean = False

    Public Sub ConnectNavDB()
        ' Dim MyNavConString As String = "DSN=EverHome;UID=sa"
        Dim MyNavConString As String = My.Settings.NavConnectionString
        'MsgBox(MyNavConString)
        MyNavConnection = New SqlConnection(MyNavConString)
        MyNavConnection.Open()
        bNavConn = True
        MyNavConnection1 = New SqlConnection(MyNavConString)
        MyNavConnection1.Open()
        bNavConn1 = True
        MyNavConnection2 = New SqlConnection(MyNavConString)
        MyNavConnection2.Open()
        bNavConn2 = True
    End Sub

    Public Sub DisconnectNavDB()
        If bNavConn = True Then
            MyNavConnection.Close()
        End If
        bNavConn = False
    End Sub
    Public Sub ExecuteNavAnotherSQL(ByVal mySQL As String)
        '  Try
        Dim MyCommand As New SqlCommand
        If bNavConn1 = False Then ConnectNavDB()
        MyCommand.Connection = MyNavConnection1
        MyCommand.CommandText = mySQL
        MyCommand.CommandTimeout = 0
        MyCommand.ExecuteNonQuery()
        'MyCommand.Connection.Close()
        'MyCommand.Dispose()

        'Catch ex As Exception
        'System.IO.File.AppendAllText("C:\ErrorLog.txt", Date.Now & " From BasNDO: " & vbCrLf & ex.Message & vbCrLf)
        'End Try
    End Sub
    Public Sub ExecuteNavSQL(ByVal mySQL As String)
        '  Try
        Dim MyCommand As New SqlCommand
        If bNavConn = False Then ConnectNavDB()
        MyCommand.Connection = MyNavConnection
        MyCommand.CommandText = mySQL
        MyCommand.CommandTimeout = 0
        MyCommand.ExecuteNonQuery()
        'MyCommand.Connection.Close()
        'MyCommand.Dispose()

        'Catch ex As Exception
        'System.IO.File.AppendAllText("C:\ErrorLog.txt", Date.Now & " From BasNDO: " & vbCrLf & ex.Message & vbCrLf)
        'End Try
    End Sub

    Public Function ReadNavRecord(ByVal mySQL As String) As SqlDataReader
        Dim MyCommand As New SqlCommand
        If bNavConn = False Then ConnectNavDB()
        Try
            'ConnectNavDB()
            MyCommand.Connection = MyNavConnection
            MyCommand.CommandText = mySQL
            MyCommand.CommandTimeout = 0
            ReadNavRecord = MyCommand.ExecuteReader
            MyCommand.Dispose()
        Catch MySQLException As SqlException
            'MsgBox(MySQLException.ToString)
            Return Nothing
        Catch MyException As Exception
            'MsgBox(MyException.ToString)
            Return Nothing
        End Try
    End Function
    Public Function ReadNavRecordAnother(ByVal mySQL As String) As SqlDataReader
        Dim MyCommand As New SqlCommand
        If bNavConn = False Then ConnectNavDB()
        Try
            'ConnectNavDB()
            MyCommand.Connection = MyNavConnection1
            MyCommand.CommandText = mySQL
            MyCommand.CommandTimeout = 0
            ReadNavRecordAnother = MyCommand.ExecuteReader
            MyCommand.Dispose()
        Catch MySQLException As SqlException
            'MsgBox(MySQLException.ToString)
            Return Nothing
        Catch MyException As Exception
            'MsgBox(MyException.ToString)
            Return Nothing
        End Try
    End Function
    Public Function ReadNavRecordAnother2(ByVal mySQL As String) As SqlDataReader
        Dim MyCommand As New SqlCommand
        If bNavConn = False Then ConnectNavDB()
        Try
            'ConnectNavDB()
            MyCommand.Connection = MyNavConnection2
            MyCommand.CommandText = mySQL
            MyCommand.CommandTimeout = 0
            ReadNavRecordAnother2 = MyCommand.ExecuteReader
            MyCommand.Dispose()
        Catch MySQLException As SqlException
            'MsgBox(MySQLException.ToString)
            Return Nothing
        Catch MyException As Exception
            'MsgBox(MyException.ToString)
            Return Nothing
        End Try
    End Function
End Module
