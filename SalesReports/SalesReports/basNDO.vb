Imports System.Data.odbc
Module basNDO
    Public MyNavConnection As OdbcConnection
    Public bNavConn As Boolean = False

    Public Sub ConnectNavDB()
        Dim MyNavConString As String = My.Settings.NavConnectionString
        'MsgBox(MyNavConString)
        MyNavConnection = New OdbcConnection(MyNavConString)
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
        Dim MyCommand As New OdbcCommand
        If bNavConn = False Then ConnectDB()
        MyCommand.Connection = MyNavConnection
        MyCommand.CommandText = mySQL
        MyCommand.CommandTimeout = 0
        MyCommand.ExecuteNonQuery()
    End Sub

    Public Function ReadNavRecord(ByVal mySQL As String) As OdbcDataReader
        Dim MyCommand As New OdbcCommand
        If bNavConn = False Then ConnectDB()
        Try
            MyCommand.Connection = MyNavConnection
            MyCommand.CommandText = mySQL
            MyCommand.CommandTimeout = 0
            ReadNavRecord = MyCommand.ExecuteReader
        Catch MySqlException As OdbcException
            MsgBox(MySqlException.ToString)
            Return Nothing
        Catch MyException As Exception
            MsgBox(MyException.ToString)
            Return Nothing
        End Try
    End Function
End Module
