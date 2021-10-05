Imports System.Data.Odbc
Public Class ScalaDO
    Private MyScalaConnection As OdbcConnection
    Public IsScalaConnected As Boolean = False
    Public ConnectionString As String

    Public Sub ConnectScalaDB()
        Dim MyConString As String = My.Settings.NavConnectionString
        MyScalaConnection = New OdbcConnection(MyConString)
        MyScalaConnection.Open()
        IsScalaConnected = True
    End Sub

    Public Sub DisconnectScalaDB()
        MyScalaConnection.Close()
    End Sub

    Public Sub ExecuteScalaSQL(ByVal mySQL As String)
        Dim MyCommand As New OdbcCommand
        If IsScalaConnected = False Then ConnectScalaDB()
        MyCommand.Connection = MyScalaConnection
        MyCommand.CommandText = mySQL
        MyCommand.ExecuteNonQuery()
    End Sub

    Public Function ReadScalaRecord(ByVal mySQL As String) As OdbcDataReader
        Dim MyCommand As New OdbcCommand
        If IsScalaConnected = False Then ConnectScalaDB()
        Try
            MyCommand.Connection = MyScalaConnection
            MyCommand.CommandText = mySQL
            ReadScalaRecord = MyCommand.ExecuteReader
        Catch MyOdbcException As OdbcException
            MsgBox(MyOdbcException.ToString)
            Return Nothing
        Catch MyException As Exception
            MsgBox(MyException.ToString)
            Return Nothing
        End Try

    End Function
    Public Sub DisconnectDB()
        MyConnection.Close()
    End Sub
    Public Sub DisconnectAnotherDB()
        MyAConnection.Close()
    End Sub
End Class
