Imports System.Data.OleDb
Public Class OledbDO
    Private MyOleDBConnection As OleDbConnection
    Public IsOleDBConnected As Boolean = False
    Public ConnectionString As String
    Public Sub ConnectOleDBDB()
        Dim MyConString As String = My.Settings.OledbConnectionString
        'Dim MyConString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Desktop\UniCurd\transfer.mdb;"
        MyOleDBConnection = New OleDbConnection(MyConString)
        MyOleDBConnection.Open()
        IsOleDBConnected = True
    End Sub

    Public Sub DisconnectOleDBDB()
        MyOleDBConnection.Close()
    End Sub

    Public Sub ExecuteOleDBSQL(ByVal mySQL As String)
        Dim MyCommand As New OleDbCommand
        MyCommand.Connection = MyOleDBConnection
        MyCommand.CommandText = mySQL
        MyCommand.ExecuteNonQuery()
    End Sub

    Public Function ReadOleDBRecord(ByVal mySQL As String) As OleDbDataReader
        Dim MyCommand As New OleDbCommand
        If IsOleDBConnected = False Then ConnectOleDBDB()
        Try
            MyCommand.Connection = MyOleDBConnection
            MyCommand.CommandText = mySQL
            ReadOleDBRecord = MyCommand.ExecuteReader
        Catch MyOledbException As OleDbException
            MsgBox(MyOledbException.ToString)
            Return Nothing
        Catch MyException As Exception
            MsgBox(MyException.ToString)
            Return Nothing
        End Try
    End Function
End Class
