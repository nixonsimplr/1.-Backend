Imports System.Data.Odbc

Public Class NaviDO
    Private MyNavConnection As OdbcConnection
    Public IsNavisionConnected As Boolean = False
    Public ConnectionString As String

    Public Sub ConnectNavDB()
        Dim MyConString As String = My.Settings.NavConnectionString
        MyNavConnection = New OdbcConnection(MyConString)
        MyNavConnection.Open()
        IsNavisionConnected = True
    End Sub

    Public Sub DisconnectNavDB()
        MyNavConnection.Close()
    End Sub

    Public Sub ExecuteNavSQL(ByVal mySQL As String)
        Dim MyCommand As New OdbcCommand
        If IsNavisionConnected = False Then ConnectNavDB()
        MyCommand.Connection = MyNavConnection
        MyCommand.CommandText = mySQL
        MyCommand.ExecuteNonQuery()
    End Sub

    Public Function SafeSQLNav(ByVal sSQL As String) As String
        Dim iPos As Integer
        sSQL = Trim(sSQL)
        iPos = InStr(sSQL, "'")
        Do Until iPos = 0
            sSQL = Left(sSQL, iPos) & "'" & Mid(sSQL, iPos + 1)
            iPos = InStr(iPos + 2, sSQL, "'")
        Loop
        SafeSQLNav = "'" & sSQL & "'"
    End Function
    Public Function ReadNavRecord(ByVal mySQL As String) As OdbcDataReader
        Dim MyCommand As New OdbcCommand
        If IsNavisionConnected = False Then ConnectNavDB()
        Try
            MyCommand.Connection = MyNavConnection
            MyCommand.CommandText = mySQL
            ReadNavRecord = MyCommand.ExecuteReader
        Catch MyOdbcException As OdbcException
            MsgBox(MyOdbcException.ToString)
            Return Nothing
        Catch MyException As Exception
            MsgBox(MyException.ToString)
            Return Nothing
        End Try
    End Function
End Class
