Imports System.Data
Imports System.Data.SqlClient
Module basDo
    Public MyConnection As SqlConnection
    Public MyAConnection As SqlConnection
    Public bConn As Boolean = False
    Public bAConn As Boolean = False
    Public Sub ConnectDB()
        Dim MyConString As String = My.Settings.ConnectionString
        MyConnection = New SqlConnection(My.Settings.ConnectionString)
        MyConnection.Open()
        bConn = True
    End Sub
    Public Sub ConnectAnotherDB()
        Dim MyAConString As String = My.Settings.ConnectionString
        MyAConnection = New SqlConnection(My.Settings.ConnectionString)
        MyAConnection.Open()
        bAConn = True
    End Sub
    Public Sub DisconnectDB()
        MyConnection.Close()
        bConn = False
    End Sub
    Public Sub DisconnectAnotherDB()
        MyAConnection.Close()
        bAConn = False
    End Sub
    Public Function ReturnDataSet(ByVal mySql As String) As DataSet
        Dim Connection As New SqlConnection(My.Settings.ConnectionString)
        Dim DA As New SqlDataAdapter(mySql, Connection)
        Dim DS As New DataSet
        DA.Fill(DS)
        Return DS
    End Function
   
    Public Sub ExecuteSQL(ByVal mySQL As String)
        Dim MyCommand As New SqlCommand
        If bConn = False Then ConnectDB()
        MyCommand.Connection = MyConnection
        MyCommand.CommandText = mySQL
        MyCommand.CommandTimeout = 0
        MyCommand.ExecuteNonQuery()
    End Sub
    Public Sub ExecuteAnotherSQL(ByVal mySQL As String)
        Dim MyACommand As New SqlCommand
        If bAConn = False Then ConnectAnotherDB()
        MyACommand.Connection = MyAConnection
        MyACommand.CommandText = mySQL
        MyACommand.CommandTimeout = 0
        MyACommand.ExecuteNonQuery()
    End Sub

    Public Function ReadRecord(ByVal mySQL As String) As SqlDataReader
        Dim MyCommand As New SqlCommand
        If bConn = False Then ConnectDB()
        Try
            MyCommand.Connection = MyConnection
            MyCommand.CommandText = mySQL
            MyCommand.CommandTimeout = 0
            ReadRecord = MyCommand.ExecuteReader
        Catch MySqlException As SqlException
            MsgBox(MySqlException.ToString)
            Return Nothing
        Catch MyException As Exception
            MsgBox(MyException.ToString)
            Return Nothing
        End Try
    End Function

    Public Function GetCount(ByVal mySQL As String) As Integer
        Dim MyCommand As New SqlCommand
        If bConn = False Then ConnectDB()
        Try
            MyCommand.Connection = MyConnection
            MyCommand.CommandText = mySQL
            MyCommand.CommandTimeout = 0
            GetCount = MyCommand.ExecuteNonQuery
            MyCommand.Cancel()
        Catch MySqlException As SqlException
            MsgBox(MySqlException.ToString)
        Catch MyException As Exception
            MsgBox(MyException.ToString)
        End Try

    End Function

    Public Function SafeSQL(ByVal sSQL As String) As String
        Dim iPos As Integer
        sSQL = Trim(sSQL)
        iPos = InStr(sSQL, "'")
        Do Until iPos = 0
            sSQL = Left(sSQL, iPos) & "'" & Mid(sSQL, iPos + 1)
            iPos = InStr(iPos + 2, sSQL, "'")
        Loop
        SafeSQL = "'" & sSQL & "'"
        'SafeSQL = "'" & sSQL & "'"
    End Function

    Public Function ReadRecordAnother(ByVal mySQL As String) As SqlDataReader
        Dim MyCommand As New SqlCommand
        If bAConn = False Then ConnectAnotherDB()
        Try
            MyCommand.Connection = MyAConnection
            MyCommand.CommandText = mySQL
            MyCommand.CommandTimeout = 0
            ReadRecordAnother = MyCommand.ExecuteReader
        Catch MySqlException As SqlException
            MsgBox(MySqlException.ToString)
            Return Nothing
        Catch MyException As Exception
            MsgBox(MyException.ToString)
            Return Nothing
        End Try
    End Function
    End Module
