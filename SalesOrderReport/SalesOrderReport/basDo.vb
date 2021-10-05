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
    End Sub
    Public Sub DisconnectAnotherDB()
        MyAConnection.Close()
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
        MyCommand.ExecuteNonQuery()
    End Sub
    Public Sub ExecuteAnotherSQL(ByVal mySQL As String)
        Dim MyACommand As New SqlCommand
        If bAConn = False Then ConnectAnotherDB()
        MyACommand.Connection = MyAConnection
        MyACommand.CommandText = mySQL
        MyACommand.ExecuteNonQuery()
    End Sub

    Public Function ReadRecord(ByVal mySQL As String) As SqlDataReader
        Dim MyCommand As New SqlCommand
        If bConn = False Then ConnectDB()
        Try
            MyCommand.Connection = MyConnection
            MyCommand.CommandText = mySQL
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
            GetCount = MyCommand.ExecuteNonQuery
            MyCommand.Cancel()
        Catch MySqlException As SqlException
            MsgBox(MySqlException.ToString)
        Catch MyException As Exception
            MsgBox(MyException.ToString)
        End Try

    End Function

    Public Function SafeSQL(ByVal sSQL As String) As String
        SafeSQL = "'" & sSQL & "'"
    End Function

    Public Function ReadRecordAnother(ByVal mySQL As String) As SqlDataReader
        Dim MyCommand As New SqlCommand
        Try
            MyCommand.Connection = New SqlConnection
            MyCommand.Connection.ConnectionString = My.Settings.ConnectionString
            MyCommand.Connection.Open()
            MyCommand.CommandText = mySQL
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
