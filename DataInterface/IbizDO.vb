Imports System.Data.SqlClient

Public Class IbizDO
    Public MyConnection As New SqlConnection
    Public MyAConnection As New SqlConnection
    Public ConnectionString As String
    Public IsConnected As Boolean = False
    Public bAConn As Boolean = False

    Public Sub ConnectDB()
        Dim MyConString As String = My.Settings.ConnectionString ' "Data Source=IBIZCS-JAGADISH\SQLEXPRESS;Initial Catalog=Sales;Integrated Security=True" 'My.Settings.ConnectionString
        MyConnection.ConnectionString = MyConString
        MyConnection.Open()
        IsConnected = True
    End Sub

    Public Sub ConnectAnotherDB()
        Dim MyAConString As String = My.Settings.ConnectionString
        MyAConnection = New SqlConnection(My.Settings.ConnectionString)
        MyAConnection.Open()
        bAConn = True
    End Sub

    Public Sub DisconnectDB()
        IsConnected = False
        MyConnection.Close()
    End Sub
    Public Sub DisconnectAnotherDB()
        IsConnected = False
        MyAConnection.Close()
    End Sub

    Public Sub ExecuteSQL(ByVal mySQL As String)
        Dim MyCommand As New SqlCommand
        If IsConnected = False Then ConnectDB()
        IsConnected = True
        MyCommand.Connection = MyConnection
        MyCommand.CommandText = mySQL
        MyCommand.CommandTimeout = 0
        MyCommand.ExecuteNonQuery()
    End Sub

    Public Sub ExecuteSQLAnother(ByVal mySQL As String)
        Dim MyCommand As New SqlCommand
        MyCommand.Connection = New SqlConnection
        If bAConn = False Then ConnectAnotherDB()
        MyCommand.Connection.ConnectionString = My.Settings.ConnectionString ' "Data Source=.\SQLEXPRESS;Initial Catalog=Sales;Integrated Security=True" 'My.Settings.ConnectionString  '"Dsn=IbizMobile;dbq=C:\My Projects\Chee Seng\Database\Sales.mdb;driverid=25;fil=MS Access;maxbuffersize=2048;pagetimeout=5"
        MyCommand.Connection.Open()
        MyCommand.CommandText = mySQL
        MyCommand.CommandTimeout = 0
        MyCommand.ExecuteNonQuery()
        MyCommand.Connection.Close()
    End Sub

    Public Function ReadRecord(ByVal mySQL As String) As SqlDataReader
        Dim MyCommand As New SqlCommand
        If IsConnected = False Then ConnectDB()
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

    Public Function GetDataAdapter(ByVal mySQL As String) As SqlDataAdapter
        Dim oda As New SqlDataAdapter(mySQL, MyConnection)
        Return oda
    End Function

    Public Function GetDataSet(ByVal mySQL As String) As DataSet
        Dim da As New SqlDataAdapter(mySQL, MyConnection)
        Dim ds As New DataSet
        da.Fill(ds.Tables(0))
        Return ds
    End Function

    Public Function GetCount(ByVal mySQL As String) As Integer
        Dim MyCommand As New SqlCommand
        If IsConnected = False Then ConnectDB()
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
        Dim iPos As Integer
        sSQL = Trim(sSQL)
        iPos = InStr(sSQL, "'")
        Do Until iPos = 0
            sSQL = Left(sSQL, iPos) & "'" & Mid(sSQL, iPos + 1)
            iPos = InStr(iPos + 2, sSQL, "'")
        Loop
        SafeSQL = "'" & sSQL & "'"
    End Function

    Public Function ReadRecordAnother(ByVal mySQL As String) As SqlDataReader
        Dim MyCommand As New SqlCommand
        If bAConn = False Then ConnectAnotherDB()
        Try
            MyCommand.Connection = MyAConnection
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
End Class
