Imports System.Data.SqlClient
Module basDO
    Public MyConnection As SqlConnection
    Public MyAnotherConnection As SqlConnection
    Public bConn As Boolean = False
    Public bAConn As Boolean = False
    
    Public Sub ConnectDB()
        Dim MyConString As String = My.Settings.ConnectionString '"Data Source=IBIZ-JAGADISH\SQLEXPRESS;Initial Catalog=Sales;Integrated Security=True" 'My.Settings.ConnectionString   '"Dsn=IbizMobile;dbq=C:\My Projects\Chee Seng\Database\Sales.mdb;driverid=25;fil=MS Access;maxbuffersize=2048;pagetimeout=5"
        MyConnection = New SqlConnection(MyConString)
        MyConnection.Open()
        bConn = True
    End Sub

    Public Sub DisconnectDB()
        If bConn = True Then MyConnection.Close()
        bConn = False
    End Sub

    Public Sub ExecuteSQL(ByVal mySQL As String)
        Try
            Dim MyCommand As New SqlCommand
            If bConn = False Then ConnectDB()
            MyCommand.Connection = MyConnection
            MyCommand.CommandText = mySQL
            MyCommand.CommandTimeout = 0
            MyCommand.ExecuteNonQuery()
            MyCommand.Dispose()
        Catch ex As Exception
            MyConnection.Close()
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'ExecuteSQL'," & SafeSQL("") & "," & SafeSQL(ex.Message) & ")")
        End Try
    End Sub

    Public Function ReadRecord(ByVal mySQL As String) As SqlDataReader
        Dim MyCommand As New SqlCommand
        If bConn = False Then ConnectDB()
        Try
            MyCommand.Connection = MyConnection
            MyCommand.CommandText = mySQL
            MyCommand.CommandTimeout = 0
            ReadRecord = MyCommand.ExecuteReader
            MyCommand.Dispose()
        Catch MySqlException As SqlException
            MyConnection.Close()
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'ReadRecord'," & SafeSQL("") & "," & SafeSQL(MySqlException.Message) & ")")
            Return Nothing
        Catch MyException As Exception
            MyConnection.Close()
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'ReadRecord'," & SafeSQL("") & "," & SafeSQL(MyException.Message) & ")")
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
            MyConnection.Close()
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'GetCount'," & SafeSQL("") & "," & SafeSQL(MySqlException.Message) & ")")

        Catch MyException As Exception
            MyConnection.Close()
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'GetCount'," & SafeSQL("") & "," & SafeSQL(MyException.Message) & ")")

        End Try

    End Function

    Public Function SafeSQL(ByVal sSQL As String) As String
        Dim ii As Integer
        ii = InStr(1, sSQL, "'")
        Do While ii <> 0
            sSQL = Left(sSQL, ii) & "'" & Mid(sSQL, ii + 1)
            ii = InStr(ii + 2, sSQL, "'")
        Loop
        Return "'" & sSQL & "'"
    End Function

    Public Sub ConnectAnotherDB()
        Dim MyConString As String = My.Settings.ConnectionString '"Data Source=IBIZ-JAGADISH\SQLEXPRESS;Initial Catalog=Sales;Integrated Security=True" 'My.Settings.ConnectionString   '"Dsn=IbizMobile;dbq=C:\My Projects\Chee Seng\Database\Sales.mdb;driverid=25;fil=MS Access;maxbuffersize=2048;pagetimeout=5"
        MyAnotherConnection = New SqlConnection(MyConString)
        MyAnotherConnection.Open()
        bAConn = True
    End Sub

    Public Sub DisconnectAnotherDB()
        If bAConn = True Then
            MyAnotherConnection.Close()
            bAConn = False
        End If
    End Sub

    Public Function ReadRecordAnother(ByVal mySQL As String) As SqlDataReader
        Dim MyCommand As New SqlCommand

        Try
            If bAConn = False Then ConnectAnotherDB()
            If MyAnotherConnection.State = ConnectionState.Closed Then
                MyAnotherConnection.Open()
            End If
            MyCommand.Connection = MyAnotherConnection
            MyCommand.CommandText = mySQL
            MyCommand.CommandTimeout = 0
            ReadRecordAnother = MyCommand.ExecuteReader
            MyCommand.Dispose()
        Catch MySqlException As SqlException
            MyAnotherConnection.Close()
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'ReadRecordAnother'," & SafeSQL("") & "," & SafeSQL(MySqlException.Message) & ")")
            Return Nothing
        Catch MyException As Exception
            MyAnotherConnection.Close()
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'ReadRecordAnother'," & SafeSQL("") & "," & SafeSQL(MyException.Message) & ")")
            Return Nothing
        End Try

    End Function

    Public Sub ExecuteSQLAnother(ByVal mySQL As String)
        Try
            If bAConn = False Then ConnectAnotherDB()
            Dim MyCommand As New SqlCommand
            If MyAnotherConnection.State = ConnectionState.Closed Then
                MyAnotherConnection.Open()
            End If
            MyCommand.Connection = MyAnotherConnection
            MyCommand.CommandText = mySQL
            MyCommand.CommandTimeout = 0
            MyCommand.ExecuteNonQuery()
            MyCommand.Connection.Close()
            MyCommand.Dispose()
        Catch ex As Exception
            MyAnotherConnection.Close()
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'ExecuteSQLAnother'," & SafeSQL("") & "," & SafeSQL(ex.Message) & ")")
        End Try
    End Sub
End Module
