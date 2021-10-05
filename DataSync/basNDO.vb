Imports System.Data.SqlClient
Module basNDO
    Public MyConnection As SqlConnection

    Public Sub ConnectDB(ByVal MyConString As String)
        MyConnection = New SqlConnection(MyConString)
        MyConnection.Open()
    End Sub

    Public Sub DisconnectDB()
        MyConnection.Close()
    End Sub

    Public Sub ExecuteSQL(ByVal mySQL As String)
        Dim MyCommand As New SqlCommand
        MyCommand.Connection = MyConnection
        MyCommand.CommandText = mySQL
        MyCommand.ExecuteNonQuery()
    End Sub

    Public Function ReadRecord(ByVal mySQL As String) As SqlDataReader
        Dim MyCommand As New SqlCommand
        Try
            MyCommand.Connection = MyConnection
            MyCommand.CommandText = mySQL
            ReadRecord = MyCommand.ExecuteReader
        Catch MyOdbcException As SqlException
            MsgBox(MyOdbcException.ToString)
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

    Public Function SafeSQL(ByVal sSQL As String) As String
        Dim iPos As Integer
        iPos = InStr(sSQL, "'")
        Do Until iPos = 0
            sSQL = Left(sSQL, iPos) & "'" & Mid(sSQL, iPos + 1)
            iPos = InStr(iPos + 2, sSQL, "'")
        Loop
        SafeSQL = "'" & sSQL & "'"
    End Function

    Public Function GetUpdateCommand(ByVal ds As DataSet, ByVal row As DataRow) As String
        Dim CommandSQL As String = ""
        Dim sField As String
        For Each column As DataColumn In ds.Tables(0).Columns
            If CommandSQL <> "" Then CommandSQL += ", "
            sField = column.ColumnName.ToString
            If sField.ToLower = "type" Then sField = "[" & sField & "]"
            CommandSQL = CommandSQL & sField & "="
            If row(column) Is DBNull.Value Then
                CommandSQL = CommandSQL & "NULL"
            Else
                If column.DataType.ToString = "System.String" Then
                    CommandSQL = CommandSQL & SafeSQL(row(column).ToString)
                ElseIf column.DataType.ToString = "System.DateTime" Then
                    CommandSQL = CommandSQL & SafeSQL(Format(CDate(row(column).ToString), "yyyyMMdd hh:mm:ss"))
                Else
                    CommandSQL = CommandSQL & row(column).ToString
                End If
            End If
        Next column
        Return CommandSQL
    End Function

    Public Function GetInsertCommand(ByVal ds As DataSet, ByVal row As DataRow) As String
        Dim CommandSQL As String = ""
        Dim ValueSQL As String = ""
        Dim sField As String
        For Each column As DataColumn In ds.Tables(0).Columns
            If CommandSQL <> "" Then CommandSQL += ", "
            sField = column.ColumnName.ToString
            If sField.ToLower = "type" Then sField = "[" & sField & "]"
            CommandSQL = CommandSQL & sField
            If ValueSQL <> "" Then ValueSQL += ", "
            If row(column) Is DBNull.Value Then
                ValueSQL = ValueSQL & "NULL"
            Else
                If column.DataType.ToString = "System.String" Then
                    ValueSQL = ValueSQL & SafeSQL(row(column).ToString)
                ElseIf column.DataType.ToString = "System.DateTime" Then
                    ValueSQL = ValueSQL & SafeSQL(Format(CDate(row(column).ToString), "yyyyMMdd hh:mm:ss"))
                Else
                    ValueSQL = ValueSQL & row(column).ToString
                End If
            End If
        Next column
        Return "(" & CommandSQL & ") Values (" & ValueSQL & ")"
    End Function

End Module
