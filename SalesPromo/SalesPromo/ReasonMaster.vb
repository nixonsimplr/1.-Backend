Imports System
Imports System.Resources
Imports System.Threading
Imports System.Reflection
Imports System.Globalization
Imports DataInterface.IbizDO
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Public Class ReasonMaster
    Implements ISalesBase



    Private objDO As New DataInterface.IbizDO
    Dim bLoad As Boolean = True
    Private Sub ReasonMaster_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim rs As SqlDataReader
        rs = objDO.ReadRecord("SELECT * FROM Reason")
        dgvExchange.Rows.Clear()
        Dim cnt As Integer = 0
        While rs.Read
            Dim row As String() = New String() _
            {rs("Code").ToString, rs("Description").ToString, _
             ""}
            dgvExchange.Rows.Add(row)
            Dim cb As New DataGridViewTextBoxCell
            dgvExchange.Item(2, cnt) = cb
            dgvExchange.Item(2, cnt).Value = rs("ReasonType").ToString
            cnt = cnt + 1
        End While
        rs.Close()
        bLoad = False
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        objDO.ExecuteSQL("Delete from Reason")
        Dim i As Integer
        Try
            For i = 0 To dgvExchange.Rows.Count - 2
                objDO.ExecuteSQL("Insert into Reason(Code, Description, ReasonType) values (" & objDO.SafeSQL(dgvExchange(0, i).Value.ToString) & "," & objDO.SafeSQL(dgvExchange(1, i).Value.ToString) & "," & objDO.SafeSQL(dgvExchange(2, i).Value.ToString) & ")")
            Next
            MsgBox("Reason Master Updated")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub dgvExchange_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvExchange.CellValueChanged
        If e.RowIndex < 0 Then Exit Sub
        If e.ColumnIndex < 0 Then Exit Sub
        If bLoad = True Then Exit Sub
        If e.ColumnIndex = 1 Then
            Dim dtr As SqlDataReader
            Dim cb As New DataGridViewComboBoxCell
            cb = dgvExchange.Item(2, e.RowIndex)
            cb.Items.Clear()
            Dim bCur As Boolean = False
            dtr = objDO.ReadRecord("Select ServiceID from ServiceMaster")
            cb.Items.Add("FOC")
            cb.Items.Add("EX")
            cb.Items.Add("RTN")
            While dtr.Read
                cb.Items.Add(dtr("ServiceID"))
                bCur = True
            End While
            dtr.Close()
        End If
    End Sub

    Public Function GetListViewForm() As String Implements SalesInterface.MobileSales.ISalesBase.GetListViewForm
        Return ""
    End Function

    Public Sub ListViewClick() Implements SalesInterface.MobileSales.ISalesBase.ListViewClick

    End Sub

    Public Event LoadDataForm(ByVal Location As String, ByVal LoadType As String, ByVal ParentLoadType As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.LoadDataForm

    Public Event LoadForm(ByVal Location As String, ByVal LoadType As String) Implements SalesInterface.MobileSales.ISalesBase.LoadForm

    Public Sub MoveFirstClick() Implements SalesInterface.MobileSales.ISalesBase.MoveFirstClick

    End Sub

    Public Sub MoveLastClick() Implements SalesInterface.MobileSales.ISalesBase.MoveLastClick

    End Sub

    Public Sub MoveNextClick() Implements SalesInterface.MobileSales.ISalesBase.MoveNextClick

    End Sub

    Public Sub MovePositionClick(ByVal Position As Long) Implements SalesInterface.MobileSales.ISalesBase.MovePositionClick

    End Sub

    Public Sub MovePreviousClick() Implements SalesInterface.MobileSales.ISalesBase.MovePreviousClick

    End Sub

    Public Sub Print() Implements SalesInterface.MobileSales.ISalesBase.Print

    End Sub

    Public Sub Print(ByVal PageNo As Integer) Implements SalesInterface.MobileSales.ISalesBase.Print

    End Sub

    Public Event ResultData(ByVal ChildLoadType As String, ByVal Value As String) Implements SalesInterface.MobileSales.ISalesBase.ResultData

    Public Sub ReturnedData(ByVal Value As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.ReturnedData

    End Sub

    Public Sub ReturnSearch(ByVal SQL As String) Implements SalesInterface.MobileSales.ISalesBase.ReturnSearch

    End Sub

    Public Event SearchField(ByVal FormName As String, ByVal FieldName As String, ByVal FieldType As String, ByVal CurValue As String) Implements SalesInterface.MobileSales.ISalesBase.SearchField

    Public Sub SetCulture(ByVal CultureName As String) Implements SalesInterface.MobileSales.ISalesBase.SetCulture

    End Sub

 End Class