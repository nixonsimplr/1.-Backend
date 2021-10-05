Imports System
Imports System.Resources
Imports System.Threading
Imports System.Reflection
Imports System.Globalization
Imports DataInterface.IbizDO
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Public Class ServiceMaster
    Implements ISalesBase



    Private objDO As New DataInterface.IbizDO
    Private Sub GoodsChkIn_btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        objDO.ExecuteSQL("Delete from ServiceMaster")
        Dim i As Integer
        Try
            For i = 0 To dgvExchange.Rows.Count - 2
                objDO.ExecuteSQL("Insert into ServiceMaster(ServiceID, ServiceType) values (" & objDO.SafeSQL(dgvExchange(0, i).Value.ToString) & "," & objDO.SafeSQL(dgvExchange(1, i).Value.ToString) & ")")
            Next
            MsgBox("Service Master Updated")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    Private Sub dgvExchange_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvExchange.CellValueChanged
        If e.RowIndex < 0 Then Exit Sub
        If e.ColumnIndex < 0 Then Exit Sub
        'If e.ColumnIndex = 2 Then
        '    dgvExchange.Item(3, e.RowIndex).Value = "0"
        'End If
    End Sub

    Private Sub ServiceMaster_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim rs As SqlDataReader
        rs = objDO.ReadRecord("SELECT * FROM ServiceMaster")
        dgvExchange.Rows.Clear()
        While rs.Read
            Dim row As String() = New String() _
            {rs("ServiceID").ToString, rs("ServiceType").ToString}
            dgvExchange.Rows.Add(row)
        End While
        rs.Close()
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