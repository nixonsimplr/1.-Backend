Imports System
Imports System.Resources
Imports System.Threading
Imports System.Reflection
Imports System.Globalization

Imports SalesInterface.MobileSales

Public Class MDTList
    Implements ISalesBase



    Private Sub MDTList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'MDTSalesDataSet.MDT' table. You can move, or remove it, as needed.
        Me.MDTTableAdapter.Fill(Me.MDTSalesDataSet.MDT)

    End Sub

    Public Function GetListViewForm() As String Implements SalesInterface.MobileSales.ISalesBase.GetListViewForm
        Return ""
    End Function

    Public Sub ListViewClick() Implements SalesInterface.MobileSales.ISalesBase.ListViewClick

    End Sub

    Public Event LoadDataForm(ByVal Location As String, ByVal LoadType As String, ByVal ParentLoadType As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.LoadDataForm

    Public Event LoadForm(ByVal Location As String, ByVal LoadType As String) Implements SalesInterface.MobileSales.ISalesBase.LoadForm

    Public Sub MoveFirstClick() Implements SalesInterface.MobileSales.ISalesBase.MoveFirstClick
        MDTBindingSource.MoveFirst()
    End Sub

    Public Sub MoveLastClick() Implements SalesInterface.MobileSales.ISalesBase.MoveLastClick
        MDTBindingSource.MoveLast()
    End Sub

    Public Sub MoveNextClick() Implements SalesInterface.MobileSales.ISalesBase.MoveNextClick
        MDTBindingSource.MoveNext()
    End Sub

    Public Sub MovePositionClick(ByVal Position As Long) Implements SalesInterface.MobileSales.ISalesBase.MovePositionClick

    End Sub

    Public Sub MovePreviousClick() Implements SalesInterface.MobileSales.ISalesBase.MovePreviousClick
        MDTBindingSource.MovePrevious()
    End Sub

    Public Sub Print() Implements SalesInterface.MobileSales.ISalesBase.Print

    End Sub

    Public Sub Print(ByVal PageNo As Integer) Implements SalesInterface.MobileSales.ISalesBase.Print

    End Sub

    Public Event ResultData(ByVal ChildLoadType As String, ByVal Value As String) Implements SalesInterface.MobileSales.ISalesBase.ResultData

    Public Sub ReturnedData(ByVal Value As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.ReturnedData

    End Sub

    Public Sub ReturnSearch(ByVal SQL As String) Implements SalesInterface.MobileSales.ISalesBase.ReturnSearch
        Try
            MDTBindingSource.Filter = SQL
        Catch ex As Exception
        End Try
    End Sub

    Public Event SearchField(ByVal FormName As String, ByVal FieldName As String, ByVal FieldType As String, ByVal CurValue As String) Implements SalesInterface.MobileSales.ISalesBase.SearchField

    Public Sub SetCulture(ByVal CultureName As String) Implements SalesInterface.MobileSales.ISalesBase.SetCulture

    End Sub

    Private Sub dgvMDT_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvMDT.CellDoubleClick
        If e.ColumnIndex < 0 Then Exit Sub
        If e.RowIndex < 0 Then Exit Sub
        RaiseEvent ResultData("SalesPromo.MDTList", dgvMDT.Item(0, e.RowIndex).Value.ToString)
        Me.Close()
    End Sub

    Private Sub dgvMDT_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvMDT.CellEnter
        Dim dcl As DataGridViewColumn
        dcl = dgvMDT.Columns(e.ColumnIndex)
        Try
            Dim sField As String
            Dim sFieldType As String
            sField = dcl.DataPropertyName
            sFieldType = dcl.ValueType.Name
            If sField <> "" Then RaiseEvent SearchField(Me.ProductName & "." & Me.Name, sField, sFieldType, dgvMDT.Item(e.ColumnIndex, e.RowIndex).Value.ToString)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub dgvMDT_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvMDT.CellMouseClick
        If e.ColumnIndex < 0 Then Exit Sub
        If e.RowIndex < 0 Then Exit Sub
    End Sub

    Private Sub dgvMDT_CellMouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvMDT.CellMouseDoubleClick
        If e.ColumnIndex < 0 Then Exit Sub
        If e.RowIndex < 0 Then Exit Sub
    End Sub

    Private Sub dgvMDT_CellMouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvMDT.CellMouseDown
        If e.ColumnIndex < 0 Then Exit Sub
        If e.RowIndex < 0 Then Exit Sub
    End Sub
End Class