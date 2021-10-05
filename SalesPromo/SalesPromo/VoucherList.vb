Imports System
Imports System.Resources
Imports System.Threading
Imports System.Reflection
Imports System.Globalization
Imports SalesInterface.MobileSales
Imports System.Data.SqlClient
Public Class VoucherList
    Implements ISalesBase

    Private objDO As New DataInterface.IbizDO
    Private myAdapter As SqlDataAdapter
    Private myDataSet As DataSet
    Private myDataView As DataView
    Private Sub VoucherList_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        objDO.DisconnectDB()
    End Sub
    Private Sub VoucherList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'VoucherDataSet.Voucher' table. You can move, or remove it, as needed.
        '  Me.VoucherTableAdapter.Fill(Me.VoucherDataSet.Voucher)
        objDO.ConnectionString = My.Settings.ConnectionString
        objDO.ConnectDB()
        LoadData("SELECT * FROM VoucherHeader order by VoucherNo")
    End Sub
    Private Sub LoadData(ByVal sSQL As String)
        Try
            myAdapter = objDO.GetDataAdapter(sSQL)
            myDataSet = New DataSet
            myAdapter.Fill(myDataSet, "VoucherHeader")
            myDataView = New DataView(myDataSet.Tables("VoucherHeader"))
            VoucherBindingSource.DataSource = myDataView
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Function GetListViewForm() As String Implements SalesInterface.MobileSales.ISalesBase.GetListViewForm
        Return ""
    End Function

    Public Sub ListViewClick() Implements SalesInterface.MobileSales.ISalesBase.ListViewClick

    End Sub

    Public Event LoadDataForm(ByVal Location As String, ByVal LoadType As String, ByVal ParentLoadType As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.LoadDataForm

    Public Event LoadForm(ByVal Location As String, ByVal LoadType As String) Implements SalesInterface.MobileSales.ISalesBase.LoadForm

    Public Sub MoveFirstClick() Implements SalesInterface.MobileSales.ISalesBase.MoveFirstClick
        VoucherBindingSource.MoveFirst()
    End Sub

    Public Sub MoveLastClick() Implements SalesInterface.MobileSales.ISalesBase.MoveLastClick
        VoucherBindingSource.MoveLast()
    End Sub

    Public Sub MoveNextClick() Implements SalesInterface.MobileSales.ISalesBase.MoveNextClick
        VoucherBindingSource.MoveNext()
    End Sub

    Public Sub MovePositionClick(ByVal Position As Long) Implements SalesInterface.MobileSales.ISalesBase.MovePositionClick

    End Sub

    Public Sub MovePreviousClick() Implements SalesInterface.MobileSales.ISalesBase.MovePreviousClick
        VoucherBindingSource.MovePrevious()
    End Sub

    Public Sub Print() Implements SalesInterface.MobileSales.ISalesBase.Print

    End Sub

    Public Sub Print(ByVal PageNo As Integer) Implements SalesInterface.MobileSales.ISalesBase.Print

    End Sub

    Public Event ResultData(ByVal ChildLoadType As String, ByVal Value As String) Implements SalesInterface.MobileSales.ISalesBase.ResultData

    Public Sub ReturnedData(ByVal Value As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.ReturnedData

    End Sub

    Public Sub ReturnSearch(ByVal SQL As String) Implements SalesInterface.MobileSales.ISalesBase.ReturnSearch
        If SQL = "" Then
            LoadData("SELECT * FROM VoucherHeader order by VoucherNo")
        Else
            LoadData("SELECT * FROM VoucherHeader where " & SQL & " order by VoucherNo")
        End If
    End Sub

    Public Event SearchField(ByVal FormName As String, ByVal FieldName As String, ByVal FieldType As String, ByVal CurValue As String) Implements SalesInterface.MobileSales.ISalesBase.SearchField

    Public Sub SetCulture(ByVal CultureName As String) Implements SalesInterface.MobileSales.ISalesBase.SetCulture
    End Sub

    Private Sub dgvVouList_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvVouList.CellDoubleClick
        RaiseEvent ResultData("SalesPromo.VoucherList", dgvVouList.Item(0, e.RowIndex).Value.ToString)
        Me.Close()
    End Sub

    Private Sub dgvVouList_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvVouList.CellEnter
        Dim dcl As DataGridViewColumn
        dcl = dgvVouList.Columns(e.ColumnIndex)
        Try
            Dim sField As String
            Dim sFieldType As String
            sField = dcl.DataPropertyName
            sFieldType = dcl.ValueType.Name
            If sField <> "" Then RaiseEvent SearchField(Me.ProductName & "." & Me.Name, sField, sFieldType, dgvVouList.Item(e.ColumnIndex, e.RowIndex).Value.ToString)
        Catch ex As Exception
        End Try
    End Sub

   End Class