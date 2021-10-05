Imports System
Imports System.Resources
Imports System.Threading
Imports System.Reflection
Imports System.Globalization

Imports SalesInterface.MobileSales
Imports System.Data.SqlClient

Public Class POList
    Implements ISalesBase


#Region "Localization Variants"
    Private rMgr As ResourceManager
    Private cInfo As CultureInfo
    Private sLang As String
#End Region

    Private objDO As New DataInterface.IbizDO
    Private myAdapter As SqlDataAdapter
    Private myDataSet As DataSet
    Private myDataView As DataView
    Private bFlag As Boolean

    Private Sub DataList_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        objDO.DisconnectDB()
    End Sub

    Private Sub DataList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'PODataSet.PO' table. You can move, or remove it, as needed.
        'Me.POTableAdapter.Fill(Me.PODataSet.PO)
        'TODO: This line of code loads data into the 'InvoiceDataSet.Invoice' table. You can move, or remove it, as needed.
        'Me.InvoiceTableAdapter.Fill(Me.InvoiceDataSet.Invoice)
        objDO.ConnectionString = My.Settings.ConnectionString
        objDO.ConnectDB()
        bFlag = False
        LoadData("SELECT * FROM PO")
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv.CellClick
        bFlag = False
    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv.CellDoubleClick
        RaiseEvent ResultData("SalesPromo.POList", dgv.Item(0, e.RowIndex).Value.ToString)
        Me.Close()
    End Sub

    Public Function GetListViewForm() As String Implements SalesInterface.MobileSales.ISalesBase.GetListViewForm
        Return String.Empty
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

    '=====================================================================================
    'Search Function
    Private Sub LoadData(ByVal sSQL As String)
        '  myAdapter = objDO.GetDataAdapter("SELECT CustNo AS [No], CustName AS Name, ChineseName AS [Chinese Name], Address, Address2 AS [Address 2], Address3 AS [Address 3], Address4 AS [Address 4], PostCode AS [Post Code], City, CountryCode AS [Country Name], CountryCode AS [Country Code], Phone, ContactPerson AS [Contact Person], Balance, CreditLimit AS [Credit Limit], ProvisionalBalance AS [Provisional Balance], ZoneCode AS [Zone Description], FaxNo AS [Fax No], Email, Website, ICPartner AS [IC Partner], PriceGroup AS [Price Group Description], PaymentTerms AS [Payment Term Description], PaymentMethod AS [Payment Method Description], ShipmentMethod AS [Shipment Method Description], SalesAgent AS [Sales Agent], ShipAgent AS [Shipping Agent], [Bill-toNo] AS [Bill-to No], InvDisGroup AS [Invoice Discount Group], Location AS [Location Name], CurrencyCode AS [Currency Code] FROM Customer where Active=1")
        Try

            myAdapter = objDO.GetDataAdapter(sSQL)
            myDataSet = New DataSet
            myAdapter.Fill(myDataSet, "PO")
            myDataView = New DataView(myDataSet.Tables("PO"))
            POBindingSource.DataSource = myDataView
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub ReturnSearch(ByVal SQL As String) Implements SalesInterface.MobileSales.ISalesBase.ReturnSearch
        bFlag = True
        If SQL = "" Then
            LoadData("SELECT * FROM PO")
        Else
            LoadData("SELECT * FROM PO where " & SQL)
        End If
    End Sub

    Public Event SearchField(ByVal FormName As String, ByVal FieldName As String, ByVal FieldType As String, ByVal CurValue As String) Implements SalesInterface.MobileSales.ISalesBase.SearchField

    Private Sub DataGridView1_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv.CellEnter
        If bFlag = True Then Exit Sub
        Dim dcl As DataGridViewColumn
        dcl = dgv.Columns(e.ColumnIndex)
        Try
            Dim sField As String
            Dim sFieldType As String
            sField = dcl.DataPropertyName
            sFieldType = dcl.ValueType.Name
            If sField <> "" Then RaiseEvent SearchField(Me.ProductName & "." & Me.Name, sField, sFieldType, dgv.Item(e.ColumnIndex, e.RowIndex).Value.ToString)
        Catch ex As Exception

        End Try
    End Sub

    'Localization
    Private Sub Localization()
        Try
            cInfo = New CultureInfo(sLang)
            Thread.CurrentThread.CurrentCulture = cInfo
            Thread.CurrentThread.CurrentUICulture = cInfo
            rMgr = New ResourceManager("Invoice.Invoice", [Assembly].GetExecutingAssembly())

            Me.Text = rMgr.GetString("DataList")
            dgv.Columns("InvNoDataGridViewTextBoxColumn").HeaderText = rMgr.GetString("Col_DataLst_InvoiceNo")
            dgv.Columns("InvDtDataGridViewTextBoxColumn").HeaderText = rMgr.GetString("Col_DataLst_Date")
            dgv.Columns("OrdNoDataGridViewTextBoxColumn").HeaderText = rMgr.GetString("Col_DataLst_OrderNo")
            dgv.Columns("CustIdDataGridViewTextBoxColumn").HeaderText = rMgr.GetString("Col_DataLst_CustNo")
            dgv.Columns("AgentIDDataGridViewTextBoxColumn").HeaderText = rMgr.GetString("Col_DataLst_AgentCode")
            dgv.Columns("PayTermsDataGridViewTextBoxColumn").HeaderText = rMgr.GetString("Col_DataLst_PayTerms")
            dgv.Columns("CurCodeDataGridViewTextBoxColumn").HeaderText = rMgr.GetString("Col_DataLst_Currny")
            dgv.Columns("CurExRateDataGridViewTextBoxColumn").HeaderText = rMgr.GetString("Col_DataLst_ExRate")
            dgv.Columns("DiscountDataGridViewTextBoxColumn").HeaderText = rMgr.GetString("Col_DataLst_Discnt")
            dgv.Columns("SubTotalDataGridViewTextBoxColumn").HeaderText = rMgr.GetString("Col_DataLst_SubTot")
            dgv.Columns("GSTAmtDataGridViewTextBoxColumn").HeaderText = rMgr.GetString("Col_DataLst_GstAmt")
            dgv.Columns("TotalAmtDataGridViewTextBoxColumn").HeaderText = rMgr.GetString("Col_DataLst_TotAmt")
            dgv.Columns("PaidAmtDataGridViewTextBoxColumn").HeaderText = rMgr.GetString("Col_DataLst_PaidAmt")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub SetCulture(ByVal CultureName As String) Implements SalesInterface.MobileSales.ISalesBase.SetCulture
        sLang = CultureName
        '   Localization()
    End Sub

 End Class