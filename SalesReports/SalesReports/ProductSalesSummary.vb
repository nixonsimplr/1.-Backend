Imports System.Data
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales

Public Class ProductSalesSummary
    Implements ISalesBase

    Private aCat As New ArrayList()
    Private aAgent As New ArrayList()
    Private aFromCust As New ArrayList
    Private aToCust As New ArrayList


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


    Public Sub loadCombo()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Code, Name from SalesAgent order by Name")
        cmbAgent.DataSource = Nothing
        aAgent.Clear()
        aAgent.Add(New ComboValues("ALL", "ALL"))
        While dtr.Read()
            aAgent.Add(New ComboValues(dtr("Code").ToString, dtr("Name").ToString))

            '    iSelIndex = iIndex
            'End If
            'iIndex = iIndex + 1
        End While
        dtr.Close()
        cmbAgent.DataSource = aAgent
        cmbAgent.DisplayMember = "Desc"
        cmbAgent.ValueMember = "Code"
        cmbAgent.SelectedIndex = 0
    End Sub

    Public Sub LoadCategory()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Code, Description from Category order by description")
        cmbCategory.DataSource = Nothing
        aCat.Clear()
        aCat.Add(New ComboValues("ALL", "ALL"))
        While dtr.Read()
            aCat.Add(New ComboValues(dtr("Code").ToString, dtr("Description").ToString))
        End While
        dtr.Close()
        cmbCategory.DataSource = aCat
        cmbCategory.DisplayMember = "Desc"
        cmbCategory.ValueMember = "Code"
        cmbCategory.SelectedIndex = 0

    End Sub

    Public Sub LoadCustomer()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select CustNo, CustName from Customer where active = 1 order by CustNo")
        cmbFromCust.DataSource = Nothing
        aFromCust.Clear()
        aToCust.Clear()
        'aFromCust.Add(New ComboValues("ALL", "ALL"))
        While dtr.Read()
            aFromCust.Add(New ComboValues(dtr("CustNo").ToString, dtr("CustNo").ToString & " - " & dtr("CustName").ToString))
            aToCust.Add(New ComboValues(dtr("CustNo").ToString, dtr("CustNo").ToString & " - " & dtr("CustName").ToString))
        End While
        dtr.Close()
        cmbFromCust.DataSource = aFromCust
        cmbFromCust.DisplayMember = "Desc"
        cmbFromCust.ValueMember = "Code"
        cmbFromCust.SelectedIndex = 0

        cmbToCust.DataSource = aToCust
        cmbToCust.DisplayMember = "Desc"
        cmbToCust.ValueMember = "Code"
        cmbToCust.SelectedIndex = 0

    End Sub

    Private Sub ProductSalesSummary_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub


    Private Sub ProductSalesSummary_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ConnectDB()
        'ExecuteSQL("Drop table InvoiceList")
        'ExecuteSQL("SELECT Customer.CustNo, Customer.CustName, Customer.PaymentMethod, Invoice.InvNo, Invoice.InvDt, Invoice.AgentID, Invoice.Discount, Invoice.SubTotal, Invoice.GST, Invoice.TotalAmt, Invoice.PayTerms into InvoiceList FROM Customer INNER JOIN Invoice ON Customer.CustNo = Invoice.CustId")
        LoadCategory()
        loadCombo()
        LoadCustomer()
        dtpFromDate.Value = Format(Date.Now, "MM/01/yyyy")
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim strSql As String = "'"
        If dtpToDate.Value.Month <> dtpFromDate.Value.Month Then
            MsgBox("Sorry, please select date on the same month !!")
            Exit Sub
        End If
        If dtpToDate.Value.Day < dtpFromDate.Value.Day Then
            MsgBox("Sorry, day value on To Date should be bigger than day value on From Date !!")
            Exit Sub
        End If
        btnPrint.Enabled = False
        Dim strAgent As String = ""
        Dim sAgentName As String = ""
        Dim strAgentRet As String = ""
        Dim strCat As String = ""
        Dim strCust As String = ""
        Dim strCustRet As String = ""

        If cmbAgent.Text = "ALL" Then
            strAgent = " "
            strAgentRet = " "
            sAgentName = "ALL"
        Else
            strAgent = "and Invoice.AgentId=" & SafeSQL(cmbAgent.SelectedValue)
            strAgentRet = "and GR.SalesPersonCode=" & SafeSQL(cmbAgent.SelectedValue)
            sAgentName = cmbAgent.Text
        End If
        If cmbCategory.Text = "ALL" Then
            strCat = " "
        Else
            strCat = " and Item.Category=" & SafeSQL(cmbCategory.SelectedValue.ToString.Trim)
        End If
        If cmbFromCust.Text = "ALL" Then
            strCust = " "
            strCustRet = " "
        Else
            'strCust = " and Invoice.CustId = " & SafeSQL(cmbFromCust.SelectedValue.ToString.Trim)
            'strCustRet = " and GR.CustNo = " & SafeSQL(cmbFromCust.SelectedValue.ToString.Trim)
            strCust = " and Invoice.CustId between " & SafeSQL(cmbFromCust.SelectedValue.ToString.Trim) & " and " & SafeSQL(cmbToCust.SelectedValue.ToString.Trim)
            strCustRet = " and GR.CustNo between " & SafeSQL(cmbFromCust.SelectedValue.ToString.Trim) & " and " & SafeSQL(cmbToCust.SelectedValue.ToString.Trim)
        End If

        ExecuteSQL("Delete from ProductSalesSumRep")

        strSql = "Select ItemNo, ItemName ,sum(SalesQty) as SalesQty, sum(SalesAmt) as SalesAmt, sum(GRNQty) as GRNQty, sum(GRNAmt) as GRNAmt, sum(ExchangeQty) as ExchangeQty, sum(ExchangeAmt) as ExchangeAmt,Category, CustNo, CustName, " & SafeSQL(sAgentName) & " as AgentName, FromDate,ToDate from ( "

        'Get Sales Qty and Amt

        strSql = strSql & " select InvItem.ItemNo,Item.Description as ItemName,(InvItem.Qty * Uom.BaseQty) as SalesQty,InvItem.SubAmt as SalesAmt,0 as GRNQty, 0 as GRNAmt, 0 as ExchangeQty, 0 as ExchangeAmt, Item.Category as Category, " & SafeSQL(cmbFromCust.SelectedValue.ToString.Trim) & " as CustNo, " & SafeSQL(cmbToCust.SelectedValue.ToString.Trim) & "  as CustName,SalesAgent.Name as AgentName, " & SafeSQL(Format(dtpFromDate.Value, "dd/MM/yyyy")) & " as FromDate, " & SafeSQL(Format(dtpToDate.Value, "dd/MM/yyyy")) & " as ToDate "
        strSql = strSql & " from Invoice,InvItem,Item,UOM,Customer,SalesAgent "
        strSql = strSql & " where Invoice.InvNo = InvITem.InvNo and (InvItem.description not like 'EX%' or InvItem.description is null) and UOM.ItemNo = InvItem.ItemNo and UOM.Uom = InvItem.Uom "
        strSql = strSql & " and Item.ItemNo = InvItem.ItemNo " & strAgent & strCust & strCat
        strSql = strSql & " and Customer.CustNo = Invoice.CustId and SalesAgent.Code = Invoice.AgentId and Customer.Active = 1"
        strSql = strSql & " and (Invoice.Void = 0 or Invoice.Void is null) and Invoice.InvDt between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59"))
        strSql = strSql & " union all "
        'Get Exchange Qty and Amt
        strSql = strSql & " select InvItem.ItemNo,Item.Description as ItemName,0 as SalesQty,0 as SalesAmt,0 as GRNQty, 0 as GRNAmt, (InvItem.Qty * Uom.BaseQty) as ExchangeQty, InvItem.SubAmt as ExchangeAmt, Item.Category as Category, " & SafeSQL(cmbFromCust.SelectedValue.ToString.Trim) & " as CustNo, " & SafeSQL(cmbToCust.SelectedValue.ToString.Trim) & "  as CustName,SalesAgent.Name as AgentName, " & SafeSQL(Format(dtpFromDate.Value, "dd/MM/yyyy")) & " as FromDate, " & SafeSQL(Format(dtpToDate.Value, "dd/MM/yyyy")) & " as ToDate "
        strSql = strSql & " from Invoice,InvItem,Item,UOM,Customer,SalesAgent  "
        strSql = strSql & " where Invoice.InvNo = InvITem.InvNo and InvItem.description like 'EX%' and UOM.ItemNo = InvItem.ItemNo and UOM.Uom = InvItem.Uom "
        strSql = strSql & " and Item.ItemNo = InvItem.ItemNo " & strAgent & strCust & strCat
        strSql = strSql & " and Customer.CustNo = Invoice.CustId and SalesAgent.Code = Invoice.AgentId and Customer.Active = 1"
        strSql = strSql & " and (Invoice.Void = 0 or Invoice.Void is null) and Invoice.InvDt between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59"))
        strSql = strSql & " union all "
        'Get Goods Return Qty and Amt
        strSql = strSql & " select GI.ItemNo,Item.Description as ItemName,0 as SalesQty,0 as SalesAmt,(GI.Quantity * Uom.BaseQty) as GRNQty, GI.Amt as GRNAmt, 0 as ExchangeQty, 0 as ExchangeAmt, Item.Category as Category, " & SafeSQL(cmbFromCust.SelectedValue.ToString.Trim) & " as CustNo, " & SafeSQL(cmbToCust.SelectedValue.ToString.Trim) & " as CustName,SalesAgent.Name as AgentName, " & SafeSQL(Format(dtpFromDate.Value, "dd/MM/yyyy")) & " as FromDate, " & SafeSQL(Format(dtpToDate.Value, "dd/MM/yyyy")) & " as ToDate "
        strSql = strSql & " from GoodsReturn GR,GoodsReturnItem GI,Item,UOM ,Customer,SalesAgent "
        strSql = strSql & " where GR.ReturnNo = GI.ReturnNo and Item.ItemNo = GI.ItemNo and UOM.ItemNo = GI.ItemNo and UOM.Uom = GI.Uom " & strAgentRet & strCustRet & strCat
        strSql = strSql & " and Customer.CustNo = GR.CustNo and SalesAgent.Code = GR.SalesPersonCode and Customer.Active = 1"
        strSql = strSql & " and (GR.Void = 0 or GR.Void is null) and GR.ReturnDate between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59"))
        strSql = strSql & " ) A group by ItemNo,ItemName,Category,CustNo,CustName,FromDate,ToDate"
        'strSql = strSql & " ) A group by ItemNo,ItemName,Category,CustNo,CustName,AgentName,FromDate,ToDate"
        'strSql = "Select distinct * from ProductSalesSumRep order by CustNo"
        ExecuteReport(strSql, "ProductSalesSummaryRep")
        btnPrint.Enabled = True
    End Sub
End Class