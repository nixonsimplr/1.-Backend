Imports System.Data
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Public Class DailySales
    Implements ISalesBase
    Private aAgent As New ArrayList()
    Private aTerms As New ArrayList()
    Dim strPay As String = " "
    Dim strTerms As String
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        'Dim sql As String = "Select * from InvoiceList where InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59"))
        'MsgBox(sql)
        If cmbAgent.Text = "ALL" Then
            strPay = " "
        Else
            strPay = "and OrderHdr.Agentid=" & "'" & cmbAgent.SelectedValue & "'"
        End If
        If cmbTerms.Text = "ALL" Then
            strTerms = ""
        Else
            strTerms = " and OrderHdr.Payterms='" & cmbTerms.SelectedValue & "'"
        End If
        'strSql = "SELECT Invoice.InvNo, Invoice.InvDt, Customer.CustNo, Customer.CustName, Invoice.PayTerms, Invoice.AgentId, InvItem.ItemNo, Item.Description, InvItem.UOM, InvItem.Qty, InvItem.Price, InvItem.SubAmt, InvItem.GstAmt, Invoice.TotalAmt FROM  Item INNER JOIN InvItem ON Item.ItemNo = InvItem.ItemNo INNER JOIN Invoice ON InvItem.InvNo = Invoice.InvNo INNER JOIN Customer ON Invoice.CustId = Customer.CustNo where Invoice.InvDt = " & SafeSQL(CmbInvNo.Text) & " and Invoice.PayTerms= " & SafeSQL(cmbPayTerms.Text)
        'Dim strSql As String = "SELECT Customer.CustNo, Customer.CustName,OrderHdr.PayTerms as PaymentMethod, OrderHdr.OrdNo as InvNo, OrderHdr.OrdDt as InvDt,OrderHdr.AgentID, OrderHdr.Discount, OrderHdr.SubTotal, OrderHdr.GSTAmt as GST, OrderHdr.TotalAmt, OrderHdr.PayTerms FROM Customer INNER JOIN OrderHdr ON Customer.CustNo = OrderHdr.CustId where OrderHdr.void=0 and OrdDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & strTerms & " Order by OrderHdr.OrdNo"
        'Dim strSql As String = "SELECT Invoice.InvNo, Invoice.InvDt, Customer.CustNo, Customer.CustName, Invoice.PayTerms, Invoice.AgentId, InvItem.ItemNo, Item.Description, InvItem.UOM, InvItem.Qty, InvItem.Price, InvItem.SubAmt, InvItem.GstAmt, Invoice.TotalAmt FROM  Item INNER JOIN InvItem ON Item.ItemNo = InvItem.ItemNo INNER JOIN Invoice ON InvItem.InvNo = Invoice.InvNo INNER JOIN Customer ON Invoice.CustId = Customer.CustNo where Invoice.InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strTerms & strPay & " Order by Invoice.InvNo"
        'Dim strSql As String = "SELECT DISTINCT Customer.CustNo, Customer.CustName, Invoice.PayTerms AS PaymentMethod, Invoice.InvNo, Invoice.InvDt, Invoice.Discount, Invoice.SubTotal, Invoice.GstAmt AS GST, Invoice.TotalAmt, PayTerms.Description as PayTerms, SalesAgent.Name as AgentID, Invoice.Void FROM Customer INNER JOIN Invoice ON Customer.CustNo = Invoice.CustId INNER JOIN SalesAgent ON Invoice.AgentId = SalesAgent.Code INNER JOIN PayTerms ON Invoice.PayTerms = PayTerms.Code where InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & strTerms & " Union SELECT DISTINCT NewCust.CustID, NewCust.CustName, 'COD' AS PaymentMethod, Invoice.InvNo, Invoice.InvDt, Invoice.Discount, Invoice.SubTotal, Invoice.GstAmt AS GST, Invoice.TotalAmt, 'COD' as PayTerms, SalesAgent.Name as AgentID , Invoice.VOID FROM NewCust INNER JOIN Invoice ON NewCust.CustID = Invoice.CustId INNER JOIN SalesAgent ON Invoice.AgentId = SalesAgent.Code where InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & strTerms
        'Dim strSql As String = "SELECT Distinct OrderHDr.OrdNo as InvNo, OrderHDr.OrdDt as InvDt, Customer.CustNo, Customer.SearchName as CustName, OrderHDr.PayTerms, SalesAgent.Name as AgentId, OrderHDr.SubTotal, OrderHDr.Discount, OrderHDr.GstAmt as GST, OrderHDr.TotalAmt, OrderHdr.CompanyName, " & SafeSQL(Format(dtpFromDate.Value, "dd/MM/yyyy")) & " as FromDate, " & SafeSQL(Format(dtpToDate.Value, "dd/MM/yyyy")) & " as ToDate  FROM  OrderHDr, Customer, SalesAgent Where Customer.CompanyName = OrderHdr.CompanyName and OrderHDr.AgentID=SalesAgent.Code and OrderHDr.CustId = Customer.CustNo and OrderHDr.OrdNo not in (Select OrdNo from Invoice) and (Void=0 or Void is Null) and OrderHdr.OrdDt Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & strTerms & strPay & " Union SELECT OrderHDr.OrdNo as InvNo, OrderHDr.OrdDt as InvDt, NewCust.CustID, NewCust.CustName as CustName, OrderHDr.PayTerms, SalesAgent.Name as AgentId, OrderHDr.SubTotal, OrderHDr.Discount, OrderHDr.GstAmt as GST, OrderHDr.TotalAmt, OrderHdr.CompanyName, " & SafeSQL(Format(dtpFromDate.Value, "dd/MM/yyyy")) & " as FromDate, " & SafeSQL(Format(dtpToDate.Value, "dd/MM/yyyy")) & " as ToDate FROM NewCust INNER JOIN OrderHdr ON NewCust.CompanyName = OrderHdr.CompanyName and NewCust.CustID = OrderHdr.CustId INNER JOIN SalesAgent ON OrderHdr.AgentId = SalesAgent.Code where OrdDt Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & strPay & strTerms & " Order by OrderHdr.OrdNo"
        Dim strSql As String = "SELECT Distinct OrderHDr.OrdNo as InvNo, OrderHDr.OrdDt as InvDt, Customer.CustNo, Customer.SearchName as CustName, OrderHDr.PayTerms, SalesAgent.Name as AgentId, OrderHDr.SubTotal, OrderHDr.Discount, OrderHDr.GstAmt as GST, OrderHDr.TotalAmt, " & SafeSQL(Format(dtpFromDate.Value, "dd/MM/yyyy")) & " as FromDate, " & SafeSQL(Format(dtpToDate.Value, "dd/MM/yyyy")) & " as ToDate  FROM  OrderHDr, Customer, SalesAgent Where OrderHDr.AgentID=SalesAgent.Code and OrderHDr.CustId = Customer.CustNo and OrderHDr.OrdNo not in (Select OrdNo from Invoice) and (Void=0 or Void is Null) and OrderHdr.OrdDt Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & strTerms & strPay & " Union SELECT OrderHDr.OrdNo as InvNo, OrderHDr.OrdDt as InvDt, NewCust.CustID, NewCust.CustName as CustName, OrderHDr.PayTerms, SalesAgent.Name as AgentId, OrderHDr.SubTotal, OrderHDr.Discount, OrderHDr.GstAmt as GST, OrderHDr.TotalAmt, " & SafeSQL(Format(dtpFromDate.Value, "dd/MM/yyyy")) & " as FromDate, " & SafeSQL(Format(dtpToDate.Value, "dd/MM/yyyy")) & " as ToDate FROM NewCust INNER JOIN OrderHdr ON NewCust.CustID = OrderHdr.CustId INNER JOIN SalesAgent ON OrderHdr.AgentId = SalesAgent.Code where OrdDt Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & strPay & strTerms & " Order by OrderHdr.OrdNo"
        'ExecuteReport(strSql, "DailySalesRep")
        ExecuteReport(strSql, "SalesOrderSummaryRep")
    End Sub

    Private Sub DailySales_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        DisconnectDB()
    End Sub
    Private Sub DailySales_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ConnectDB()
        'ExecuteSQL("Drop table InvoiceList")
        'ExecuteSQL("SELECT Customer.CustNo, Customer.CustName, Customer.PaymentMethod, Invoice.InvNo, Invoice.InvDt, Invoice.AgentID, Invoice.Discount, Invoice.SubTotal, Invoice.GST, Invoice.TotalAmt, Invoice.PayTerms into InvoiceList FROM Customer INNER JOIN Invoice ON Customer.CustNo = Invoice.CustId")
        LoadDescription()
        loadCombo()
        cmbTerms.SelectedIndex = 0
    End Sub
    Public Sub LoadDescription()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Code, Description from PayTerms order by Description")
        cmbTerms.DataSource = Nothing
        aTerms.Clear()
        aTerms.Add(New ComboValues("ALL", "ALL"))
        While dtr.Read()
            aTerms.Add(New ComboValues(dtr("Code").ToString, dtr("Description").ToString))
        End While
        dtr.Close()
        cmbTerms.DataSource = aTerms
        cmbTerms.DisplayMember = "Desc"
        cmbTerms.ValueMember = "Code"
        cmbTerms.SelectedIndex = 0
    End Sub
    Public Sub loadCombo()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select distinct Code, Name from SalesAgent order by Name")
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

    Private Sub cmbAgent_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAgent.SelectedIndexChanged

    End Sub

    Private Sub cmbTerms_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbTerms.SelectedIndexChanged

    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpToDate.ValueChanged

    End Sub

    Private Sub Label20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label20.Click

    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

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