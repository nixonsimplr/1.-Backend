Imports System.Data.SqlClient
Imports SalesInterface.MobileSales

Public Class Daily
    Implements ISalesBase
    Private aAgent As New ArrayList()
    Private aTerms As New ArrayList()
    Dim strPay As String = " "
    Dim strTerms As String
    Private Sub Daily_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub
    Private Sub Daily_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ConnectDB()
        'ExecuteSQL("Drop table InvoiceList")
        'ExecuteSQL("SELECT Customer.CustNo, Customer.CustName, Customer.PaymentMethod, Invoice.InvNo, Invoice.InvDt, Invoice.AgentID, Invoice.Discount, Invoice.SubTotal, Invoice.GST, Invoice.TotalAmt, Invoice.PayTerms into InvoiceList FROM Customer INNER JOIN Invoice ON Customer.CustNo = Invoice.CustId")
        LoadDescription()
        loadCombo()
        cmbTerms.SelectedIndex = 0
    End Sub
    Public Sub LoadDescription()
        'Dim dtr As SqlDataReader
        'dtr = ReadRecord("Select Description from PayTerms")
        'Do While dtr.Read = True
        '    If dtr("Description") <> "" Then
        '        cmbTerms.Items.Add(dtr(0))
        '    End If
        'Loop
        'dtr.Close()
        'If cmbTerms.Items.Count > 0 Then
        '    cmbTerms.SelectedIndex = 0
        'End If

        'Dim iIndex As Integer = 0
        'Dim iSelIndex As Integer
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
        dtr = ReadRecord("Select Distinct Code, Name from SalesAgent order by Name")
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

    Private Sub btnReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
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
        'Dim strSql As String = "SELECT Customer.CustNo, Customer.CustName," & SafeSQL(cmbTerms.Text) & " as PaymentMethod, Invoice.InvNo, Invoice.InvDt," & SafeSQL(cmbAgent.Text) & " as AgentID, Invoice.Discount, Invoice.SubTotal, Invoice.GSTAmt as GST, Invoice.TotalAmt, Invoice.PayTerms FROM Customer INNER JOIN Invoice ON Customer.CustNo = Invoice.CustId where InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & strTerms
        'Dim strSql As String = "SELECT OrderHdr.OrdNo as InvNo, OrderHdr.OrdDt as InvDt, Customer.CustNo, Customer.CustName, OrderHdr.PayTerms, SalesAgent.Name as AgentId, OrdItem.ItemNo, OrdItem.Description, OrdItem.UOM, OrdItem.Qty, OrdItem.Price, OrderHdr.SubTotal as SubAmt, OrderHdr.Discount, OrderHdr.GSTAmt, OrderHdr.TotalAmt, OrderHdr.CompanyName, OrderHdr.Remarks, Item.Category, " & SafeSQL(Format(dtpFromDate.Value, "dd/MM/yyyy")) & " as FromDate, " & SafeSQL(Format(dtpToDate.Value, "dd/MM/yyyy")) & " as ToDate FROM  Item INNER JOIN OrdItem ON Item.ItemNo = OrdItem.ItemNo INNER JOIN OrderHdr ON Item.CompanyName = OrderHdr.CompanyName and OrdItem.OrdNo = OrderHdr.OrdNo INNER JOIN Customer ON OrderHdr.CompanyName = Customer.CompanyName and OrderHdr.CustId = Customer.CustNo INNER JOIN SalesAgent ON OrderHdr.AgentId = SalesAgent.Code where OrderHdr.OrdNo not in (Select OrdNo from Invoice) and OrderHdr.void=0 and OrderHdr.OrdDt Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & strTerms & strPay & " Union SELECT Distinct OrderHDr.OrdNo as InvNo , OrderHDr.OrdDt as InvDT, NewCust.CustID as CustNo, NewCust.CustName, 'COD' as PayTerms, SalesAgent.Name as AgentId, OrdItem.ItemNo, OrdItem.Description, OrdItem.UOM, OrdItem.Qty, OrdItem.Price, OrderHDr.SubTotal as SubAmt, OrdItem.Discount AS Discount, OrderHDr.GstAmt, OrderHDr.TotalAmt, OrderHdr.CompanyName, OrderHdr.Remarks, Item.Category, " & SafeSQL(Format(dtpFromDate.Value, "dd/MM/yyyy 00:00:00")) & " as FromDate, " & SafeSQL(Format(dtpToDate.Value, "dd/MM/yyyy 23:59:59")) & " as ToDate  FROM Item INNER JOIN OrdItem ON Item.ItemNo = OrdItem.ItemNo INNER JOIN OrderHdr ON OrdItem.OrdNo = OrderHDr.OrdNo INNER JOIN NewCust ON NewCust.CompanyName = orderHdr.CompanyName and OrderHDr.CustId = NewCust.CustID INNER JOIN SalesAgent ON OrderHDr.AgentId = SalesAgent.Code where OrderHdr.OrdNo not in (Select OrdNo from Invoice) and OrderHDr.OrdDt Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & strTerms & strPay & " Order by Category "
        Dim strSql As String = "SELECT OrderHdr.OrdNo as InvNo, OrderHdr.OrdDt as InvDt, Customer.CustNo, Customer.CustName, OrderHdr.PayTerms, SalesAgent.Name as AgentId, OrdItem.ItemNo, OrdItem.Description, OrdItem.UOM, OrdItem.Qty, OrdItem.Price, OrderHdr.SubTotal as SubAmt, OrderHdr.Discount, OrderHdr.GSTAmt, OrderHdr.TotalAmt, OrderHdr.Remarks, Item.Category, " & SafeSQL(Format(dtpFromDate.Value, "dd/MM/yyyy")) & " as FromDate, " & SafeSQL(Format(dtpToDate.Value, "dd/MM/yyyy")) & " as ToDate FROM  Item INNER JOIN OrdItem ON Item.ItemNo = OrdItem.ItemNo INNER JOIN OrderHdr ON OrdItem.OrdNo = OrderHdr.OrdNo INNER JOIN Customer ON OrderHdr.CustId = Customer.CustNo INNER JOIN SalesAgent ON OrderHdr.AgentId = SalesAgent.Code where OrderHdr.OrdNo not in (Select OrdNo from Invoice) and OrderHdr.void=0 and OrderHdr.OrdDt Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & strTerms & strPay & " Union SELECT Distinct OrderHDr.OrdNo as InvNo , OrderHDr.OrdDt as InvDT, NewCust.CustID as CustNo, NewCust.CustName, 'COD' as PayTerms, SalesAgent.Name as AgentId, OrdItem.ItemNo, OrdItem.Description, OrdItem.UOM, OrdItem.Qty, OrdItem.Price, OrderHDr.SubTotal as SubAmt, OrdItem.Discount AS Discount, OrderHDr.GstAmt, OrderHDr.TotalAmt, OrderHdr.Remarks, Item.Category, " & SafeSQL(Format(dtpFromDate.Value, "dd/MM/yyyy 00:00:00")) & " as FromDate, " & SafeSQL(Format(dtpToDate.Value, "dd/MM/yyyy 23:59:59")) & " as ToDate  FROM Item INNER JOIN OrdItem ON Item.ItemNo = OrdItem.ItemNo INNER JOIN OrderHdr ON OrdItem.OrdNo = OrderHDr.OrdNo INNER JOIN NewCust ON OrderHDr.CustId = NewCust.CustID INNER JOIN SalesAgent ON OrderHDr.AgentId = SalesAgent.Code where OrderHdr.OrdNo not in (Select OrdNo from Invoice) and OrderHDr.OrdDt Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & strTerms & strPay & " Order by Category "
        'ExecuteReport(strSql, "DailySalesRep")
        ExecuteReport(strSql, "SalesOrderDetailRep")
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