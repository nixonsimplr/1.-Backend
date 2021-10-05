Imports System.Data
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Public Class DailySales
    Implements ISalesBase
    Private aTerms As New ArrayList()
    Private aAgent As New ArrayList()
    Dim strPay As String = " "
    Dim strTerms As String
    Dim sAgentID As String
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        'Dim sql As String = "Select * from InvoiceList where InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59"))
        'MsgBox(sql)
        If cmbAgent.Text = "ALL" Then
            strPay = " "
        Else
            strPay = "and Invoice.Agentid=" & "'" & cmbAgent.SelectedValue & "'"
        End If
        If cmbTerms.Text = "ALL" Then
            strTerms = ""
        Else
            strTerms = " and Invoice.Payterms='" & cmbTerms.SelectedValue & "'"
        End If
        Dim dtr As SqlDataReader
        dtr = ReadRecord("SELECT  PreInvNo as Agent FROM  MDT WHERE AgentID=" & SafeSQL(cmbAgent.SelectedValue))
        If dtr.Read = True Then
            sAgentID = dtr("Agent").ToString
        Else
            sAgentID = "0"
        End If
        dtr.Close()
        'strSql = "SELECT Invoice.InvNo, Invoice.InvDt, Customer.CustNo, Customer.CustName, Invoice.PayTerms, Invoice.AgentId, InvItem.ItemNo, Item.Description, InvItem.UOM, InvItem.Qty, InvItem.Price, InvItem.SubAmt, InvItem.GstAmt, Invoice.TotalAmt FROM  Item INNER JOIN InvItem ON Item.ItemNo = InvItem.ItemNo INNER JOIN Invoice ON InvItem.InvNo = Invoice.InvNo INNER JOIN Customer ON Invoice.CustId = Customer.CustNo where Invoice.InvDt = " & SafeSQL(CmbInvNo.Text) & " and Invoice.PayTerms= " & SafeSQL(cmbPayTerms.Text)
        'Dim strSql As String = "SELECT Customer.CustNo, Customer.CustName,OrderHdr.PayTerms as PaymentMethod, OrderHdr.OrdNo as InvNo, OrderHdr.OrdDt as InvDt,OrderHdr.AgentID, OrderHdr.Discount, OrderHdr.SubTotal, OrderHdr.GSTAmt as GST, OrderHdr.TotalAmt, OrderHdr.PayTerms FROM Customer INNER JOIN OrderHdr ON Customer.CustNo = OrderHdr.CustId where OrderHdr.void=0 and OrdDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & strTerms & " Order by OrderHdr.OrdNo"
        '12032008New Cust) Dim strSql As String = "SELECT DISTINCT Customer.CustNo, Customer.CustName, Invoice.PayTerms AS PaymentMethod, Invoice.InvNo, Invoice.InvDt, Invoice.Discount, Invoice.SubTotal, Invoice.GstAmt AS GST, Invoice.TotalAmt, PayTerms.Description as PayTerms, SalesAgent.Name as AgentID FROM Customer INNER JOIN Invoice ON Customer.CustNo = Invoice.CustId INNER JOIN SalesAgent ON Invoice.AgentId = SalesAgent.Code INNER JOIN PayTerms ON Invoice.PayTerms = PayTerms.Code where Invoice.void=0 and InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & strTerms & " Order by Invoice.InvNo"
        'Dim strSql As String = "SELECT Distinct Invoice.InvNo, Invoice.InvDt, Customer.CustNo, Customer.CustName, Invoice.PayTerms, Invoice.AgentId, InvItem.ItemNo, Item.Description, InvItem.UOM, InvItem.Qty, InvItem.Price, InvItem.SubAmt, InvItem.GstAmt, Invoice.TotalAmt FROM  Item INNER JOIN InvItem ON Item.ItemNo = InvItem.ItemNo INNER JOIN Invoice ON InvItem.InvNo = Invoice.InvNo INNER JOIN Customer ON Invoice.CustId = Customer.CustNo where Invoice.InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strTerms & strPay & " order by Invoice.Invno"
        Dim strSql As String
        If cmbAgent.Text = "ALL" Then
            strSql = "SELECT DISTINCT Customer.CustNo, Customer.CustName, Invoice.PayTerms AS PaymentMethod, Invoice.InvNo, Invoice.InvDt, Invoice.Discount, Invoice.SubTotal, Invoice.PrintNo AS GST, Invoice.TotalAmt, PayTerms.Description as PayTerms, " & SafeSQL(sAgentID) & " + ' / ' + SalesAgent.Name as AgentID, Invoice.Void FROM Customer INNER JOIN Invoice ON Customer.CustNo = Invoice.CustId INNER JOIN SalesAgent ON Invoice.AgentId = SalesAgent.Code INNER JOIN PayTerms ON Invoice.PayTerms = PayTerms.Code where InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & strTerms '& " Order By Products.ProductType, Products.CategoryID, Products.Brand, Products.Packsize, Products.ShortDesc" '& " Union SELECT DISTINCT NewCust.CustID, NewCust.CustName, 'COD' AS PaymentMethod, Invoice.InvNo, Invoice.InvDt, Invoice.Discount, Invoice.SubTotal, Invoice.GstAmt AS GST, Invoice.TotalAmt, 'COD' as PayTerms, SalesAgent.Name as AgentID , Invoice.VOID FROM NewCust INNER JOIN Invoice ON NewCust.CustID = Invoice.CustId INNER JOIN SalesAgent ON Invoice.AgentId = SalesAgent.Code where InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & strTerms
        Else
            strSql = "SELECT DISTINCT Customer.CustNo, Customer.CustName, Invoice.PayTerms AS PaymentMethod, Invoice.InvNo, Invoice.InvDt, Invoice.Discount, Invoice.SubTotal, Invoice.PrintNo AS GST, Invoice.TotalAmt, PayTerms.Description as PayTerms, " & SafeSQL(sAgentID) & " + ' / ' + SalesAgent.Name as AgentID, Invoice.Void FROM Customer INNER JOIN Invoice ON Customer.CustNo = Invoice.CustId INNER JOIN SalesAgent ON Invoice.AgentId = SalesAgent.Code INNER JOIN PayTerms ON Invoice.PayTerms = PayTerms.Code where InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & strTerms '&'& " Union SELECT DISTINCT NewCust.CustID, NewCust.CustName, 'COD' AS PaymentMethod, Invoice.InvNo, Invoice.InvDt, Invoice.Discount, Invoice.SubTotal, Invoice.GstAmt AS GST, Invoice.TotalAmt, 'COD' as PayTerms, SalesAgent.Name as AgentID , Invoice.VOID FROM NewCust INNER JOIN Invoice ON NewCust.CustID = Invoice.CustId INNER JOIN SalesAgent ON Invoice.AgentId = SalesAgent.Code where InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & strTerms
        End If
        ExecuteReport(strSql, "DailySalesRep")
    End Sub

    Private Sub DailySales_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'ExecuteSQL("Drop table InvoiceList")
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
        dtr = ReadRecord("Select Code, Description from PayTerms")
        cmbTerms.DataSource = Nothing
        aTerms.Clear()
        aTerms.Add(New ComboValues("ALL", "ALL"))
        While dtr.Read()
            aTerms.Add(New ComboValues(dtr("Code").ToString, dtr("Description").ToString))
            '    iSelIndex = iIndex
            'End If
            'iIndex = iIndex + 1
        End While
        dtr.Close()
        cmbTerms.DataSource = aTerms
        cmbTerms.DisplayMember = "Desc"
        cmbTerms.ValueMember = "Code"
        cmbTerms.SelectedIndex = 0

    End Sub
    Public Sub loadCombo()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Code, Name from SalesAgent Order By Name")
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

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged

    End Sub

    Private Sub Label20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

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