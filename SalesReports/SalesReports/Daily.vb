Imports System.Data.SqlClient
Imports SalesInterface.MobileSales

Public Class Daily
    Implements ISalesBase
    Private aTerms As New ArrayList()
    Private aAgent As New ArrayList()
    Dim strPay As String = " "
    Dim strTerms As String
    Dim sAgentID As String

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
        dtr = ReadRecord("Select Distinct Code, Name from SalesAgent order by name")
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
        btnPrint.Enabled = False
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
        'Dim strSql As String
        'strSql = "SELECT Invoice.InvNo, Invoice.InvDt, Customer.CustNo, Customer.CustName, Invoice.PayTerms, Invoice.AgentId, InvItem.ItemNo, Item.Description, InvItem.UOM, InvItem.Qty, InvItem.Price, InvItem.SubAmt, InvItem.GstAmt, Invoice.TotalAmt FROM  Item INNER JOIN InvItem ON Item.ItemNo = InvItem.ItemNo INNER JOIN Invoice ON InvItem.InvNo = Invoice.InvNo INNER JOIN Customer ON Invoice.CustId = Customer.CustNo where Invoice.InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & strTerms
        'Dim strSql As String = "SELECT Customer.CustNo, Customer.CustName," & SafeSQL(cmbTerms.Text) & " as PaymentMethod, Invoice.InvNo, Invoice.InvDt," & SafeSQL(cmbAgent.Text) & " as AgentID, Invoice.Discount, Invoice.SubTotal, Invoice.GSTAmt as GST, Invoice.TotalAmt, Invoice.PayTerms FROM Customer INNER JOIN Invoice ON Customer.CustNo = Invoice.CustId where InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & strTerms
        'Dim strSql As String = "SELECT OrderHdr.OrdNo as InvNo, OrderHdr.OrdDt as InvDt, Customer.CustNo, Customer.CustName, OrderHdr.PayTerms, OrderHdr.AgentId, OrdItem.ItemNo, Item.Description, OrdItem.UOM, OrdItem.Qty, OrdItem.Price, OrdItem.SubAmt as SubAmtItem, OrderHdr.SubTotal as SubAmt, OrderHdr.Discount, OrderHdr.GSTAmt, OrderHdr.TotalAmt FROM  Item INNER JOIN OrdItem ON Item.ItemNo = OrdItem.ItemNo INNER JOIN OrderHdr ON OrdItem.OrdNo = OrderHdr.OrdNo INNER JOIN Customer ON OrderHdr.CustId = Customer.CustNo where OrderHdr.void=0 and OrderHdr.OrdDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strTerms & strPay
        'ExecuteReport(strSql, "DailySalesRep")
        'Dim strSql As String = "SELECT Invoice.InvNo, Invoice.InvDt, Customer.CustNo, Customer.CustName, Invoice.PayTerms, Invoice.AgentId, InvItem.ItemNo, Item.Description, InvItem.UOM, InvItem.Qty, InvItem.Price, Invoice.SubTotal as SubAmt,Invoice.Discount, Invoice.GSTAmt as GstAmt, Invoice.TotalAmt FROM  Item INNER JOIN InvItem ON Item.ItemNo = InvItem.ItemNo INNER JOIN Invoice ON InvItem.InvNo = Invoice.InvNo INNER JOIN Customer ON Invoice.CustId = Customer.CustNo where Invoice.void=0 and Invoice.InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strTerms & strPay
        '(Comm on 12032008 to add the New Cust)Dim strSql As String = "SELECT Distinct Invoice.InvNo, Invoice.InvDt, Customer.CustNo, Customer.CustName, Payterms.Description as PayTerms, SalesAgent.Name as AgentId, InvItem.ItemNo, Item.Description, InvItem.UOM, InvItem.Qty, InvItem.Price, InvItem.Qty*InvItem.Price AS Discount, Invoice.SubTotal as SubAmt, Invoice.GstAmt, Invoice.TotalAmt FROM Item INNER JOIN InvItem ON Item.ItemNo = InvItem.ItemNo INNER JOIN Invoice ON InvItem.InvNo = Invoice.InvNo INNER JOIN Customer ON Invoice.CustId = Customer.CustNo INNER JOIN SalesAgent ON Invoice.AgentId = SalesAgent.Code INNER JOIN PayTerms ON Invoice.PayTerms = PayTerms.Code where Invoice.void=0 and Invoice.InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strTerms & strPay
        Dim strSql As String
        'If cmbAgent.Text = "ALL" Then
        '    strSql = "SELECT Distinct Invoice.InvNo, Invoice.InvDt, Customer.CustNo, Customer.CustName, Payterms.Description as PayTerms, 'ALL' as AgentId, InvItem.ItemNo, InvItem.Description, InvItem.UOM, InvItem.Qty, InvItem.Price, InvItem.Qty*InvItem.Price AS Discount, Invoice.SubTotal as SubAmt, Invoice.GstAmt, Invoice.TotalAmt FROM Item INNER JOIN InvItem ON Item.ItemNo = InvItem.ItemNo INNER JOIN Invoice ON InvItem.InvNo = Invoice.InvNo INNER JOIN Customer ON Invoice.CustId = Customer.CustNo INNER JOIN SalesAgent ON Invoice.AgentId = SalesAgent.Code INNER JOIN PayTerms ON Invoice.PayTerms = PayTerms.Code where Invoice.InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strTerms & strPay & "Order By Item.ProdType, Item.Category, Item.Brand, Item.Packsize, Item.Description"
        '    '& " Union SELECT Distinct Invoice.InvNo, Invoice.InvDt, NewCust.CustID, NewCust.CustName, 'COD' as PayTerms, SalesAgent.Name as AgentId, InvItem.ItemNo, Item.Description, InvItem.UOM, InvItem.Qty, InvItem.Price, InvItem.Qty*InvItem.Price AS Discount, Invoice.SubTotal as SubAmt, Invoice.GstAmt, Invoice.TotalAmt FROM Item INNER JOIN InvItem ON Item.ItemNo = InvItem.ItemNo INNER JOIN Invoice ON InvItem.InvNo = Invoice.InvNo INNER JOIN NewCust ON Invoice.CustId = NewCust.CustID INNER JOIN SalesAgent ON Invoice.AgentId = SalesAgent.Code where Invoice.InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strTerms & strPay
        'Else
        '    strSql = "SELECT Distinct Invoice.InvNo, Invoice.InvDt, Customer.CustNo, Customer.CustName, Payterms.Description as PayTerms, " & SafeSQL(sAgentID) & " + ' / ' + SalesAgent.Name as AgentId, InvItem.ItemNo, InvItem.Description, InvItem.UOM, InvItem.Qty, InvItem.Price, InvItem.Qty*InvItem.Price AS Discount, Invoice.SubTotal as SubAmt, Invoice.GstAmt, Invoice.TotalAmt FROM Item INNER JOIN InvItem ON Item.ItemNo = InvItem.ItemNo INNER JOIN Invoice ON InvItem.InvNo = Invoice.InvNo INNER JOIN Customer ON Invoice.CustId = Customer.CustNo INNER JOIN SalesAgent ON Invoice.AgentId = SalesAgent.Code INNER JOIN PayTerms ON Invoice.PayTerms = PayTerms.Code where Invoice.InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strTerms & strPay & "Order By Item.ProdType, Item.Category, Item.Brand, Item.Packsize, Item.Description"
        '    '& " Union SELECT Distinct Invoice.InvNo, Invoice.InvDt, NewCust.CustID, NewCust.CustName, 'COD' as PayTerms, SalesAgent.Name as AgentId, InvItem.ItemNo, Item.Description, InvItem.UOM, InvItem.Qty, InvItem.Price, InvItem.Qty*InvItem.Price AS Discount, Invoice.SubTotal as SubAmt, Invoice.GstAmt, Invoice.TotalAmt FROM Item INNER JOIN InvItem ON Item.ItemNo = InvItem.ItemNo INNER JOIN Invoice ON InvItem.InvNo = Invoice.InvNo INNER JOIN NewCust ON Invoice.CustId = NewCust.CustID INNER JOIN SalesAgent ON Invoice.AgentId = SalesAgent.Code where Invoice.InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strTerms & strPay
        'End If
        ''     Dim strSql As String = "SELECT Distinct Invoice.InvNo, Invoice.InvDt, Customer.CustNo, Customer.CustName, Payterms.Description as PayTerms, " & SafeSQL(sAgentID) & " + ' / ' + SalesAgent.Name as AgentId, InvItem.ItemNo, InvItem.Description, InvItem.UOM, InvItem.Qty, InvItem.Price, InvItem.Qty*InvItem.Price AS Discount, Invoice.SubTotal as SubAmt, Invoice.GstAmt, Invoice.TotalAmt FROM Item INNER JOIN InvItem ON Item.ItemNo = InvItem.ItemNo INNER JOIN Invoice ON InvItem.InvNo = Invoice.InvNo INNER JOIN Customer ON Invoice.CustId = Customer.CustNo INNER JOIN SalesAgent ON Invoice.AgentId = SalesAgent.Code INNER JOIN PayTerms ON Invoice.PayTerms = PayTerms.Code where Invoice.InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strTerms & strPay & " Union SELECT Distinct Invoice.InvNo, Invoice.InvDt, NewCust.CustID, NewCust.CustName, 'COD' as PayTerms, SalesAgent.Name as AgentId, InvItem.ItemNo, Item.Description, InvItem.UOM, InvItem.Qty, InvItem.Price, InvItem.Qty*InvItem.Price AS Discount, Invoice.SubTotal as SubAmt, Invoice.GstAmt, Invoice.TotalAmt FROM Item INNER JOIN InvItem ON Item.ItemNo = InvItem.ItemNo INNER JOIN Invoice ON InvItem.InvNo = Invoice.InvNo INNER JOIN NewCust ON Invoice.CustId = NewCust.CustID INNER JOIN SalesAgent ON Invoice.AgentId = SalesAgent.Code where Invoice.InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strTerms & strPay

        If cmbAgent.Text = "ALL" Then
            strSql = "SELECT Distinct Invoice.InvNo, Invoice.InvDt, Customer.CustNo, Customer.CustName, Payterms.Description as PayTerms, 'ALL' as AgentId, InvItem.ItemNo, case left(InvItem.Description,3) when 'EX:' then 'EX: ' + Item.Description when 'FOC' then 'FOC: ' + Item.Description else Item.Description end as Description, InvItem.UOM, InvItem.Qty, InvItem.Price, InvItem.Qty*InvItem.Price AS Discount, Invoice.SubTotal as SubAmt, Invoice.GstAmt, Invoice.TotalAmt,Item.Category, Item.Brand FROM Item INNER JOIN InvItem ON Item.ItemNo = InvItem.ItemNo INNER JOIN Invoice ON InvItem.InvNo = Invoice.InvNo INNER JOIN Customer ON Invoice.CustId = Customer.CustNo INNER JOIN SalesAgent ON Invoice.AgentId = SalesAgent.Code INNER JOIN PayTerms ON Invoice.PayTerms = PayTerms.Code where Invoice.InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strTerms & strPay & " Order By InvItem.ItemNo"
            '& " Union SELECT Distinct Invoice.InvNo, Invoice.InvDt, NewCust.CustID, NewCust.CustName, 'COD' as PayTerms, SalesAgent.Name as AgentId, InvItem.ItemNo, Item.Description, InvItem.UOM, InvItem.Qty, InvItem.Price, InvItem.Qty*InvItem.Price AS Discount, Invoice.SubTotal as SubAmt, Invoice.GstAmt, Invoice.TotalAmt FROM Item INNER JOIN InvItem ON Item.ItemNo = InvItem.ItemNo INNER JOIN Invoice ON InvItem.InvNo = Invoice.InvNo INNER JOIN NewCust ON Invoice.CustId = NewCust.CustID INNER JOIN SalesAgent ON Invoice.AgentId = SalesAgent.Code where Invoice.InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strTerms & strPay
        Else
            strSql = "SELECT Distinct Invoice.InvNo, Invoice.InvDt, Customer.CustNo, Customer.CustName, Payterms.Description as PayTerms, " & SafeSQL(sAgentID) & " + ' / ' + SalesAgent.Name as AgentId, InvItem.ItemNo,  case left(InvItem.Description,3) when 'EX:' then 'EX: ' + Item.Description when 'FOC' then 'FOC: ' + Item.Description else Item.Description end as Description, InvItem.UOM, InvItem.Qty, InvItem.Price, InvItem.Qty*InvItem.Price AS Discount, Invoice.SubTotal as SubAmt, Invoice.GstAmt, Invoice.TotalAmt,Item.Category, Item.Brand FROM Item INNER JOIN InvItem ON Item.ItemNo = InvItem.ItemNo INNER JOIN Invoice ON InvItem.InvNo = Invoice.InvNo INNER JOIN Customer ON Invoice.CustId = Customer.CustNo INNER JOIN SalesAgent ON Invoice.AgentId = SalesAgent.Code INNER JOIN PayTerms ON Invoice.PayTerms = PayTerms.Code where Invoice.InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strTerms & strPay & " Order By InvItem.ItemNo"
            '& " Union SELECT Distinct Invoice.InvNo, Invoice.InvDt, NewCust.CustID, NewCust.CustName, 'COD' as PayTerms, SalesAgent.Name as AgentId, InvItem.ItemNo, Item.Description, InvItem.UOM, InvItem.Qty, InvItem.Price, InvItem.Qty*InvItem.Price AS Discount, Invoice.SubTotal as SubAmt, Invoice.GstAmt, Invoice.TotalAmt FROM Item INNER JOIN InvItem ON Item.ItemNo = InvItem.ItemNo INNER JOIN Invoice ON InvItem.InvNo = Invoice.InvNo INNER JOIN NewCust ON Invoice.CustId = NewCust.CustID INNER JOIN SalesAgent ON Invoice.AgentId = SalesAgent.Code where Invoice.InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strTerms & strPay
        End If

        ExecuteReport(strSql, "DailyReport")
        btnPrint.Enabled = True
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

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub cmbTerms_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbTerms.SelectedIndexChanged

    End Sub
End Class