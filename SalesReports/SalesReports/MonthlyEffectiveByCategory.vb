Imports System.Data
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales

Public Class MonthlyEffectiveByCategory
    Implements ISalesBase
    Private aCat As New ArrayList()
    Private aAgent As New ArrayList()
    Private aCust As New ArrayList
    Dim strPay As String = " "
    Dim strPay1 As String = " "

    Dim strAgent As String = " "
    Dim strCat As String = " "
    Dim strCat1 As String = " "
    Dim strTerms As String

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


    Private Sub MonthlyEffectiveByCategory_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub

    Private Sub MonthlyEffectiveByCategory_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ConnectDB()
        'ExecuteSQL("Drop table InvoiceList")
        'ExecuteSQL("SELECT Customer.CustNo, Customer.CustName, Customer.PaymentMethod, Invoice.InvNo, Invoice.InvDt, Invoice.AgentID, Invoice.Discount, Invoice.SubTotal, Invoice.GST, Invoice.TotalAmt, Invoice.PayTerms into InvoiceList FROM Customer INNER JOIN Invoice ON Customer.CustNo = Invoice.CustId")
        LoadCategory()
        loadCombo()
        loadCustomer()
        dtpFromDate.Value = Format(Date.Now, "MM/01/yyyy")
    End Sub

    Public Sub loadCombo()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Code, Name from SalesAgent")
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
        dtr = ReadRecord("Select CustNo, CustName from Customer where Active = 1 order by CustNo")
        cmbCust.DataSource = Nothing
        aCust.Clear()
        aCust.Add(New ComboValues("ALL", "ALL"))
        While dtr.Read()
            aCust.Add(New ComboValues(dtr("CustNo").ToString, dtr("CustNo").ToString & " - " & dtr("CustName").ToString))
        End While
        dtr.Close()
        cmbCust.DataSource = aCust
        cmbCust.DisplayMember = "Desc"
        cmbCust.ValueMember = "Code"
        cmbCust.SelectedIndex = 0

    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        If dtpToDate.Value.Month <> dtpFromDate.Value.Month Then
            MsgBox("Sorry, please select date on the same month !!")
            Exit Sub
        End If
        If dtpToDate.Value.Day < dtpFromDate.Value.Day Then
            MsgBox("Sorry, day value on To Date should be bigger than day value on From Date !!")
            Exit Sub
        End If
        btnPrint.Enabled = False
        If cmbAgent.Text = "ALL" Then
            strAgent = " "
        Else
            strAgent = "and SalesAgent.Code=" & "'" & cmbAgent.SelectedValue & "'"
        End If
        If cmbCategory.Text = "ALL" Then
            strCat = ""
            strCat1 = ""
        Else
            strCat = " and InvItem.ItemNo in (Select ItemNo from Item where Category='" & cmbCategory.SelectedValue.ToString.Trim & "')"
            strCat1 = " and OrdItem.ItemNo in (Select ItemNo from Item where Category='" & cmbCategory.SelectedValue.ToString.Trim & "')"
        End If
        ExecuteSQL("Delete from MonthlyEfByCatRep")
        Dim dtr As SqlDataReader
        'change the query, only take the data from invoice 29/05/2008
        'dtr = ReadRecord("SELECT Customer.CustNo, Customer.CustName, SalesAgent.Name as AgentID FROM Customer, OrderHdr, SalesAgent, OrdItem where OrderHdr.OrdNo = OrdItem.OrdNo and Customer.CustNo = OrderHdr.CustId and SalesAgent.Code=OrderHdr.AgentID and OrderHdr.OrdNo not in (select OrdNo from Invoice) and OrdDt Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & strAgent & strCat1 & _
        '        " UNION SELECT Customer.CustNo, Customer.CustName, SalesAgent.Name as AgentID FROM Customer, Invoice, SalesAgent, InvItem where Invoice.InvNo = InvItem.InvNo and Customer.CustNo = Invoice.CustId and SalesAgent.Code=Invoice.AgentID and InvDt Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & strAgent & strCat)
        dtr = ReadRecord("SELECT distinct Customer.CustNo, Customer.CustName, SalesAgent.Name as AgentID, InvItem.ItemNo, InvItem.Description as ItemName FROM Customer, Invoice, SalesAgent, InvItem where Invoice.InvNo = InvItem.InvNo and ((InvItem.Description not like 'EX%' and InvItem.Description not like 'FOC:%') or InvItem.Description is null) and Customer.CustNo = Invoice.CustId and SalesAgent.Code=Invoice.AgentID and InvDt Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & strAgent & strCat)
        While dtr.Read = True
            If dtpToDate.Value.Day >= dtpFromDate.Value.Day Then
                For i As Integer = dtpToDate.Value.Day To dtpFromDate.Value.Day Step -1
                    ExecuteAnotherSQL("Insert into MonthlyEfByCatRep(CustNo, CustName, DayOfMonth, FromDate, ToDate, AgentId,ItemNo, ITemName,Cnt) values(" & SafeSQL(dtr("CustNo").ToString) & "," & SafeSQL(dtr("CustName").ToString) & "," & i & "," & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 23:59:59")) & "," & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & "," & SafeSQL(dtr("AgentID").ToString) & "," & SafeSQL(dtr("ItemNo").ToString) & "," & SafeSQL(dtr("ItemName").ToString) & ",0)")
                Next
            End If
        End While
        dtr.Close()
        ExecuteAnotherSQL("Update MonthlyEfByCatRep set Category = " & SafeSQL(cmbCategory.Text))
        'change the query, only take the data from invoice 29/05/2008
        'dtr = ReadRecord("SELECT Customer.CustNo,day(OrdDt) DayOfMonth FROM Customer, OrderHdr, SalesAgent, OrdItem where OrderHdr.OrdNo = OrdItem.OrdNo and Customer.CustNo = OrderHdr.CustId and SalesAgent.Code=OrderHdr.AgentID and OrderHdr.OrdNo not in (select OrdNo from Invoice) and OrdDt Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & strAgent & strCat1 & _
        '        " UNION SELECT Customer.CustNo, Day(InvDt) DayOfMonth FROM Customer, Invoice, SalesAgent, InvItem where Invoice.InvNo = InvItem.InvNo and Customer.CustNo = Invoice.CustId and SalesAgent.Code=Invoice.AgentID and InvDt Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & strAgent & strCat)
        dtr = ReadRecord("SELECT distinct Customer.CustNo, Day(InvDt) DayOfMonth, InvItem.ItemNo, sum(Qty * UOM.BaseQty) as Qty FROM Customer, Invoice, SalesAgent, InvItem, UOM where Invoice.InvNo = InvItem.InvNo and ((InvItem.Description not like 'EX%' and InvItem.Description not like 'FOC:%') or InvItem.Description is null) and Customer.CustNo = Invoice.CustId and SalesAgent.Code=Invoice.AgentID and UOM.ItemNo = InvItem.ItemNo and InvDt Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & strAgent & strCat & " group by Customer.CustNo, Day(InvDt), InvItem.ItemNo")
        While dtr.Read = True
            ExecuteAnotherSQL("Update MonthlyEfByCatRep set Cnt = cnt + " & dtr("Qty").ToString & " where CustNo = " & SafeSQL(dtr("CustNo").ToString) & " and DayOfMonth = " & SafeSQL(dtr("DayOfMonth").ToString) & " and ItemNo = " & SafeSQL(dtr("ItemNo").ToString))
        End While
        dtr.Close()
        Dim strSql As String = ""

        ExecuteAnotherSQL("Update MonthlyEfByCatRep set MonthlyEfByCatRep.ItemName = Item.ShortDesc from MonthlyEfByCatRep, Item where MonthlyEfByCatRep.ItemNo = Item.ItemNo and (MonthlyEfByCatRep.ItemName = '' or MonthlyEfByCatRep.ItemName is null)")

        strsql = "Select distinct * from MonthlyEfByCatRep order by CustNo"
        ExecuteReport(strsql, "MonthlyEffectiveByCategoryRep")
        btnPrint.Enabled = True
    End Sub
End Class