Imports System.Data
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales


Public Class MonthlySales
    Implements ISalesBase
    Private aTerms As New ArrayList()
    Private aAgent As New ArrayList()
    Dim strPay As String = " "
    Dim strPay1 As String = " "
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

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        If cmbAgent.Text = "ALL" Then
            strPay = " "
            strPay1 = " "
        Else
            strPay = " and Invoice.Agentid=" & SafeSQL(cmbAgent.SelectedValue)
            strPay1 = " and RH.SalesPersonCode=" & SafeSQL(cmbAgent.SelectedValue)
        End If
        ExecuteSQL("Delete from MonthlySalesReport")
        Dim strSql As String = "SELECT Customer.CustNo, Customer.CustName, SalesAgent.Name as AgentID, sum(Invoice.SubTotal) Amt FROM Customer, Invoice, SalesAgent where Customer.CustNo = Invoice.CustId and SalesAgent.Code=Invoice.AgentID and (Invoice.Void <> 1 or Invoice.Void is null) and InvDt Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & strPay & " group by Customer.CustNo, Customer.CustName, SalesAgent.Name"
        Dim dtr As SqlDataReader
        '        Dim iCnt As Integer = 0
        dtr = ReadRecord(strSql)
        While dtr.Read = True
            ExecuteAnotherSQL("Insert into MonthlySalesReport(AgentID, CustNo, CustName, GRNAmt, Amt, FromDate, ToDate) values(" & SafeSQL(dtr("AgentID").ToString) & "," & SafeSQL(dtr("CustNo").ToString) & "," & SafeSQL(dtr("CustName").ToString) & ",0," & SafeSQL(Format(dtr("Amt"), "0.00")) & "," & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & "," & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & ")")
        End While
        dtr.Close()
        'change the query, only take the data from invoice 29/05/2008
        'strSql = "SELECT Customer.CustNo, Customer.CustName, SalesAgent.Name as AgentID, sum(TotalAmt) Amt FROM Customer, OrderHdr, SalesAgent where Customer.CustNo = OrderHdr.CustId and SalesAgent.Code=OrderHdr.AgentID and (OrderHdr.Void <> 1 or OrderHdr.Void is null) and OrdDt Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & " and Ordno not in (select Ordno from Invoice where (Invoice.Void <> 1 or Invoice.Void is null)) " & strPay1 & " group by Customer.CustNo, Customer.CustName, SalesAgent.Name"
        'dtr = ReadRecord(strSql)
        'While dtr.Read = True
        '    If IsCustomerExists(dtr("CustNo").ToString) Then
        '        ExecuteAnotherSQL("Update MonthlySalesReport set Amt = Amt+" & SafeSQL(Format(dtr("Amt"), "0.00")) & " where CustNo=" & SafeSQL(dtr("CustNo").ToString))
        '    Else
        '        ExecuteAnotherSQL("Insert into MonthlySalesReport(AgentID, CustNo, CustName, Amt, FromDate, ToDate) values(" & SafeSQL(dtr("AgentID").ToString) & "," & SafeSQL(dtr("CustNo").ToString) & "," & SafeSQL(dtr("CustName").ToString) & "," & SafeSQL(Format(dtr("Amt"), "0.00")) & "," & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & "," & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & ")")
        '    End If
        'End While
        'dtr.Close()
        'Add Goods return value
        strSql = "SELECT Customer.CustNo, Customer.CustName, SalesAgent.Name as AgentID, sum(RH.SubTotal) Amt FROM Customer, GoodsReturn as RH, SalesAgent where Customer.CustNo = RH.CustNo and SalesAgent.Code=RH.SalesPersonCode and (RH.Void <> 1 or RH.Void is null) and ReturnDate Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & strPay1 & " group by Customer.CustNo, Customer.CustName, SalesAgent.Name"
        dtr = ReadRecord(strSql)
        While dtr.Read = True
            If IsCustomerExists(dtr("CustNo").ToString) Then
                ExecuteAnotherSQL("Update MonthlySalesReport set GRNAmt = GRNAmt+" & SafeSQL(Format(dtr("Amt"), "0.00")) & " where CustNo=" & SafeSQL(dtr("CustNo").ToString))
            Else
                ExecuteAnotherSQL("Insert into MonthlySalesReport(AgentID, CustNo, CustName,GRNAmt, Amt, FromDate, ToDate) values(" & SafeSQL(dtr("AgentID").ToString) & "," & SafeSQL(dtr("CustNo").ToString) & "," & SafeSQL(dtr("CustName").ToString) & "," & SafeSQL(Format(dtr("Amt"), "0.00")) & ",0," & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & "," & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & ")")
            End If
        End While
        dtr.Close()
        strSql = "Select * from MonthlySalesReport order by Amt desc"
        '& " Union SELECT DISTINCT Customer.CustNo, Customer.CustName, .InvNo, Invoice.InvDt, Invoice.Discount, Invoice.SubTotal, Invoice.GstAmt AS GST, Invoice.TotalAmt, 'COD' as PayTerms, SalesAgent.Name as AgentID , Invoice.VOID FROM NewCust INNER JOIN Invoice ON NewCust.CustID = Invoice.CustId INNER JOIN SalesAgent ON Invoice.AgentId = SalesAgent.Code where InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & strTerms
        ExecuteReport(strSql, "MonthlySalesRep")

    End Sub
    Private Function IsCustomerExists(ByVal sCustNo As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = ReadRecordAnother("Select CustNo from MonthlySalesReport where CustNo = " & SafeSQL(sCustNo))
        bAns = dtr.Read
        dtr.Close()
        Return bAns
    End Function


    Private Sub MonthlySales_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub

    Private Sub MonthlySales_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ConnectDB()
        'ExecuteSQL("Drop table InvoiceList")
        'ExecuteSQL("SELECT Customer.CustNo, Customer.CustName, Customer.PaymentMethod, Invoice.InvNo, Invoice.InvDt, Invoice.AgentID, Invoice.Discount, Invoice.SubTotal, Invoice.GST, Invoice.TotalAmt, Invoice.PayTerms into InvoiceList FROM Customer INNER JOIN Invoice ON Customer.CustNo = Invoice.CustId")

        loadCombo()
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
End Class