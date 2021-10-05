Imports System.Data
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Public Class ReturnNote
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

    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        If cmbAgent.Text = "ALL" Then
            strPay = " "
            strPay1 = " "
        Else
            strPay = "and OrderHdr.Agentid=" & "'" & cmbAgent.SelectedValue & "'"
            strPay1 = "and GoodsReturn.SalesPersonCode=" & "'" & cmbAgent.SelectedValue & "'"
        End If
        ExecuteSQL("Delete from ReturnNoteRep")
        Dim strSql As String = "SELECT DISTINCT Customer.CustNo, Customer.CustName, OrderHdr.OrdNo, OrderHdr.OrdDt, SalesAgent.Name as AgentID FROM Customer, OrderHdr, SalesAgent where Customer.CustNo = OrderHdr.CustId and SalesAgent.Code=OrderHdr.AgentID  and OrdDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & " and Ordno not in (select Ordno from Invoice) " & strPay
        Dim dtr As SqlDataReader
        '        Dim iCnt As Integer = 0
        dtr = ReadRecord(strSql)
        While dtr.Read = True
            ExecuteAnotherSQL("Insert into ReturnNoteRep(AgentID, CustNo, CustName, OrdDate, OrdNo, ReturnNo) values(" & SafeSQL(dtr("AgentID").ToString) & "," & SafeSQL(dtr("CustNo").ToString) & "," & SafeSQL(dtr("CustName").ToString) & "," & SafeSQL(Format(dtpDate.Value, "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("OrdNo").ToString) & ",'')")
        End While
        dtr.Close()
        strSql = "SELECT DISTINCT Customer.CustNo, Customer.CustName, ReturnNo, ReturnDate as OrdDate, SalesAgent.Name as AgentID FROM Customer, GoodsReturn, SalesAgent where Customer.CustNo = GoodsReturn.CustNo and SalesAgent.Code=GoodsReturn.SalesPersonCode  and ReturnDate Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & " and ReturnNo not in (select GoodsReturnNo from CreditNote)" & strPay1
        dtr = ReadRecord(strSql)
        While dtr.Read = True
            If IsCustomerExists(dtr("CustNo").ToString) Then
                ExecuteAnotherSQL("Update ReturnNoteRep set ReturnNo =" & SafeSQL(dtr("ReturnNo").ToString) & " where CustNo=" & SafeSQL(dtr("CustNo").ToString))
            Else
                ExecuteAnotherSQL("Insert into ReturnNoteRep(AgentID, CustNo, CustName, OrdDate, OrdNo, ReturnNo) values(" & SafeSQL(dtr("AgentID").ToString) & "," & SafeSQL(dtr("CustNo").ToString) & "," & SafeSQL(dtr("CustName").ToString) & "," & SafeSQL(Format(dtpDate.Value, "yyyyMMdd HH:mm:ss")) & ",''," & SafeSQL(dtr("ReturnNo").ToString) & ")")
            End If
        End While
        dtr.Close()
        strSql = "Select * from ReturnNoteRep order by CustNo"
        '& " Union SELECT DISTINCT Customer.CustNo, Customer.CustName, .InvNo, Invoice.InvDt, Invoice.Discount, Invoice.SubTotal, Invoice.GstAmt AS GST, Invoice.TotalAmt, 'COD' as PayTerms, SalesAgent.Name as AgentID , Invoice.VOID FROM NewCust INNER JOIN Invoice ON NewCust.CustID = Invoice.CustId INNER JOIN SalesAgent ON Invoice.AgentId = SalesAgent.Code where InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & strTerms
        ExecuteReport(strSql, "ReturnNoteRep")
    End Sub
    Private Function IsCustomerExists(ByVal sCustNo As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = ReadRecordAnother("Select CustNo from ReturnNoteRep where CustNo = " & SafeSQL(sCustNo))
        bAns = dtr.Read
        dtr.Close()
        Return bAns
    End Function
    Private Sub ReturnNote_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub

    Private Sub ReturnNote_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
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