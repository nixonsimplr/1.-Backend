Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Public Class PayCollection
    Implements ISalesBase
    Private aAgent As New ArrayList()
    Private aTerms As New ArrayList()
    Dim strPay As String = "and PayMethod='Cash'"
    Dim strPay1 As String = " "

    Private Sub PayCollection_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        DisconnectDB()
    End Sub

    Private Sub PayCollection_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ConnectDB()
        loadCombo()
        LoadDescription()
    End Sub

    Public Sub LoadDescription()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Code, Description from PayMethod order by Description")
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
        cmbTerms.SelectedIndex = 0
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        btnPrint.Enabled = False
        If cmbAgent.Text = "ALL" Then
            strPay1 = " "
        Else
            strPay1 = "and Receipt.AgentID=" & SafeSQL(cmbAgent.SelectedValue)
        End If
        If cmbTerms.Text = "ALL" Then
            strPay = " "
        Else
            strPay = "and PayMethod.Description='" & cmbTerms.Text & "'"
        End If
        'Dim strSql As String = "SELECT Distinct Receipt.RcptNo, Receipt.RcptDt, Customer.CustNo, Customer.CustName, RcptItem.InvNo, RcptItem.AmtPaid as PaidAmt,SalesAgent.Name as AgentID, PayMethod.Description as PayMethod, (Receipt.BankName + ' ' + Receipt.ChqNo + ' ' + convert(varchar(10),Receipt.ChqDt,101)) as ChqNo, Receipt.CompanyName FROM Receipt, Customer, RcptItem, SalesAgent, PayMethod WHERE Receipt.CompanyName = Customer.CompanyName and (Receipt.void=0 or Receipt.Void is Null) and Receipt.CustId = CustNo and SalesAgent.Code=Receipt.AgentID and PayMethod.code=Receipt.Paymethod and Receipt.RcptNo = RcptItem.RcptNo and Receipt.RcptDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & strPay1 & " Order by Receipt.RcptDt, Receipt.RcptNo, RcptItem.InvNo"
        '(isnull(TotalAmt,0) - isnull(PaidAmt,0)) as SubTotal
        Dim strSql As String = ""
        strSql = "SELECT Distinct Receipt.RcptNo, Receipt.RcptDt, Customer.CustNo, Customer.CustName, RcptItem.InvNo, RcptItem.AmtPaid as PaidAmt,SalesAgent.Name as AgentID, PayMethod.Description as PayMethod, (Receipt.BankName + '   ' + Receipt.ChqNo + '    ' + convert(varchar(10),Receipt.ChqDt,101)) as ChqNo, (isnull(Invoice.TotalAmt,0) - isnull(Invoice.PaidAmt,0)) as SubTotal, Receipt.Remarks as CompanyName  FROM Receipt, Customer, RcptItem, SalesAgent, PayMethod, Invoice WHERE Invoice.InvNo = RcptItem.InvNo and (Receipt.void=0 or Receipt.Void is Null) and Receipt.CustId = Customer.CustNo and SalesAgent.Code=Receipt.AgentID and PayMethod.code=Receipt.Paymethod and Receipt.RcptNo = RcptItem.RcptNo and Receipt.RcptDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & strPay1
        strSql = strSql & " UNION "
        strSql = strSql & "SELECT Distinct Receipt.RcptNo, Receipt.RcptDt, Customer.CustNo, Customer.CustName, RcptItem.InvNo, RcptItem.AmtPaid as PaidAmt,SalesAgent.Name as AgentID, PayMethod.Description as PayMethod, (Receipt.BankName + '   ' + Receipt.ChqNo + '    ' + convert(varchar(10),Receipt.ChqDt,101)) as ChqNo, (isnull(CreditNote.TotalAmt,0) - isnull(CreditNote.PaidAmt,0)) as SubTotal, Receipt.Remarks as CompanyName FROM Receipt, Customer, RcptItem, SalesAgent, PayMethod, CreditNote WHERE CreditNote.CreditNoteNo = RcptItem.InvNo and (Receipt.void=0 or Receipt.Void is Null) and Receipt.CustId = Customer.CustNo and SalesAgent.Code=Receipt.AgentID and PayMethod.code=Receipt.Paymethod and Receipt.RcptNo = RcptItem.RcptNo and Receipt.RcptDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & strPay1
        strSql = strSql & " UNION "
        strSql = strSql & "SELECT DISTINCT Receipt.RcptNo, Receipt.RcptDt, Customer.CustNo, Customer.CustName, RcptItem.InvNo, isnull(RcptItem.AmtPaid,0) AS PaidAmt, " & _
        "SalesAgent.Name, PayMethod.Description AS PayMethod, Receipt.BankName + '   ' + Receipt.ChqNo + '    ' + CONVERT(varchar(10), Receipt.ChqDt, 101) AS ChqNo " & _
        ", -(isnull(RcptItem.AmtPaid,0)) AS SubTotal, Receipt.Remarks as CompanyName FROM RcptItem INNER JOIN Receipt ON RcptItem.RcptNo = Receipt.RcptNo INNER JOIN SalesAgent ON " & _
        "Receipt.AgentID = SalesAgent.Code INNER JOIN Customer ON Receipt.CustID = Customer.CustNo INNER JOIN PayMethod ON Receipt.PayMethod = PayMethod.Code " & _
        "WHERE (RcptItem.InvNo = 'ADVPAYMENT') and (Receipt.Void = 0 or Receipt.Void is null) and Receipt.RcptDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & strPay1
        strSql = strSql & " Order by RcptDt, RcptNo, InvNo "
        ExecuteReport(strSql, "PayCollectionRep")
        btnPrint.Enabled = True
    End Sub

    Private Sub cmbTerms_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbTerms.SelectedIndexChanged

    End Sub

    Private Sub cmbAgent_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAgent.SelectedIndexChanged

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