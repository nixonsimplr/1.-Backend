Imports System.Data
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Public Class MonthlyStatement
    Implements ISalesBase

    Private aCust As New ArrayList()

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim strSql As String = ""
        strSql = "Select CustBill.CustNo, CustBill.CustName, CustBill.Address as Add1, CustBill.Address2 as Add2, CustBill.City + ' ' + CustBill.PostCode as Add3,Customer.CustName as Attn, Invoice.InvNo, Invoice.GstAmt, Invoice.SubTotal, Invoice.TotalAmt, " & SafeSQL(Format(dtpDate.Value, "MMMM yyyy")) & " as RefNo  from Invoice,Customer,Customer as CustBill where Customer.CustNo = Invoice.CustId and Customer.[Bill-toNo] = " & SafeSQL(cmbCust.SelectedValue) & " and CustBill.CustNo = Customer.[Bill-toNo] and month(Invoice.InvDt)= " & dtpDate.Value.Month & " and year(Invoice.InvDt) = " & dtpDate.Value.Year
        strSql = strSql & " UNION "
        strSql = strSql & " Select CustBill.CustNo, CustBill.CustName, CustBill.Address as Add1, CustBill.Address2 as Add2, CustBill.City + ' ' + CustBill.PostCode as Add3,Customer.CustName as Attn, CreditNote.CreditNoteNo, CreditNote.GST, CreditNote.SubTotal, CreditNote.TotalAmt, " & SafeSQL(Format(dtpDate.Value, "MMMM yyyy")) & " as RefNo  from CreditNote,Customer,Customer as CustBill where Customer.CustNo = CreditNote.CustNo and Customer.[Bill-toNo] = " & SafeSQL(cmbCust.SelectedValue) & " and CustBill.CustNo = Customer.[Bill-toNo] and month(CreditNote.CreditDate)= " & dtpDate.Value.Month & " and year(CreditNote.CreditDate) = " & dtpDate.Value.Year
        ExecuteReport(strSql, "MonthlyStatementB")
    End Sub
    Public Sub loadCombo()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select CustNo, CustName from Customer where active = 1 order by CustNo")
        cmbCust.DataSource = Nothing
        aCust.Clear()
        While dtr.Read()
            aCust.Add(New ComboValues(dtr("CustNo").ToString, dtr("CustNo").ToString & " - " & dtr("CustName").ToString))
        End While
        dtr.Close()
        cmbCust.DataSource = aCust
        cmbCust.DisplayMember = "Desc"
        cmbCust.ValueMember = "Code"
        cmbCust.SelectedIndex = 0
    End Sub

    Private Sub MonthlyStatement_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub

    Private Sub MonthlyStatement_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ConnectDB()
        loadCombo()
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

    Private Sub btnPrintA_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintA.Click
        Dim strSql As String = ""
        strSql = "Select CustBill.CustNo, CustBill.CustName, CustBill.Address as Add1, CustBill.Address2 as Add2, CustBill.City + ' ' + CustBill.PostCode as Add3,Customer.CustName as Attn, Invoice.InvNo, Invoice.GstAmt, Invoice.SubTotal, Invoice.TotalAmt, " & SafeSQL(Format(dtpDate.Value, "MMMM yyyy")) & " as RefNo  from Invoice,Customer,Customer as CustBill where Customer.CustNo = Invoice.CustId and Customer.[Bill-toNo] = " & SafeSQL(cmbCust.SelectedValue) & " and CustBill.CustNo = Customer.[Bill-toNo] and month(Invoice.InvDt)= " & dtpDate.Value.Month & " and year(Invoice.InvDt) = " & dtpDate.Value.Year
        strSql = strSql & " UNION "
        strSql = strSql & " Select CustBill.CustNo, CustBill.CustName, CustBill.Address as Add1, CustBill.Address2 as Add2, CustBill.City + ' ' + CustBill.PostCode as Add3,Customer.CustName as Attn, CreditNote.CreditNoteNo, CreditNote.GST, CreditNote.SubTotal, CreditNote.TotalAmt, " & SafeSQL(Format(dtpDate.Value, "MMMM yyyy")) & " as RefNo  from CreditNote,Customer,Customer as CustBill where Customer.CustNo = CreditNote.CustNo and Customer.[Bill-toNo] = " & SafeSQL(cmbCust.SelectedValue) & " and CustBill.CustNo = Customer.[Bill-toNo] and month(CreditNote.CreditDate)= " & dtpDate.Value.Month & " and year(CreditNote.CreditDate) = " & dtpDate.Value.Year
        ExecuteReport(strSql, "MonthlyStatementA")
    End Sub
End Class