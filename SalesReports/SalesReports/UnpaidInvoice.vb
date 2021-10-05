Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Public Class UnpaidInvoice
    Implements ISalesBase


    Dim strPay As String = " "
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim strSql As String = "SELECT Invoice.InvNo, Invoice.InvDt, Customer.CustName, Invoice.AgentID, Invoice.TotalAmt, Invoice.PaidAmt, Customer.CustNo, Invoice.DoNo FROM Invoice INNER JOIN Customer ON Invoice.CustId = Customer.CustNo  where InvDt Between" & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & " and (Invoice.PaidAmt<1) order by Invoice.TotalAmt Desc" '& "and " & strTerms
        ExecuteReport(strSql, "UnpaidInvRep")
    End Sub

    Private Sub UnpaidInvoice_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        DisconnectDB()
    End Sub
    Private Sub SalesReportItem_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ConnectDB()
        loadCombo()
    End Sub
    Public Sub loadCombo()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select AgentID from MDT")
        cmbAgent.Items.Add("All")
        Do While dtr.Read = True
            cmbAgent.Items.Add(dtr(0))
        Loop
        dtr.Close()
        If cmbAgent.Items.Count > 0 Then
            cmbAgent.SelectedIndex = 0
        End If
        dtpDate.Value = Date.Now
    End Sub

    Private Sub cmbAgent_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAgent.SelectedIndexChanged
        If cmbAgent.Text = "All" Then
            strPay = " "
        Else
            strPay = "and Agentid=" & "'" & cmbAgent.Text & "'"
        End If
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