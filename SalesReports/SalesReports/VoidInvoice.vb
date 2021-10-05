Imports System.Data.SqlClient
Imports SalesInterface.MobileSales

Public Class VoidInvoice
    Implements ISalesBase


    Dim strPay As String = ""
    Private Sub VoidInvoice_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ConnectDB()
        Dim strDel = "Drop table VoidInvoice"
        ExecuteSQL(strDel)
        Dim strSqlSales = "Select Receipt.RcptNo, Invoice.InvNo, Invoice.PaidAmt, Invoice.Void, Invoice.AgentId, Invoice.InvDt, Customer.CustNo, Invoice.VoidDate INTO VoidInvoice FROM Invoice INNER JOIN Receipt ON Invoice.CustId = Receipt.CustID INNER JOIN Customer ON Invoice.CustId = Customer.CustNo"
        ExecuteSQL(strSqlSales)
        loadCombo()
    End Sub
    Public Sub loadCombo()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Code from SalesAgent")
        cmbAgent.Items.Add("Select")
        Do While dtr.Read = True
            cmbAgent.Items.Add(dtr(0))
        Loop
        dtr.Close()
        If cmbAgent.Items.Count > 0 Then
            cmbAgent.SelectedIndex = 0
        End If
    End Sub
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim strSql As String = "Select * from VoidInvoice where Void=1" & strPay
        ExecuteReport(strSql, "VoidInvoiceRep")
    End Sub

    Private Sub cmbAgent_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAgent.SelectedIndexChanged
        If cmbAgent.Text = "All" Then
            strPay = " "
        Else
            strPay = " and Agentid=" & SafeSQL(cmbAgent.Text)
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