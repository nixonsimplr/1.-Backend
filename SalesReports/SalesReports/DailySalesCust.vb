Imports System.Data
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales

Public Class DailySalesCust
    Implements ISalesBase

    Dim strPay As String = " "
    Public Sub LoadItem()
        Dim dtr As SqlDataReader
        cmbCustomer.Items.Clear()
        dtr = ReadRecord("Select ItemNo from Item")
        cmbCustomer.Items.Add("All")
        Do While dtr.Read = True
            cmbCustomer.Items.Add(dtr("ItemNo"))
        Loop
        dtr.Close()
        If cmbCustomer.Items.Count > 0 Then
            cmbCustomer.SelectedIndex = 0
        End If
        dtpDate.Value = Date.Now
    End Sub
    Public Sub LoadCustomer()
        Dim dtr As SqlDataReader
        cmbCustomer.Items.Clear()
        dtr = ReadRecord("Select CustNo from Customer")
        cmbCustomer.Items.Add("All")
        Do While dtr.Read = True
            cmbCustomer.Items.Add(dtr("CustNo"))
        Loop
        dtr.Close()
        If cmbCustomer.Items.Count > 0 Then
            cmbCustomer.SelectedIndex = 0
        End If
        dtpDate.Value = Date.Now
    End Sub
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim strSql As String = ""
        If rbCustomer.Checked = True Then
            strSql = "SELECT Customer.CustNo, Customer.CustName, Invoice.InvNo, Invoice.InvDt, Invoice.TotalAmt, Invoice.PayTerms FROM Customer INNER JOIN Invoice ON Customer.CustNo = Invoice.CustId where InvDt Between" & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & "Order by TotalAmt"
        End If
        If rbItem.Checked = True Then
            strSql = "SELECT Invoice.InvDt, InvItem.InvNo, InvItem.ItemNo as CustNo, InvItem.Qty as TotalAmt, Item.Description as CustName, Invoice.PayTerms FROM Item INNER JOIN InvItem ON Item.ItemNo = InvItem.ItemNo INNER JOIN Invoice ON InvItem.InvNo = Invoice.InvNo  where InvDt Between" & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & " Order by Qty"
        End If
        ExecuteReport(strSql, "DailySalesAmtRpt")
    End Sub

    Private Sub cmbCustomer_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCustomer.SelectedIndexChanged
        If cmbCustomer.Text = "All" Then
            strPay = " "
        ElseIf rbCustomer.Checked = True Then
            strPay = "and CustNo=" & SafeSQL(cmbCustomer.Text)
        Else
            strPay = "and Item.ItemNo=" & SafeSQL(cmbCustomer.Text)
        End If
    End Sub

    Private Sub DailySalesCust_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        DisconnectDB()
    End Sub

    Private Sub DailySalesCust_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ConnectDB()
        rbCustomer.Checked = True
    End Sub

    Private Sub rbCustomer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbCustomer.CheckedChanged
        If rbCustomer.Checked = True Then
            lblItem.Text = "Customer. . . . . . . . . . . . ."
            LoadCustomer()
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

    Private Sub rbItem_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbItem.CheckedChanged
        If rbItem.Checked = True Then
            lblItem.Text = "Item. . . . . . . . . . . . . ."
            LoadItem()
        End If
    End Sub

    Public Sub ReturnSearch(ByVal SQL As String) Implements SalesInterface.MobileSales.ISalesBase.ReturnSearch

    End Sub

    Public Event SearchField(ByVal FormName As String, ByVal FieldName As String, ByVal FieldType As String, ByVal CurValue As String) Implements SalesInterface.MobileSales.ISalesBase.SearchField

    Public Sub SetCulture(ByVal CultureName As String) Implements SalesInterface.MobileSales.ISalesBase.SetCulture

    End Sub
End Class