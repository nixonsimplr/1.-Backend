Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Public Class DailySalesOrder
    Implements ISalesBase


    Private aTerms As New ArrayList()
    Dim strPay As String = " "
    Dim strTerms As String
    Dim strDelivery As String = "and Delivered='1'"
    Private Sub DailySalesOrder_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub
    Private Sub DailySalesOrder_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ConnectDB()
        LoadDescription()
        loadCombo()
    End Sub
    Public Sub LoadDescription()
        Dim iIndex As Integer = 0
        Dim iSelIndex As Integer
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Code, Description from PayTerms")
        cmbTerms.DataSource = Nothing
        aTerms.Clear()
        While dtr.Read()
            aTerms.Add(New ComboValues(dtr("Code").ToString, dtr("Description").ToString))

            '    iSelIndex = iIndex
            'End If
            iIndex = iIndex + 1
        End While
        dtr.Close()
        cmbTerms.DataSource = aTerms
        cmbTerms.DisplayMember = "Desc"
        cmbTerms.ValueMember = "Code"
        cmbTerms.SelectedIndex = iSelIndex
    End Sub
    Public Sub loadCombo()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select distinct(Code) from SalesAgent")
        cmbAgent.Items.Add("All")
        Do While dtr.Read = True
            cmbAgent.Items.Add(dtr(0))
        Loop
        dtr.Close()
        If cmbAgent.Items.Count > 0 Then
            cmbAgent.SelectedIndex = 0
        End If

        If cmbDelivered.Items.Count > 0 Then
            cmbDelivered.SelectedIndex = 2
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

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        strTerms = "PayTerms=" & SafeSQL(cmbTerms.SelectedValue)
        Dim strSql As String = "SELECT OrderHdr.OrdNo, OrderHdr.OrdDt, OrderHdr.CustId, Customer.CustName, OrderHdr.SubTotal, OrderHdr.Discount, OrderHdr.GSTAmt, OrderHdr.TotalAmt, OrderHdr.PayTerms, OrderHdr.Delivered, OrderHdr.AgentID FROM OrderHdr, Customer where Customer.CustNo=OrderHdr.CustID and OrdDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & "and " & strTerms & strDelivery
        ExecuteReport(strSql, "DailySalesOrderRep")
    End Sub

    Private Sub cmbDelivered_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbDelivered.SelectedIndexChanged
        If cmbDelivered.Text = "All" Then
            strDelivery = " "
        ElseIf cmbDelivered.Text = "Yes" Then
            strDelivery = "and Delivered='1'"
        Else
            strDelivery = "and Delivered='0'"
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