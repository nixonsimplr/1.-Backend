Imports System.Data.SqlClient
Imports SalesInterface.MobileSales

Public Class BankInDate
    Implements ISalesBase
    Private aAgent As New ArrayList()
    Dim strAgent As String = " "

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

    Public Sub loadCombo()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Code, Name from SalesAgent, MDT where MDT.AgentID=SalesAgent.Code order by Name")
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

    Private Sub BankInDate_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub

    Private Sub BankInDate_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ConnectDB()
        loadCombo()
    End Sub

    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        btnPrint.Enabled = False
        If cmbAgent.Text = "ALL" Then
            strAgent = " "
        Else
            strAgent = " and Receipt.AgentId=" & SafeSQL(cmbAgent.SelectedValue) & " "
        End If
        Dim strSql As String
        strSql = "SELECT Receipt.RcptNo, SalesAgent.Code + ' / ' + SalesAgent.Name AS AgentId, Customer.CustNo as CustId, Customer.CustName, RcptItem.AmtPaid, " & _
        "Receipt.ChqNo, Receipt.ChqDt, Receipt.BankInDate, Receipt.RcptDt, RcptItem.InvNo, " & _
        "" & SafeSQL(Format(dtpFromDate.Value, "dd/MM/yyyy")) & " as FromDate, " & SafeSQL(Format(dtpToDate.Value, "dd/MM/yyyy")) & " as ToDate " & _
        "FROM RcptItem INNER JOIN Receipt ON RcptItem.RcptNo = Receipt.RcptNo INNER JOIN " & _
        "Customer ON Receipt.CustID = Customer.CustNo INNER JOIN SalesAgent ON Receipt.AgentID = SalesAgent.Code " & _
        "WHERE (Receipt.PayMethod = 'Cheque') and Receipt.RcptDt Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & strAgent
        'strSql = "StockTakeDet.UOM, StockTakeDet.Qty," & SafeSQL(Format(dtpFromDate.Value, "dd/MM/yyyy")) & " as FromDate, " & SafeSQL(Format(dtpToDate.Value, "dd/MM/yyyy")) & " as ToDate from StockTakeHdr, StockTakeDet, SalesAgent, Customer, Item where StockTakeHdr.StockTakeNo = StockTakeDet.StockTakeNo and StockTakeHdr.AgentId = SalesAgent.Code and Item.ItemNo = StockTakeDet.ItemId and StockTakeHdr.CustomerId = Customer.CustNo and StockTakeHdr.StockTakeDt Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & strAgent

        ExecuteReport(strSql, "BankInDateRep")
        btnPrint.Enabled = True
    End Sub
End Class