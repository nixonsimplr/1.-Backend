Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Public Class StockOrder
    Implements ISalesBase
    Private aAgent As New ArrayList()
    Dim strAgent As String = " "
    
    Private Sub StockOrder_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
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
    Private Sub StockOrder_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ConnectDB()
        loadCombo()
        'ExecuteSQL("Drop table InvoiceList")
        'ExecuteSQL("SELECT Customer.CustNo, Customer.CustName, Customer.PaymentMethod, Invoice.InvNo, Invoice.InvDt, Invoice.AgentID, Invoice.Discount, Invoice.SubTotal, Invoice.GST, Invoice.TotalAmt, Invoice.PayTerms into InvoiceList FROM Customer INNER JOIN Invoice ON Customer.CustNo = Invoice.CustId")
        ' loadCombo()
        'cmbTerms.SelectedIndex = 0
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

    Private Sub btnPrint_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        btnPrint.Enabled = False
        If cmbAgent.Text = "ALL" Then
            strAgent = " "
        Else
            strAgent = "and StockOrder.Agentid=" & "'" & cmbAgent.SelectedValue & "'"
        End If
        Dim sAgentID As String
        Dim dtr As SqlDataReader
        dtr = ReadRecord("SELECT  PreInvNo as Agent FROM  MDT WHERE AgentID=" & SafeSQL(cmbAgent.SelectedValue))
        If dtr.Read = True Then
            sAgentID = dtr("Agent").ToString
        Else
            sAgentID = "0"
        End If
        dtr.Close()


      
        Dim strSql As String
        If cmbAgent.Text = "ALL" Then
            strSql = "SELECT Item.Description as ItemName, StockOrder.StockNo, StockOrder.OrdDt as StockDt, StockOrder.Status, 'ALL' as AgentId, StockOrderItem.ItemNo, StockOrderItem.UOM, StockOrderItem.Qty FROM StockOrder, StockOrderItem, Item, SalesAgent Where StockOrder.StockNo = StockOrderItem.StockNo and StockOrderItem.ItemNo = Item.ItemNo and StockOrder.AgentID=SalesAgent.Code and OrdDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strAgent
        Else
            strSql = "SELECT Item.Description as ItemName, StockOrder.StockNo, StockOrder.OrdDt as StockDt, StockOrder.Status, " & SafeSQL(sAgentID) & " + ' / ' + SalesAgent.Name as AgentId, StockOrderItem.ItemNo, StockOrderItem.UOM, StockOrderItem.Qty FROM StockOrder, StockOrderItem, Item, SalesAgent Where StockOrder.StockNo = StockOrderItem.StockNo and StockOrderItem.ItemNo = Item.ItemNo and StockOrder.AgentID=SalesAgent.Code and OrdDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strAgent
        End If

        ExecuteReport(strSql, "StockOrderRep")
        btnPrint.Enabled = True
    End Sub
End Class