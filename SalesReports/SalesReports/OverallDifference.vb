Imports System.Data.SqlClient
Imports SalesInterface.MobileSales

Public Class OverallDifference
    Implements ISalesBase
    Dim strLoc As String = " "
    Dim ArList As New ArrayList

    Private Structure ArrItemPrice
        Dim ItemCode As String
        Dim Desc As String
        Dim cPrice As Double
        Dim sUOM As String
    End Structure

    Private Sub Inventory_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub

    Private Sub Inventory_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ConnectDB()
        dtpFromDate.Value = Date.Now.AddDays(-Date.Now.Day + 1)
        'ExecuteSQL("Drop table InvoiceList")
        'ExecuteSQL("SELECT Customer.CustNo, Customer.CustName, Customer.PaymentMethod, Invoice.InvNo, Invoice.InvDt, Invoice.AgentID, Invoice.Discount, Invoice.SubTotal, Invoice.GST, Invoice.TotalAmt, Invoice.PayTerms into InvoiceList FROM Customer INNER JOIN Invoice ON Customer.CustNo = Invoice.CustId")
        loadLocation()
    End Sub

    Public Sub loadLocation()
        Dim aLoc As New ArrayList()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select distinct Location from customer where Location <> '' and Consignment = 1 order by [Location]")
        cmbLocation.DataSource = Nothing
        aLoc.Clear()
        aLoc.Add(New ComboValues("ALL", "ALL"))
        While dtr.Read()
            aLoc.Add(New ComboValues(dtr("Location").ToString, dtr("Location").ToString))
        End While
        dtr.Close()
        cmbLocation.DataSource = aLoc
        cmbLocation.DisplayMember = "Desc"
        cmbLocation.ValueMember = "Code"
        cmbLocation.SelectedIndex = 0
    End Sub

    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim str As String = ""

        If cmbLocation.Text = "ALL" Then
            strLoc = " and LocationCode in (Select distinct Location from customer where Location <> '' and Consignment = 1)"
        Else
            strLoc = " and LocationCode = " & "'" & cmbLocation.SelectedValue & "'"
        End If

        Dim strInvnLoc As String = ""
        If cmbLocation.Text = "ALL" Then
            strInvnLoc = " and Location in (Select distinct Location from customer where Location <> '' and Consignment = 1)"
        Else
            strInvnLoc = " and Location = " & "'" & cmbLocation.SelectedValue & "'"
        End If

        ExecuteSQL("Delete from InventoryRep")

        ExecuteSQL("Insert InventoryRep(ItemID, ItemName, InvnQty, StockTakeQty, OrdQty, FromDate, ToDate, Location) Select GoodsInvn.ItemNo, ItemName, Qty, 0, 0, " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & ", " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & ", Location from GoodsInvn, Item where GoodsInvn.ItemNo = Item.ItemNo " & strInvnLoc & " group by Location, GoodsInvn.ItemNo, ItemName, Qty")

        ExecuteSQL("Insert InventoryRep(ItemID, ItemName, InvnQty, StockTakeQty, OrdQty, FromDate, ToDate, Location) Select distinct StockTakeDet.ItemId, Item.ItemName, 0, sum(StockTakeDet.Qty), 0, " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & ", " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & ", StockTakeHdr.LocationCode from StockTakeDet, StockTakeHdr, Item where StockTakeHdr.StockTakeNo = StockTakeDet.StockTakeNo and Item.ItemNo = StockTakeDet.ItemID and StockTakeDet.ItemId + StockTakeHdr.LocationCode not in (Select distinct ItemNo + Location from GoodsInvn) and StockTakeHdr.StockTakeDate Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & " " & strLoc & " group by LocationCode, StockTakeDet.ItemId, Item.ItemName ")

        If cmbLocation.Text = "ALL" Then
            ExecuteSQL("Update InventoryRep set StockTakeQty = Qty from (Select LocationCode, ItemId, sum(Qty) * BaseQty as Qty from StockTakeHdr, StockTakeDet, " & _
            "Uom where StockTakeHdr.StockTakeNo = StockTakeDet.StockTakeNo and StockTakeDet.Uom = Uom.Uom and StockTakeHdr.StockTakeDate between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & " and Uom.ItemNo = " & _
            "StockTakeDet.ItemId Group by LocationCode, ItemId, BaseQty) A where A.ItemId = InventoryRep.ItemID and A.LocationCode = InventoryRep.Location")
        Else
            ExecuteSQL("Update InventoryRep set StockTakeQty = Qty from (Select ItemId, sum(Qty) * BaseQty as Qty from StockTakeHdr, StockTakeDet, " & _
            "Uom where StockTakeHdr.StockTakeNo = StockTakeDet.StockTakeNo and StockTakeDet.Uom = Uom.Uom and StockTakeHdr.StockTakeDate between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & " and Uom.ItemNo = " & _
            "StockTakeDet.ItemId and StockTakeHdr.LocationCode = " & SafeSQL(cmbLocation.SelectedValue) & " Group by ItemId, BaseQty) A where A.ItemId = InventoryRep.ItemID")
        End If

        ExecuteSQL("Delete from InventoryRep where InvnQty - StockTakeQty = 0")

        Dim strSql As String = "SELECT * from InventoryRep order by ItemID"
        ExecuteReport(strSql, "OverallDifferenceRep")
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