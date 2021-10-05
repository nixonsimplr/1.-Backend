Imports System.Data
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales

Public Class AllReports
    Implements ISalesBase
    Private aTerms As New ArrayList()
    Private aAgent As New ArrayList()
    Dim sAgentName As String
    Dim sAgentID As String
    Dim sVehicleID As String = " "
    Dim sLocation As String = " "
    Dim ArList As New ArrayList
    Private Structure ArrItemPrice
        Dim ItemCode As String
        Dim Desc As String
        Dim UOM As String
    End Structure
    Dim ArCustList As New ArrayList
    Private Structure ArrCust
        Dim sCustNo As String
        Dim sCustName As String
        Dim sAgent As String
        Dim iNo As Integer
    End Structure

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

    Private Sub AllReports_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        DisconnectDB()
    End Sub

    Private Sub AllReports_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ConnectDB()
        loadCombo()
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

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim dtr As SqlDataReader
        dtr = ReadRecord("SELECT  PreInvNo as Agent, VehicleID,Location FROM  MDT WHERE AgentID=" & SafeSQL(cmbAgent.SelectedValue))
        If dtr.Read = True Then
            sAgentID = dtr("Agent").ToString
            sVehicleID = dtr("VehicleID").ToString
            sLocation = dtr("Location").ToString
        Else
            sAgentID = " "
        End If
        dtr.Close()

        dtr = ReadRecord("SELECT  Name as Agent FROM  SalesAgent where Code=" & SafeSQL(cmbAgent.SelectedValue))
        If dtr.Read = True Then
            sAgentName = sAgentID & " / " & dtr("Agent").ToString
        Else
            sAgentName = "0"
        End If
        dtr.Close()

        PrintDailySales()
        PrintPayCollection()
        PrintStockOrder()
        PrintVanInventory()
        PrintGoodsReturn()
        PrintSchedule()
        PrintVoidInvoice()

    End Sub

    Private Sub PrintDailySales()
        Dim strPay As String = " "
        Dim strTerms As String = " "
        If cmbAgent.Text = "ALL" Then
            strPay = " "
        Else
            strPay = "and Invoice.Agentid=" & "'" & cmbAgent.SelectedValue & "'"
        End If
        Dim strSql As String
        If cmbAgent.Text = "ALL" Then
            strSql = "SELECT DISTINCT Customer.CustNo, Customer.CustName, Invoice.PayTerms AS PaymentMethod, Invoice.InvNo, Invoice.InvDt, Invoice.Discount, Invoice.SubTotal, Invoice.PrintNo AS GST, Invoice.TotalAmt, PayTerms.Description as PayTerms, " & SafeSQL(sAgentID) & " + ' / ' + SalesAgent.Name as AgentID, Invoice.Void FROM Customer INNER JOIN Invoice ON Customer.CustNo = Invoice.CustId INNER JOIN SalesAgent ON Invoice.AgentId = SalesAgent.Code INNER JOIN PayTerms ON Invoice.PayTerms = PayTerms.Code where InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & strTerms '& " Order By Products.ProductType, Products.CategoryID, Products.Brand, Products.Packsize, Products.ShortDesc" '& " Union SELECT DISTINCT NewCust.CustID, NewCust.CustName, 'COD' AS PaymentMethod, Invoice.InvNo, Invoice.InvDt, Invoice.Discount, Invoice.SubTotal, Invoice.GstAmt AS GST, Invoice.TotalAmt, 'COD' as PayTerms, SalesAgent.Name as AgentID , Invoice.VOID FROM NewCust INNER JOIN Invoice ON NewCust.CustID = Invoice.CustId INNER JOIN SalesAgent ON Invoice.AgentId = SalesAgent.Code where InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & strTerms
        Else
            strSql = "SELECT DISTINCT Customer.CustNo, Customer.CustName, Invoice.PayTerms AS PaymentMethod, Invoice.InvNo, Invoice.InvDt, Invoice.Discount, Invoice.SubTotal, Invoice.PrintNo AS GST, Invoice.TotalAmt, PayTerms.Description as PayTerms, " & SafeSQL(sAgentID) & " + ' / ' + SalesAgent.Name as AgentID, Invoice.Void FROM Customer INNER JOIN Invoice ON Customer.CustNo = Invoice.CustId INNER JOIN SalesAgent ON Invoice.AgentId = SalesAgent.Code INNER JOIN PayTerms ON Invoice.PayTerms = PayTerms.Code where InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & strTerms '&'& " Union SELECT DISTINCT NewCust.CustID, NewCust.CustName, 'COD' AS PaymentMethod, Invoice.InvNo, Invoice.InvDt, Invoice.Discount, Invoice.SubTotal, Invoice.GstAmt AS GST, Invoice.TotalAmt, 'COD' as PayTerms, SalesAgent.Name as AgentID , Invoice.VOID FROM NewCust INNER JOIN Invoice ON NewCust.CustID = Invoice.CustId INNER JOIN SalesAgent ON Invoice.AgentId = SalesAgent.Code where InvDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & strTerms
        End If
        PrintAllReport(strSql, "DailySalesRep")
    End Sub

    Private Sub PrintPayCollection()
        Dim strPay1 As String = " "
        Dim strPay As String = " "
        If cmbAgent.Text = "ALL" Then
            strPay1 = " "
        Else
            strPay1 = "and Receipt.AgentID=" & SafeSQL(cmbAgent.SelectedValue)
        End If
        Dim strSql As String = "SELECT Distinct Receipt.RcptNo, Receipt.RcptDt, Customer.CustNo, Customer.CustName, RcptItem.InvNo, RcptItem.AmtPaid as PaidAmt,SalesAgent.Name as AgentID, PayMethod.Description as PayMethod, (Receipt.BankName + ' ' + Receipt.ChqNo + ' ' + convert(varchar(10),Receipt.ChqDt,101)) as ChqNo FROM Receipt, Customer, RcptItem, SalesAgent, PayMethod WHERE (Receipt.void=0 or Receipt.Void is Null) and Receipt.CustId = CustNo and SalesAgent.Code=Receipt.AgentID and PayMethod.code=Receipt.Paymethod and Receipt.RcptNo = RcptItem.RcptNo and Receipt.RcptDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & strPay1 & " Order by Receipt.RcptDt, Receipt.RcptNo, RcptItem.InvNo"
        PrintAllReport(strSql, "PayCollectionRep")
    End Sub

    Private Sub PrintStockOrder()
        Dim strAgent As String = " "
        If cmbAgent.Text = "ALL" Then
            strAgent = " "
        Else
            strAgent = "and StockOrder.Agentid=" & "'" & cmbAgent.SelectedValue & "'"
        End If
        

        Dim strSql As String
        If cmbAgent.Text = "ALL" Then
            strSql = "SELECT Item.Description as ItemName, StockOrder.StockNo, StockOrder.OrdDt as StockDt, StockOrder.Status, 'ALL' as AgentId, StockOrderItem.ItemNo, StockOrderItem.UOM, StockOrderItem.Qty FROM StockOrder, StockOrderItem, Item, SalesAgent Where StockOrder.StockNo = StockOrderItem.StockNo and StockOrderItem.ItemNo = Item.ItemNo and StockOrder.AgentID=SalesAgent.Code and OrdDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strAgent
        Else
            strSql = "SELECT Item.Description as ItemName, StockOrder.StockNo, StockOrder.OrdDt as StockDt, StockOrder.Status, " & SafeSQL(sAgentID) & " + ' / ' + SalesAgent.Name as AgentId, StockOrderItem.ItemNo, StockOrderItem.UOM, StockOrderItem.Qty FROM StockOrder, StockOrderItem, Item, SalesAgent Where StockOrder.StockNo = StockOrderItem.StockNo and StockOrderItem.ItemNo = Item.ItemNo and StockOrder.AgentID=SalesAgent.Code and OrdDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strAgent
        End If

        PrintAllReport(strSql, "StockOrderRep")
    End Sub

    Private Sub PrintVanInventory()
        Dim dSyncDate As Date = CDate("1/1/2000")
        Dim dLastSyncDate As Date = CDate("1/1/2000")
        'Dim iCnt As Integer
        Dim dtr As SqlDataReader
        Try
            'dSyncDate = dgvSync.Item(0, dgvSync.CurrentCell.RowIndex).Value
            'sAgent = dgvSync.Item(1, dgvSync.CurrentCell.RowIndex).Value
            'sLocation = dgvSync.Item(2, dgvSync.CurrentCell.RowIndex).Value
            'For iCnt = dgvSync.CurrentCell.RowIndex + 1 To dgvSync.Rows.Count - 1
            '    If sAgent = dgvSync.Item(1, iCnt).Value Then
            '        dLastSyncDate = dgvSync.Item(0, iCnt).Value
            '        Exit For
            '    End If
            'Next
            dSyncDate = dtpDate.Value
            dLastSyncDate = dtpDate.Value
        Catch ex As Exception
            MsgBox("Please select a date and try again.")
            btnPrint.Enabled = True
            Exit Sub
        End Try


        If dSyncDate = CDate("1/1/2000") Or dLastSyncDate = CDate("1/1/2000") Then
            MsgBox("Please select a valid date and try again.")
            btnPrint.Enabled = True
            Exit Sub
        End If
        'Dim dtr As SqlDataReader
        Dim rs1 As SqlDataReader
        Dim fFoc As Double
        Dim dOut, dDel, dEx, dRt, dRtn As Double
        Dim str As String
        ConnectDB()
        Dim strDel = "Delete from VanDailySales"
        ExecuteSQL(strDel)
        dtr = ReadRecord("SELECT Item.ItemNo, Description, ItemTrans.UOM FROM Item, ItemTrans where Item.ItemNo=ItemTrans.ItemID and DateDiff(s, " & SafeSQL(Format(dLastSyncDate, "yyyyMMdd 23:59:59")) & ", DocDt) >= 0 and DateDiff(s, " & SafeSQL(Format(dSyncDate, "yyyyMMdd 00:00:00")) & ", DocDt) < 0")
        While dtr.Read = True
            Dim aPr1 As ArrItemPrice
            aPr1.ItemCode = dtr("ItemNo").ToString
            aPr1.Desc = dtr("Description").ToString
            aPr1.UOM = dtr("UOM").ToString
            ArList.Add(aPr1)
        End While
        dtr.Close()
        Dim i As Integer
        For i = 0 To ArList.Count - 1
            Dim aPr As ArrItemPrice
            aPr = ArList(i)
            ''''Van Inventory
            str = "Select Sum(Qty) as Out from ItemTrans where LOCATION = " & SafeSQL(sLocation) & " and  DOCTYPE = 'VANINVN' and ItemTrans.ItemId=" & SafeSQL(aPr.ItemCode) & " and DateDiff(s, " & SafeSQL(Format(dLastSyncDate, "yyyyMMdd 23:59:59")) & ", DocDt) >= 0 and DateDiff(s, " & SafeSQL(Format(dSyncDate, "yyyyMMdd 00:00:00")) & ", DocDt) < 0 "
            rs1 = ReadRecord(str)
            If rs1.Read = True Then
                If rs1(0) Is System.DBNull.Value = True Then
                    dOut = 0
                Else
                    dOut = rs1(0)
                End If
            End If
            rs1.Close()
            If dOut = 0 Then Continue For
            Dim insSql As String = "Insert into VanDailySales(DocNo, ItemNo, ItemName, Uom, Out, Delivered, Exchange, Foc, AgentID, Date, GVar, ReturnItem) values (''," & SafeSQL(aPr.ItemCode) & "," & SafeSQL(aPr.Desc) & "," & SafeSQL(aPr.UOM) & "," & dOut & "," & dDel & "," & dEx & "," & fFoc & "," & SafeSQL(sAgentName) & ",'" & Format(dSyncDate, "yyyyMMdd HH:mm:ss") & "'," & dRt & "," & dRtn & ")"
            ExecuteSQL(insSql)
            'dtr.Close()
        Next
        Dim strSql As String = "Select Distinct VanDailySales.*, Item.ProdType, Item.Category, Item.Brand, Item.Packsize, Item.Description  from VanDailySales, Item where VanDailySales.ItemNo=Item.ItemNo order by Item.ProdType, Item.Category, Item.Brand, Item.Packsize, Item.Description"
        PrintAllReport(strSql, "VanInventoryRep")
    End Sub

    Private Sub PrintGoodsReturn()
        Dim strTerms As String = " "
        Dim strRPay As String = " "
        Dim strEPay As String = " "
        Dim strAsset As String = " "
        ExecuteSQL("Delete from Goodsretrep")
        If cmbAgent.Text = "ALL" Then
            strRPay = " "
            strEPay = " "
            strAsset = " "
        Else
            strRPay = "and SalesPersonCode=" & "'" & cmbAgent.SelectedValue & "'"
            strEPay = "and Invoice.AgentID=" & "'" & cmbAgent.SelectedValue & "'"
            strAsset = "and AssetTrans.AgentID=" & "'" & cmbAgent.SelectedValue & "'"
        End If
        Dim dtr As SqlDataReader
        Dim strRSql As String
        Dim strESql As String
        Dim strASql As String
        If cmbAgent.Text = "ALL" Then
            strRSql = "SELECT GoodsReturn.ReturnNo, min(ReturnDate) as ReturnDate, GoodsReturn.CustNo, 'ALL' as SalesPersonCode, isnull(GoodsReturn.CreditNoteNo,'') as CreditNoteNo, GoodsReturnItem.UOM, sum(GoodsReturnItem.Quantity) as Quantity, GoodsReturnItem.ItemNo, 'Return' as CustName, Item.Description FROM  GoodsReturn INNER JOIN GoodsReturnItem ON GoodsReturn.ReturnNo = GoodsReturnItem.ReturnNo INNER JOIN Item ON GoodsReturnItem.ItemNo = Item.ItemNo INNER JOIN Customer ON GoodsReturn.CustNo = Customer.CustNo INNER JOIN SalesAgent ON GoodsReturn.SalesPersonCode = SalesAgent.Code where GoodsReturn.ReturnDate Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strRPay & " GROUP BY GoodsReturn.ReturnNo,GoodsReturn.CustNo,GoodsReturn.CreditNoteNo,item.Description,GoodsReturnItem.ItemNo, GoodsReturnItem.UOM, Item.shortdesc"
            strASql = "SELECT AssetTrans.AssetID as ReturnNo, min(AssetTrans.DocDt) as ReturnDate, AssetTrans.CustNo, SalesAgent.Name as SalesPersonCode, '' as CreditNoteNo, AssetTrans.UOM, sum(AssetTrans.InQty) as Quantity, AssetTrans.AssetID as ItemNo, 'Asset Trans' as CustName, Asset.AssetDesc as Description  FROM  Asset INNER JOIN AssetTrans ON Asset.AssetID = AssetTrans.AssetID INNER JOIN  SalesAgent ON AssetTrans.AgentID = SalesAgent.Code where DocType='O' and DocDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strAsset & " group by AssetTrans.AssetID,AssetTrans.CustNo,SalesAgent.Name,AssetTrans.UOM,Asset.AssetDesc"
            strESql = "SELECT Invoice.InvNo as ReturnNo, min(Invoice.InvDT) as ReturnDate, Invoice.CustID as CustNo, 'ALL' as SalesPersonCode, '' as CreditNoteNo, InvItem.UOM, sum(InvItem.Qty)  as Quantity, InvItem.ItemNo, 'Exchange' as CustName, Item.Description FROM  Invoice INNER JOIN InvItem ON Invoice.InvNo = InvItem.InvNo INNER JOIN Item ON InvItem.ItemNo = Item.ItemNo INNER JOIN Customer ON Invoice.CustID = Customer.CustNo INNER JOIN SalesAgent ON Invoice.AgentID = SalesAgent.Code where InvItem.Description like 'EX:%' and InvDT Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strEPay & " group by Invoice.InvNo , Invoice.CustID,InvItem.UOM,InvItem.ItemNo,Item.Description"
        Else
            strRSql = "SELECT GoodsReturn.ReturnNo, min(GoodsReturn.ReturnDate) as ReturnDate, GoodsReturn.CustNo, " & SafeSQL(sAgentID) & " + ' / ' + SalesAgent.Name as SalesPersonCode, GoodsReturn.CreditNoteNo, GoodsReturnItem.UOM,sum(GoodsReturnItem.Quantity) as Quantity, GoodsReturnItem.ItemNo, 'Return' as CustName, Item.Description FROM  GoodsReturn INNER JOIN GoodsReturnItem ON GoodsReturn.ReturnNo = GoodsReturnItem.ReturnNo INNER JOIN Item ON GoodsReturnItem.ItemNo = Item.ItemNo INNER JOIN Customer ON GoodsReturn.CustNo = Customer.CustNo INNER JOIN SalesAgent ON GoodsReturn.SalesPersonCode = SalesAgent.Code where GoodsReturn.ReturnDate Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strRPay & " GROUP BY SalesAgent.Name,GoodsReturn.ReturnNo,GoodsReturn.CustNo,GoodsReturn.CreditNoteNo,item.Description,GoodsReturnItem.ItemNo, GoodsReturnItem.UOM, Item.shortdesc"
            strASql = "SELECT AssetTrans.AssetID as ReturnNo, min(AssetTrans.DocDt) as ReturnDate, AssetTrans.CustNo, " & SafeSQL(sAgentID) & " + ' / ' + SalesAgent.Name as SalesPersonCode, '' as CreditNoteNo, AssetTrans.UOM,sum(AssetTrans.InQty) as Quantity, AssetTrans.AssetID as ItemNo, 'Asset Trans' as CustName, Asset.AssetDesc as Description FROM  Asset INNER JOIN AssetTrans ON Asset.AssetID = AssetTrans.AssetID INNER JOIN SalesAgent ON AssetTrans.AgentID = SalesAgent.Code where DocType='O' and DocDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strAsset & " group by AssetTrans.AssetID,AssetTrans.CustNo,SalesAgent.Name,AssetTrans.UOM,Asset.AssetDesc"
            strESql = "SELECT Item.ProdType, Item.Category, Item.Brand, Item.Packsize, Item.Description, Invoice.InvNo as ReturnNo,min(Invoice.InvDT) as ReturnDate, Invoice.CustID as CustNo, " & SafeSQL(sAgentID) & " + ' / ' + SalesAgent.Name as SalesPersonCode, '' as CreditNoteNo, InvItem.UOM, sum(InvItem.Qty)  as Quantity, InvItem.ItemNo, 'Exchange' as CustName FROM  Invoice INNER JOIN InvItem ON Invoice.InvNo = InvItem.InvNo INNER JOIN Item ON InvItem.ItemNo = Item.ItemNo INNER JOIN Customer ON Invoice.CustID = Customer.CustNo INNER JOIN SalesAgent ON Invoice.AgentID = SalesAgent.Code where InvItem.Description like 'EX:%' and InvDT Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strEPay & " group by Item.ProdType, Item.Category, Item.Brand, Item.Packsize, Item.Description, Invoice.InvNo , SalesAgent.Name,Invoice.CustID,InvItem.UOM,InvItem.ItemNo"
        End If
        dtr = ReadRecord(strRSql)
        While dtr.Read = True
            ExecuteAnotherSQL("Insert into Goodsretrep values (" & SafeSQL(dtr("ReturnNo")) & "," & SafeSQL(dtr("ReturnDate")) & "," & SafeSQL(dtr("CustNo")) & "," & SafeSQL(dtr("SalesPersonCode")) & "," & SafeSQL(dtr("CreditNoteNo")) & "," & SafeSQL(dtr("UOM")) & "," & dtr("Quantity") & "," & SafeSQL(dtr("ItemNo")) & "," & SafeSQL(dtr("CustName")) & "," & SafeSQL(dtr("Description")) & ")")
        End While
        dtr.Close()
        dtr = ReadRecord(strESql)
        While dtr.Read = True
            ExecuteAnotherSQL("Insert into Goodsretrep values (" & SafeSQL(dtr("ReturnNo")) & "," & SafeSQL(dtr("ReturnDate")) & "," & SafeSQL(dtr("CustNo")) & "," & SafeSQL(dtr("SalesPersonCode")) & "," & SafeSQL(dtr("CreditNoteNo")) & "," & SafeSQL(dtr("UOM")) & "," & dtr("Quantity").ToString & "," & SafeSQL(dtr("ItemNo")) & "," & SafeSQL(dtr("CustName")) & "," & SafeSQL(dtr("Description")) & ")")
        End While
        dtr.Close()
        dtr = ReadRecord(strASql)
        While dtr.Read = True
            ExecuteAnotherSQL("Insert into Goodsretrep values (" & SafeSQL(dtr("ReturnNo")) & "," & SafeSQL(dtr("ReturnDate")) & "," & SafeSQL(dtr("CustNo")) & "," & SafeSQL(dtr("SalesPersonCode")) & "," & SafeSQL(dtr("CreditNoteNo")) & "," & SafeSQL(dtr("UOM")) & "," & dtr("Quantity") & "," & SafeSQL(dtr("ItemNo")) & "," & SafeSQL(dtr("CustName")) & "," & SafeSQL(dtr("Description")) & ")")
        End While
        dtr.Close()
        PrintAllReport("select Item.ItemNo, Item.Description, UOM ,sum(Quantity) as Quantity,SalesPersonCode,min(ReturnDate) as ReturnDate,CustName from Goodsretrep, Item where GoodsRetRep.ItemNo=Item.ItemNo group by Item.ProdType, Item.Category, Item.Brand, Item.Packsize, Item.Description,Item.ItemNo,UOM,CustName,SalesPersonCode ", "GoodsRetRep")
    End Sub

    Private Sub PrintSchedule()
        Dim strPay As String = " "

        If cmbAgent.Text = "ALL" Then
            strPay = " "
        Else
            strPay = "and Customer.SalesAgent=" & "'" & cmbAgent.SelectedValue & "'"
        End If
        Dim dtr As SqlDataReader
        
        ArCustList.Clear()
        ExecuteSQL("Delete from ScheduleRep")
        dtr = ReadRecord("SELECT Customer.CustNo, CustName, " & SafeSQL(sAgentName) & "  as Name, StopNo FROM  RouteMaster INNER JOIN RouteDet ON RouteMaster.RouteNo = RouteDet.RouteNo INNER JOIN Customer ON RouteDet.CustNo = Customer.CustNo WHERE (RouteMaster.RouteDay = " & dtpDate.Value.DayOfWeek & ") and VehicleNo=" & SafeSQL(sVehicleID) & " union Select Distinct CustID as CustNo, CustName, " & SafeSQL(sAgentName) & "  as Name, '9999' as stopno from CustVisit, Customer where CustVisit.CustID=Customer.CustNo  and TransDate Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & " and AgentID=" & SafeSQL(cmbAgent.SelectedValue) & " and CustID not in (SELECT Customer.CustNo FROM  RouteMaster INNER JOIN RouteDet ON RouteMaster.RouteNo = RouteDet.RouteNo INNER JOIN Customer ON RouteDet.CustNo = Customer.CustNo  WHERE RouteMaster.RouteDay = " & dtpDate.Value.DayOfWeek & " and VehicleNo=" & SafeSQL(sVehicleID) & ")")
        While dtr.Read = True
            '      ExecuteAnotherSQL("Insert into ScheduleRep(AgentID, NoofCustActual, NoofCustServed, EffDate) values(" & SafeSQL(dtr("Name").ToString) & "," & dtr("cnt") & ",0," & SafeSQL(Format(dtpDate.Value, "yyyyMMdd HH:mm:ss")) & ")")
            Dim aPr1 As ArrCust
            aPr1.sCustNo = dtr("CustNo").ToString
            aPr1.sCustName = dtr("CustName").ToString
            aPr1.sAgent = dtr("Name").ToString
            aPr1.iNo = dtr("StopNo")
            ArCustList.Add(aPr1)
        End While
        dtr.Close()
        Dim i As Integer
        For i = 0 To ArCustList.Count - 1
            Dim aPr As ArrCust
            aPr = ArCustList(i)
            ExecuteSQL("Insert into ScheduleRep(CustNo, CustName, SalesAgent, Status, VisitedDate, NextVisitDate, StopNo) values(" & SafeSQL(aPr.sCustNo) & "," & SafeSQL(aPr.sCustName) & "," & SafeSQL(aPr.sAgent) & ",''," & SafeSQL(Format(dtpDate.Value, "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(Format(dtpDate.Value, "yyyyMMdd HH:mm:ss")) & "," & aPr.iNo & ")")
            dtr = ReadRecord("Select * from CustVisit where CustID=" & SafeSQL(aPr.sCustNo) & " and TransDate Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")))
            If dtr.Read = True Then
                ExecuteAnotherSQL("Update ScheduleRep Set Status='Visited' where CustNo=" & SafeSQL(aPr.sCustNo))
            Else
                ExecuteAnotherSQL("Update ScheduleRep Set Status='Not Visited' where CustNo=" & SafeSQL(aPr.sCustNo))
            End If
            dtr.Close()
            dtr = ReadRecord("Select * from Service where CustomerID=" & SafeSQL(aPr.sCustNo) & " and ServiceDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")))
            If dtr.Read = True Then
                ExecuteAnotherSQL("Update ScheduleRep Set Status=" & SafeSQL("Visited-" & dtr("Details")) & "  where CustNo=" & SafeSQL(aPr.sCustNo))
            Else
                '       ExecuteAnotherSQL("Update ScheduleRep Set Status='Not Visited' where CustNo=" & SafeSQL(aPr.sCustNo))
            End If
            dtr.Close()
        Next i
        Dim strSql As String = "SELECT * from ScheduleRep order by StopNo" ', CustName, " & SafeSQL(sAgentName) & " + ' / ' + SalesAgent.Name as SalesAgent from DailySchedule, Customer, SalesAgent where DailySchedule.CustNo=Customer.CustNo and Customer.SalesAgent=SalesAgent.Code and Visited=1 and LastVisitedDate Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay
        PrintAllReport(strSql, "ScheduleRep")
    End Sub

    Private Sub PrintVoidInvoice()
        Dim strPay As String = " "
        If cmbAgent.Text = "All" Then
            strPay = " "
        Else
            strPay = " and Agentid=" & SafeSQL(cmbAgent.Text)
        End If
        Dim strSql As String = "Select * from VoidInvoice where Void=1" & strPay
        PrintAllReport(strSql, "VoidInvoiceRep")
    End Sub

    Private Sub PrintAllReport(ByVal strSql As String, ByVal RptName As String)
        Dim DA As New SqlDataAdapter(strSql, My.Settings.ConnectionString)
        Dim DS As New DataSet
        DA.Fill(DS)
        Dim ConInfo As New CrystalDecisions.Shared.TableLogOnInfo
        ConInfo.ConnectionInfo.IntegratedSecurity = True
        'ConInfo.ConnectionInfo.UserID = objDataBase.UserName
        'ConInfo.ConnectionInfo.Password = objDataBase.Password
        ConInfo.ConnectionInfo.ServerName = ".\SQLEXPRESS"
        ConInfo.ConnectionInfo.DatabaseName = "Sales"

        Dim strReportPath As String = Application.StartupPath & "\" & RptName & ".rpt"
        If Not IO.File.Exists(strReportPath) Then
            Throw (New Exception("Unable to locate report file:" & vbCrLf & strReportPath))
        End If
        Dim rptDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        rptDocument.Load(strReportPath)
        rptDocument.SetDataSource(DS.Tables(0))
        rptDocument.PrintOptions.PrinterName = "SHARP MX-2300N PCL6"
        rptDocument.PrintToPrinter(1, False, 0, 0)
        'Dim frm As New ViewReport
        'frm.crvReport.ShowRefreshButton = False
        'frm.crvReport.ShowCloseButton = False
        'frm.crvReport.ShowGroupTreeButton = False
        'frm.crvReport.ReportSource = rptDocument
        ''frm.crvReport.PrintOptions.PrinterName = "\\Epson LQ-300 ESC/P 2"
        ''frm.Show()

        'frm.crvReport.PrintReport()
        'frm.crvReport.Dispose()
    End Sub
End Class