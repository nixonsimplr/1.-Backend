Imports DataTransfer.UpDown
Imports System.Data.SqlClient
Imports DataInterface.IbizDO
Imports SalesInterface.MobileSales
Imports System.IO

Public Class PPCSync
    Implements ISalesBase


    Private Structure DelCust
        Dim CustID As String
        Dim PrGroup As String
    End Structure

    Private ImgCount As Integer = -1
    Private arrImages() As String = {}

    Private arrOrdNo As New ArrayList()
    Private IsChangeUser As Boolean = False
    Private dtup As New DataTransfer.UpDown
    Private objDo As New DataInterface.IbizDO
    Private sConStr As String = My.Settings.SalesConnectionString
    ' Private sConStr As String = "Data Source=IBIZCS-JAGADISH\SQLEXPRESS;Initial Catalog=Sales;Integrated Security=True" 'My.Settings.ConnectionString
    Private sHistoryunit, sAgentID, sLocation, sDefPrGroup As String
    Private iHistoryNo As Integer
    Private arrDet As ArrayList
    Private dGST As Double
    Private sAgt As String = ""
    Dim sMDTNo As String = ""
    Dim sVehicle As String = ""
    Dim dInvHistory As Double = 0

    Private Sub btnUpload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpload.Click
        'GetPrice()
        Upload()
        MsgBox("Completed")
    End Sub

    Private Function GetPictureFolder()
        Dim ds As New DataSet
        Dim dataDirectory As String
        'Instantiate the collection for configuration info.
        'If AppDomain.CurrentDomain.DomainManager IsNot Nothing AndAlso AppDomain.CurrentDomain.DomainManager.ToString().Contains("VSHost") Then
        '    dataDirectory = Windows.Forms.Application.StartupPath
        'Else
        '    dataDirectory = Windows.Forms.Application.UserAppDataPath
        'End If
        dataDirectory = Windows.Forms.Application.StartupPath
        ds.ReadXml(dataDirectory & "\Simplr.xml")
        Dim table As DataTable
        For Each table In ds.Tables
            Dim row As DataRow
            If table.TableName = "PictureFolder" Then
                For Each row In table.Rows
                    Return row("value").ToString()
                Next row
            End If
        Next table
        Return ""
    End Function

    Public Function GetPrice(ByVal sItemId As String, ByVal sCustNo As String, ByVal sPrGroup As String) As Double
        objDo.ConnectAnotherDB()
        Dim dtr As SqlDataReader
        Dim dPr As Double = 0
        dtr = objDo.ReadRecordAnother("Select UnitPrice from ItemPr where ItemNo = " & objDo.SafeSQL(sItemId) & " and PriceGroup = " & objDo.SafeSQL(sCustNo) & "and SalesType = 'Customer' Order by MinQty")
        If dtr.Read Then
            dPr = dtr("UnitPrice")
        End If
        dtr.Close()
        If dPr > 0 Then GoTo ExitNow

        dtr = objDo.ReadRecordAnother("Select UnitPrice from ItemPr where ItemNo = " & objDo.SafeSQL(sItemId) & " and PriceGroup = " & objDo.SafeSQL(sPrGroup) & " and SalesType = 'Customer Price Group' Order by MinQty")
        If dtr.Read Then
            dPr = dtr("UnitPrice")
        End If
        dtr.Close()
        If dPr > 0 Then GoTo ExitNow

        dtr = objDo.ReadRecordAnother("Select UnitPrice from ItemPr where ItemNo = " & objDo.SafeSQL(sItemId) & " and SalesType = 'All Customers' Order by MinQty")
        If dtr.Read Then
            dPr = dtr("UnitPrice")
        End If
        dtr.Close()
        If dPr > 0 Then GoTo ExitNow

        dtr = objDo.ReadRecordAnother("Select UnitPrice from Item where ItemNo = " & objDo.SafeSQL(sItemId))
        If dtr.Read Then
            dPr = dtr("UnitPrice")
        End If
        dtr.Close()
ExitNow:
        objDo.DisconnectAnotherDB()
        Return dPr
    End Function

    Public Function Upload() As Boolean
        sLocation = ""
        Dim sBal As String
        Dim bStatus As Boolean = True
        If dtup.Connect(sConStr) = False Then
            MsgBox("Please connect device and try again.")
            dtup.Disconnect()
            Return False
        End If
        objDo.ConnectDB()
        Dim arr As New ArrayList
        Dim frm As New status
        frm.Show()
        frm.Refresh()
        dtup.StartUpload()
        'New Code for Calculating Balance
        '''''
        Dim dtr As SqlDataReader
        dtr = objDo.ReadRecord("SELECT Custid, isnull(SUM(TotalAmt) - Sum(PaidAmt),0) AS Balance FROM  Invoice GROUP BY CustId")
        frm.Refresh()
        While dtr.Read = True
            sBal = "Update Customer set Balance = " & dtr("Balance") & " where CustNo = " & objDo.SafeSQL(dtr("CustId").ToString)
            objDo.ExecuteSQLAnother(sBal)
        End While
        dtr.Close()
        frm.Refresh()
        'objDo.DisconnectDB()
        'Dim dtr As SqlDataReader
        sDefPrGroup = ""
        dtr = objDo.ReadRecord("Select HistoryUnit, Historyno, DefPriceGroup from system")
        If dtr.Read = True Then
            sHistoryunit = dtr("HistoryUnit")
            iHistoryNo = dtr("HistoryNo")
            sDefPrGroup = dtr("DefPriceGroup").ToString
        End If
        dtr.Close()
        If sDefPrGroup <> "" Then
            objDo.ExecuteSQL("Update Customer Set PriceGroup=" & objDo.SafeSQL(sDefPrGroup))
        End If
        dtup.StartUpload()
        Try
            frm.Refresh()
            Dim sMDT As String = ""
            If dtup.GetMDTFile = False Or IsChangeUser = True Then
NewMDT:
                'frm.Visible = False
                frm.Refresh()
                If MsgBox("Do you want to upload data for MDT " & dgvSync.Item(0, dgvSync.CurrentCell.RowIndex).Value, MsgBoxStyle.YesNo, "Download Data?") = MsgBoxResult.Yes Then
                    sMDT = dgvSync.Item(0, dgvSync.CurrentCell.RowIndex).Value
                    Dim dtrMDT As SqlDataReader
                    dtrMDT = objDo.ReadRecord("Select AgentID, Location, VehicleId, InvHistory from MDT where MDTNo=" & objDo.SafeSQL(sMDT))
                    If dtrMDT.Read = True Then
                        sAgentID = dtrMDT("AgentID")
                        sLocation = dtrMDT("Location")
                        sVehicle = dtrMDT("VehicleId")
                        If IsDBNull(dtrMDT("InvHistory")) = False Then dInvHistory = dtrMDT("InvHistory")
                    End If
                    dtrMDT.Close()
                    frm.Visible = True
                    frm.Refresh()
                Else
                    dtup.EndUpload()
                    dtup.Disconnect()
                    objDo.DisconnectDB()
                    frm.Close()
                    Return False
                End If
            Else
                Dim path As String
                path = System.AppDomain.CurrentDomain.BaseDirectory & "download\"
                sMDT = GetMDTName(path & "System.xml", path & "system.xsd")
                If sMDT = "" Then
                    GoTo NewMDT
                End If
                Dim dtrMDT As SqlDataReader
                dtrMDT = objDo.ReadRecord("Select AgentID, Location, VehicleId, InvHistory from MDT where MDTNo=" & objDo.SafeSQL(sMDT))
                If dtrMDT.Read = True Then
                    sAgentID = dtrMDT("AgentID")
                    sLocation = dtrMDT("Location")
                    sVehicle = dtrMDT("VehicleId")
                    If IsDBNull(dtrMDT("InvHistory")) = False Then dInvHistory = dtrMDT("InvHistory")
                End If
                dtrMDT.Close()
            End If

            'dtr = objDo.ReadRecord("SELECT CustProd.Custid, ItemNo, PriceGroup FROM CustProd, Customer where CustProd.custID=Customer.Custno and SalesAgent=" & objDo.SafeSQL(sAgentID))
            'While dtr.Read = True
            '    Dim str(2) As String
            '    str(0) = dtr("CustId")
            '    str(1) = dtr("ItemNo")
            '    str(2) = dtr("PriceGroup")
            '    arr.Add(str)
            'End While
            'dtr.Close()

            dtr = objDo.ReadRecord("SELECT isNull(Sum(SubTotal),0) as TotalAmt from Invoice where (void is null or void = 0) and Month(InvDt)=  " & Date.Now.Month & " and Year(InvDt) = " & Date.Now.Year & " and AgentID=" & objDo.SafeSQL(sAgentID))
            If dtr.Read = True Then
                objDo.ExecuteSQLAnother("Update SalesTargetByMonth Set Actual=" & dtr("TotalAmt") & " Where Month=" & objDo.SafeSQL(Format(Date.Now, "MMMM").ToString) & " and AgentID=" & objDo.SafeSQL(sAgentID))
            End If
            dtr.Close()


            objDo.ExecuteSQL("INSERT into SyncHistory(SyncTime, AgentId, Location) values (" & objDo.SafeSQL(Format(DateTime.Now, "yyyyMMdd HH:mm:ss")) & "," & objDo.SafeSQL(sAgentID) & "," & objDo.SafeSQL(sLocation) & ")")

            ''Customer Category

            dtup.UploadFile("SELECT Distinct rtrim(Customer.CustNo) as CustomerID, CustName as Customername,  ContactPerson as contact, Address, Phone, Email, Address2, Address3, Address4, PostCode as Pin, City, CountryCode as Country, Balance, CreditLimit, '' as DefCurCode, '' as ChineseName, PriceGroup, ZoneCode as ZoneId, SalesAgent as AgentNo, PaymentTerms as DefPayTerms, PaymentMethod as DefPayMethod,DisplayNo, GSTType, SearchName, [Print] as PrintDoInv, Remarks, SupplierCode,ShipName as BilltoName, ShipAddr as BillAdd1, ShipAddr2 as BillAdd2, ShipAddr3 as BillAdd3, ShipPost as BillPin, ShipAddr4 as BillAdd4, Shipcity as BillCity, AcCustCode, [Bill-toNo] as AcBillRef, Location, CASE BillMultiple WHEN 1 Then 1 Else 0 END as BillMultiple, CompanyName, Category, '' as Blocked, Case when CustType is null then '' Else CustType End as CustType FROM Customer, CustAgent where SalesAgent = CustAgentID and CustAgent.AgentID=" & objDo.SafeSQL(sAgentID) & " and Active =1 and ToPDA=1", "Customers")
            dtup.UploadFile("SELECT Distinct CustomerBill.CustName as BillName, CustomerBill.Address AS Address1, CustomerBill.Address2, CustomerBill.Address3, CustomerBill.PostCode as Pin, CustomerBill.Address4, CustomerBill.City, Customer.AcCustCode, CustomerBill.AcBillRef FROM CustomerBill, Customer where Customer.AcCustCode=CustomerBill.CustNo and BillMultiple=1 and Customer.ToPDA=1", "BillRef")
            dtup.UploadFile("SELECT RouteNo, RouteName, RouteDay FROM RouteMaster where  VehicleNo = " & objDo.SafeSQL(sVehicle), "RouteMaster")
            dtup.UploadFile("SELECT RouteDet.RouteNo, CustNo as CustomerID, StopNo, CompanyName FROM RouteDet, RouteMaster where RouteMaster.RouteNo = RouteDet.RouteNo and VehicleNo = " & objDo.SafeSQL(sVehicle), "RouteDet")
            dtup.UploadFile("SELECT Code as CategoryID , Description  FROM Category ", "Category")
            dtup.UploadFile("SELECT Code as AgentNo, [Name] as AgentName, UserId, Password, Access as AccessLevel, SalesTarget, CurMonTarget, CurCreditNote, CurPDCChqCollection, CurChqCollection, CurCashCollection, LastInvDate as LastInvDt FROM SalesAgent where Len(UserID) < 20 and (UserID<>'' or UserID is Null)", "Agent")
            dtup.UploadFile("SELECT Code as PayModeCode, Description as PayModeDesc FROM PayMethod where Active=1 ", "PayMode")
            dtup.UploadFile("SELECT Code as TermCode, Description as TermDesc, DueDateCalc as TermCalc,DisDateCalc as DueDateCalc, DiscountPercent FROM PayTerms where Active=1", "payterms")
            dtup.UploadFile("SELECT UOM.ItemNo as ItemId, UOM.Uom, UOM.BaseQty FROM UOM, Item where Item.ItemNo=Uom.ItemNo and Item.ToPDA=1 and Item.Active=1", "UOM")
            dtup.UploadFile("SELECT ItemNo as ItemId, PackageUom, BaseQty FROM Package", "Package")
            dtup.UploadFile("SELECT Code, MinAmount, DiscountPercent FROM CustInvDiscount", "CustInvDiscount")
            dtup.UploadFile("SELECT Code, Description from Brand", "Brand")

            dtup.UploadFile("SELECT ItemPr.PriceGroup, ItemPr.ItemNo as Itemid, ItemPr.SalesType, ItemPr.MinQty, ItemPr.UnitPrice, ItemPr.UOM, ItemPr.FromDate, ItemPr.ToDate FROM Item, ItemPr, Customer Where ItemPr.PriceGroup = Customer.CustNo and Item.ItemNo=ItemPr.ItemNo  and Item.Active=1 and Item.ToPDA=1 and SalesAgent in (Select CustAgentID from CustAgent where AgentID=" & objDo.SafeSQL(sAgentID) & ") and Customer.Active =1 and Customer.ToPDA=1 and SalesType = 'Customer' union SELECT PriceGroup, ItemPr.ItemNo as Itemid, ItemPr.SalesType, ItemPr.MinQty, ItemPr.UnitPrice, ItemPr.UOM, ItemPr.FromDate, ItemPr.ToDate FROM ItemPr, Item Where Item.ItemNo=ItemPr.ItemNo and Item.Active=1 and Item.ToPDA=1 and PriceGroup in (Select Distinct PriceGroup FROM Customer where Customer.TopDA=1 and SalesAgent in (Select CustAgentID from CustAgent where AgentID=" & objDo.SafeSQL(sAgentID) & ") and Customer.Active =1) and SalesType = 'Customer Price Group'", "ItemPr")

            dtup.UploadFile("SELECT Code , Description  FROM Zone where code in (Select distinct ZoneCode from Customer where ToPDA=1 and Customer.Active=1 and SalesAgent=" & objDo.SafeSQL(sAgentID) & ")", "Zone")
            dtup.UploadFile("SELECT CustProd.CustID, CustProd.ItemNo as Itemid, CustProd.ItemName, CustProd.UOM, CustProd.Price, CustProd.CustProdCode FROM CustProd, Customer, Item Where CustProd.CustId = Customer.CustNo and Item.ItemNo=CustProd.ItemNo and Item.Active=1 and Item.ToPDA=1 and Customer.Active = 1 and Customer.ToPDA=1 and SalesAgent in (Select CustAgentID from CustAgent where AgentID=" & objDo.SafeSQL(sAgentID) & ")", "CustProd")
            dtup.UploadFile("SELECT ItemNo as itemid, Description as itemname, ItemName as ChineseName, Description as ShortDesc, BaseUOM as uom, UnitPrice as Price, 1 as Available, Category as CategoryId, Description as fulldesc, 0 as IncGst, DisplayNo,BarCode, Brand, CASE ToPDA WHEN 1 Then 1 Else 0 END as Hidden, CompanyName FROM Item where Active=1 and ToPDA=1", "Products")

            objDo.ExecuteSQL("Delete from GoodsInvn where Qty = 0 and BadQty = 0 and Location <> " & objDo.SafeSQL(sLocation))

            dtup.UploadFile("Select Distinct Item.ItemNo as ItemId, GoodsInvn.Uom, GoodsInvn.Qty, 0 as BadQty, GoodsInvn.Location from GoodsInvn, Item where GoodsInvn.ItemNo = Item.ItemNo and Item.Active=1 and Item.ToPDA = 1 and (GoodsInvn.Location = " & objDo.SafeSQL(sLocation) & " or Location in (Select distinct Location from Customer where SalesAgent = " & objDo.SafeSQL(sAgentID) & " and Consignment = 1)) union SELECT Distinct Item.ItemNo, Item.BaseUOM, 0 as Qty, 0 as BadQty, " & objDo.SafeSQL(sLocation) & " as Location FROM Item Where Item.Active=1 and Item.ToPDA=1 and (Item.ItemNo NOT IN (SELECT GoodsInvn.ItemNo FROM GoodsInvn WHERE Location =" & objDo.SafeSQL(sLocation) & "))", "VanInventory")

            Dim sITDocNo As String = GetNewTransNo()
            dtup.UploadFile("Select " & objDo.SafeSQL(sITDocNo) & "  as DocNo,'" & Format(DateTime.Now, "yyyyMMdd HH:mm:ss") & "' as DocDate, 'VANINVN' as DocType, GoodsInvn.ItemNo as itemid, GoodsInvn.uom, GoodsInvn.qty, GoodsInvn.Location from GoodsInvn, Item where GoodsInvn.ItemNo = Item.ItemNo and Item.Active=1 and Item.ToPDA = 1 and GoodsInvn.Location=" & objDo.SafeSQL(sLocation), "ItemTrans")
            dtup.UploadFile("SELECT Code, Description, CASE InvoiceDiscount WHEN 1 Then 1 Else 0 END as InvoiceDiscount, CASE LineDiscount WHEN 1 Then 1 Else 0 END as LineDiscount FROM PriceGroup", "PriceGroup")
            dtup.UploadFile("SELECT InvNo, InvDt, OrdNo as OrderNo, Discount, SubTotal as netamt, GstAmt as gst, TotalAmt as amount,  PayTerms as Terms, TermDays, '' as CurCode, 1 as CurExRate, PaidAmt, CustId as CustomerID, Invoice.DTG, AgentID, PrintNo, 0 as Void, Invoice.AcBillRef, 1 as Uploaded, 1 as UploadedToPDA, Invoice.CompanyName FROM Invoice, Customer Where (Void <>1 or Void is Null) and ((PaidAmt < TotalAmt) or (DateDiff(d, InvDt, " & objDo.SafeSQL(Format(DateAdd(DateInterval.Month, (-1) * iHistoryNo, Date.Now.Date), "yyyyMMdd HH:mm:ss")) & ") <= 0)) and  Invoice.CustID = Customer.custno and Customer.SalesAgent in (Select CustAgentID from CustAgent where AgentID=" & objDo.SafeSQL(sAgentID) & ")", "Invoices")
            'Arkar 20090507 Change ShortDesc to description
            dtup.UploadFile("SELECT InvItem.InvNo, InvItem.ItemNo AS itemid, InvItem.UOM, InvItem.Qty, InvItem.Price, InvItem.SubAmt AS amount, InvItem.Description as ItemName FROM InvItem, Item, Invoice, Customer where (Void <>1 or Void is Null) and ((PaidAmt < TotalAmt) or (DateDiff(d, InvDt, " & objDo.SafeSQL(Format(DateAdd(DateInterval.Month, (-1) * iHistoryNo, Date.Now.Date), "yyyyMMdd HH:mm:ss")) & ") <= 0)) and Invoice.CustID = Customer.custno and Invoice.InvNo = InvItem.InvNo and InvItem.ItemNo = Item.ItemNo and Customer.SalesAgent in (Select CustAgentID from CustAgent where AgentID=" & objDo.SafeSQL(sAgentID) & ")", "Invdet")
            'dtup.UploadFile("SELECT InvItem.InvNo, InvItem.ItemNo AS itemid, InvItem.UOM, InvItem.Qty, InvItem.Price, InvItem.SubAmt AS amount, Item.ShortDesc as ItemName FROM InvItem, Item, Invoice, Customer where (Void <>1 or Void is Null) and (DateDiff(d, Invoice.InvDt, " & objDo.SafeSQL(Format(DateAdd(DateInterval.Day, (-1) * dInvHistory, Date.Now.Date), "yyyyMMdd HH:mm:ss")) & ") <= 0)  and Invoice.CustID = Customer.custno and Invoice.InvNo = InvItem.InvNo and InvItem.ItemNo = Item.ItemNo and Customer.SalesAgent in (Select CustAgentID from CustAgent where AgentID=" & objDo.SafeSQL(sAgentID) & ")", "Invdet")

            'dtup.UploadFile("Select BinID, Description, ItemID, Location, DTG from Bin", "Bin")
            'dtup.UploadFile("Select WarehouseId as LocationID, ItemNo, Uom, Qty, ShelfID, AreaID, BinID from ConsignmentQty", "ConsignmentQty")

            dtup.UploadFile("SELECT MDT.MDTNo AS PDAID,  MDT.LastOrdNo as LastOrdNo, MDT.PreOrdNo as PreOrdNo, MDT.LenOrdNo as LenordNo, MDT.LastInvNo as LastInvNo, MDT.PreInvNo as PreInvNo, MDT.LenInvNo as LenInvNo, MDT.LastRcptNo as LastRcptNo, MDT.PreRcptNo as PreRcptNo, MDT.LenRcptNo as LenRcptNo, System.GST as GST, '' as ServerIP, 0 as  ServerPort, 'N' as Online, MDT.Printer as Printer,MDT.PrintPort as PrintPort, MDT.PrintBaud as PrintBaud,MDT.dotPrintPort as dotPrintPort, MDT.dotPrintBaud as dotPrintBaud, MDT.CompanyName as CompanyName, System.Header1 as Header1, System.Header2 as Header2, System.Header3 as Header3, System.Header4 as Header4, System.Tail1 as Tail1, System.Tail2 as Tail2, System.Tail3 as Tail3,  System.Tail4 as Tail4, MDT.LastRtnNo as LastRetNo, MDT.PreRtnNo as PreRetNo, MDT.LenRtnNo as LenRetNo, MDT.LastExNo as LastExNo, MDT.PreExNo as PreExNo,MDT.Location as Location, MDT.LenExNo as LenExNo, MDT.AgentID as AgentID, MDT.LastITNo as LastITNo, MDT.PreITNo as PreITNo, MDT.LenITNo as LenITNo, MDT.PreSerNo, MDT.LenSerNo, MDT.LastSerNo, MDT.PreCRNo as PreCRNo, MDT.LenCRNo as LenCRNo, MDT.LastCRNo as LastCRNo, System.TradeDealPriceGroup, '" & Format(Date.Now, "yyyy-MM-dd") & "' as sysdate, CashCustNo as CashCustId, LastCustNo, PreCustNo, LenCustNo, ExchangePer, DefCategory, CASE IsNewMDT WHEN 1 Then 1 Else 0 END as IsNewMDT FROM MDT, System where MDT.MDTNO =" & objDo.SafeSQL(sMDT), "System")

            Dim sStr As String = ""
            Dim Cnt As Integer
            Dim arrItemCon As New ArrayList
            objDo.ExecuteSQL("Delete from PromoConditionTemp")
            objDo.ExecuteSQL("Delete from PromoOfferTemp")
            dtr = objDo.ReadRecord("SELECT PromoCondition.*, Promotion.AllItems from PromoCondition, Promotion where PromoCondition.PromoID=Promotion.PromoID  and ToDate >= '" & Format(Date.Now, "yyyyMMdd") & "' and FromDate <= '" & Format(DateAdd(DateInterval.Day, 7, Date.Now), "yyyyMMdd") & "'")
            frm.Refresh()
            Cnt = 0

            While dtr.Read = True
                Dim bolAllItems As Boolean = False
                If Not IsDBNull(dtr("AllItems")) Then bolAllItems = dtr("AllItems")
                If bolAllItems = True Then
                    Cnt += 1
                    Dim str(1) As String
                    str(0) = dtr("PromoID").ToString
                    str(1) = dtr("ItemID").ToString
                    arrItemCon.Add(str)
                    sStr = "Insert into PromoConditionTemp (ParentPromoID, PromoID, ItemID, UOM, Minqty, Maxqty,LineType,IsRequired) Values (" & objDo.SafeSQL(dtr("PromoID").ToString) & "," & objDo.SafeSQL(dtr("PromoID").ToString & "-" & Cnt) & "," & objDo.SafeSQL(dtr("ItemID").ToString) & " , " & objDo.SafeSQL(dtr("UOM").ToString) & "," & dtr("MinQty") & "," & dtr("MaxQty") & "," & objDo.SafeSQL(dtr("LineType").ToString) & "," & objDo.SafeSQL(dtr("ISRequired").ToString) & ")"
                    objDo.ExecuteSQLAnother(sStr)
                Else
                    sStr = "Insert into PromoConditionTemp (ParentPromoID, PromoID, ItemID, UOM, Minqty, Maxqty,LineType,IsRequired) Values (" & objDo.SafeSQL(dtr("PromoID").ToString) & "," & objDo.SafeSQL(dtr("PromoID").ToString) & "," & objDo.SafeSQL(dtr("ItemID").ToString) & " , " & objDo.SafeSQL(dtr("UOM").ToString) & "," & dtr("MinQty") & "," & dtr("MaxQty") & "," & objDo.SafeSQL(dtr("LineType").ToString) & "," & objDo.SafeSQL(dtr("ISRequired").ToString) & ")"
                    objDo.ExecuteSQLAnother(sStr)
                End If
            End While

            dtr.Close()
            Dim arrItemOff As New ArrayList
            Dim iC As Integer
            dtr = objDo.ReadRecord("SELECT PromoOffer.*, AllItems from PromoOffer, Promotion where PromoOffer.PromoID=Promotion.PromoID  and ToDate >= '" & Format(Date.Now, "yyyyMMdd") & "' and FromDate <= '" & Format(DateAdd(DateInterval.Day, 7, Date.Now), "yyyyMMdd") & "'")
            frm.Refresh()
            While dtr.Read = True
                Dim bolAllItems As Boolean = False
                If Not IsDBNull(dtr("AllItems")) Then bolAllItems = dtr("AllItems")
                If bolAllItems = True Then
                    For iC = 0 To arrItemCon.Count - 1
                        Dim str(1) As String
                        str = arrItemCon(iC)
                        If dtr("PromoID").ToString = str(0) And dtr("ItemID").ToString = str(1) Then
                            sStr = "Insert into PromoofferTemp(ParentPromoID, PromoID, ItemID, UOM, FOCQty, DisPrice,Discount) Values (" & objDo.SafeSQL(dtr("PromoID").ToString) & "," & objDo.SafeSQL(dtr("PromoID").ToString & "-" & iC + 1) & "," & objDo.SafeSQL(dtr("ItemID").ToString) & " , " & objDo.SafeSQL(dtr("UOM").ToString) & "," & dtr("FOCQty") & "," & dtr("DisPrice") & "," & dtr("Discount") & ")"
                            objDo.ExecuteSQLAnother(sStr)
                        End If
                    Next
                Else
                    sStr = "Insert into PromoofferTemp(ParentPromoID, PromoID, ItemID, UOM, FOCQty, DisPrice,Discount) Values (" & objDo.SafeSQL(dtr("PromoID").ToString) & "," & objDo.SafeSQL(dtr("PromoID").ToString) & "," & objDo.SafeSQL(dtr("ItemID").ToString) & " , " & objDo.SafeSQL(dtr("UOM").ToString) & "," & dtr("FOCQty") & "," & dtr("DisPrice") & "," & dtr("Discount") & ")"
                    objDo.ExecuteSQLAnother(sStr)
                End If
            End While
            dtr.Close()

            ''''''''''''Before Changing to all Items Promotion 04022008
            'dtup.UploadFile("SELECT PromoID, PromoName, ApType, FromDate, ToDate, PromoType, Multiply, Priority, MinAmt, MaxAmt, DisAmt, DisPer, Entitle, EntitleType, CASE CATBased WHEN 1 Then 1 Else 0 END as CATBased, CASE ItemCondition WHEN 1 Then 1 Else 0 END as ItemCondition from Promotion Where ToDate >= '" & Format(Date.Now, "yyyyMMdd") & "' and FromDate <= '" & Format(DateAdd(DateInterval.Day, 7, Date.Now), "yyyyMMdd") & "'", "Promotion")
            'dtup.UploadFile("SELECT PromoApply.*  from PromoApply, Promotion where PromoApply.PromoID=Promotion.PromoID  and ToDate >= '" & Format(Date.Now, "yyyyMMdd") & "' and FromDate <= '" & Format(DateAdd(DateInterval.Day, 7, Date.Now), "yyyyMMdd") & "'", "PromoApply")
            'dtup.UploadFile("SELECT PromoCondition.* from PromoCondition, Promotion where PromoCondition.PromoID=Promotion.PromoID  and ToDate >= '" & Format(Date.Now, "yyyyMMdd") & "' and FromDate <= '" & Format(DateAdd(DateInterval.Day, 7, Date.Now), "yyyyMMdd") & "'", "PromoCondition")
            'dtup.UploadFile("SELECT PromoOffer.[PromoID],PromoOffer.[ItemID],PromoOffer.[UOM],PromoOffer.[FOCQty],PromoOffer.[DisPrice],PromoOffer.[Discount],PromoOffer.[LineType] from PromoOffer, Promotion where PromoOffer.PromoID=Promotion.PromoID  and ToDate >= '" & Format(Date.Now, "yyyyMMdd") & "' and FromDate <= '" & Format(DateAdd(DateInterval.Day, 7, Date.Now), "yyyyMMdd") & "'", "PromoOffer")
            'dtup.UploadFile("SELECT PromoCategory.* from PromoCategory, Promotion where PromoCategory.PromoID=Promotion.PromoID  and ToDate >= '" & Format(Date.Now, "yyyyMMdd") & "' and FromDate <= '" & Format(DateAdd(DateInterval.Day, 7, Date.Now), "yyyyMMdd") & "'", "PromoCategory")
            'dtup.UploadFile("SELECT PromoEntitlement.* from PromoEntitlement, Promotion where PromoEntitlement.PromoID=Promotion.PromoID  and ToDate >= '" & Format(Date.Now, "yyyyMMdd") & "' and FromDate <= '" & Format(DateAdd(DateInterval.Day, 7, Date.Now), "yyyyMMdd") & "'", "PromoEntitlement")
            'dtup.UploadFile("SELECT * from PromoGroup", "PromoGroup")
            ''''''''''''End

            ''''''''''''After Changing to all Items Promotion 04022008
            'Arkar 20090507 Add Event Column
            dtup.UploadFile("SELECT Distinct PromoConditionTemp.PromoID, Promotion.PromoName, ApType, FromDate, ToDate, PromoType, Multiply, Priority, MinAmt, MaxAmt, DisAmt, DisPer, Entitle, EntitleType, CASE CATBased WHEN 1 Then 1 Else 0 END as CATBased, CASE ItemCondition WHEN 1 Then 1 Else 0 END as ItemCondition, Promotion.Event from Promotion, PromoConditionTemp Where Promotion.PromoId = PromoConditionTemp.ParentPromoId and ToDate >= '" & Format(Date.Now, "yyyyMMdd") & "' and FromDate <= '" & Format(DateAdd(DateInterval.Day, 7, Date.Now), "yyyyMMdd") & "' union SELECT PromoID, PromoName, ApType, FromDate, ToDate, PromoType, Multiply, Priority, MinAmt, MaxAmt, DisAmt, DisPer, Entitle, EntitleType, CASE CATBased WHEN 1 Then 1 Else 0 END as CATBased, CASE ItemCondition WHEN 1 Then 1 Else 0 END as ItemCondition, Promotion.Event from Promotion  where PromoType = 'Invoice Promotion' and ItemCondition = 0 and ToDate >= '" & Format(Date.Now, "yyyyMMdd") & "' and FromDate <= '" & Format(DateAdd(DateInterval.Day, 7, Date.Now), "yyyyMMdd") & "'", "Promotion")
            dtup.UploadFile("SELECT Distinct PromoConditionTemp.PromoId, PromoApply.ID from PromoApply, Promotion, PromoConditionTemp where Promotion.PromoId = PromoConditionTemp.ParentPromoId and PromoApply.PromoID=Promotion.PromoID  and ToDate >= '" & Format(Date.Now, "yyyyMMdd") & "' and FromDate <= '" & Format(DateAdd(DateInterval.Day, 7, Date.Now), "yyyyMMdd") & "' union SELECT PromoApply.PromoId, PromoApply.ID  from PromoApply, Promotion where PromoApply.PromoID=Promotion.PromoID and PromoType = 'Invoice Promotion' and ItemCondition = 0 and ToDate >= '" & Format(Date.Now, "yyyyMMdd") & "' and FromDate <= '" & Format(DateAdd(DateInterval.Day, 7, Date.Now), "yyyyMMdd") & "' ", "PromoApply")
            dtup.UploadFile("SELECT PromoConditionTemp.PromoID, ItemID, UOM, MinQty, MaxQty, LineType, IsRequired from PromoConditionTemp, Promotion where PromoConditionTemp.ParentPromoID=Promotion.PromoID  and ToDate >= '" & Format(Date.Now, "yyyyMMdd") & "' and FromDate <= '" & Format(DateAdd(DateInterval.Day, 7, Date.Now), "yyyyMMdd") & "'", "PromoCondition")
            dtup.UploadFile("SELECT PromoOfferTemp.PromoID , ItemID, UOM, FOCQty, DisPrice, Discount  from PromoOfferTemp, Promotion where PromoOfferTemp.ParentPromoID=Promotion.PromoID  and ToDate >= '" & Format(Date.Now, "yyyyMMdd") & "' and FromDate <= '" & Format(DateAdd(DateInterval.Day, 7, Date.Now), "yyyyMMdd") & "'", "PromoOffer")
            dtup.UploadFile("SELECT Distinct PromoConditionTemp.PromoId, CategoryID, Description from PromoCategory, Promotion,PromoConditionTemp where Promotion.PromoId = PromoConditionTemp.ParentPromoId and  PromoCategory.PromoID=Promotion.PromoID  and ToDate >= '" & Format(Date.Now, "yyyyMMdd") & "' and FromDate <= '" & Format(DateAdd(DateInterval.Day, 7, Date.Now), "yyyyMMdd") & "' union SELECT PromoCategory.PromoId, CategoryID, Description from PromoCategory, Promotion where PromoCategory.PromoID=Promotion.PromoID and PromoType = 'Invoice Promotion' and ItemCondition = 0 and ToDate >= '" & Format(Date.Now, "yyyyMMdd") & "' and FromDate <= '" & Format(DateAdd(DateInterval.Day, 7, Date.Now), "yyyyMMdd") & "' ", "PromoCategory")
            dtup.UploadFile("SELECT Distinct PromoConditionTemp.PromoId, CustID,OrdNo, OrderDate,PromoCount from PromoEntitlement, Promotion,PromoConditionTemp where Promotion.PromoId = PromoConditionTemp.ParentPromoId and  PromoEntitlement.PromoID=Promotion.PromoID  and ToDate >= '" & Format(Date.Now, "yyyyMMdd") & "' and FromDate <= '" & Format(DateAdd(DateInterval.Day, 7, Date.Now), "yyyyMMdd") & "' union SELECT PromoEntitlement.PromoId, CustID,OrdNo, OrderDate,PromoCount from PromoEntitlement, Promotion where PromoEntitlement.PromoID=Promotion.PromoID and PromoType = 'Invoice Promotion' and ItemCondition = 0 and ToDate >= '" & Format(Date.Now, "yyyyMMdd") & "' and FromDate <= '" & Format(DateAdd(DateInterval.Day, 7, Date.Now), "yyyyMMdd") & "' ", "PromoEntitlement")
            dtup.UploadFile("SELECT * from PromoGroup", "PromoGroup")
            '
            dtup.UploadFile("Select Code,Description from events", "Events")
            dtup.UploadFile("Select * from Bank", "Bank")
            dtup.UploadFile("SELECT CreditNoteNo as CrNoteNo, CreditDate as CrNoteDate, CreditNote.CustNo as CustomerID, GoodsReturnNo as RtnNo, SalesPersonCode as AgentID, Discount, Subtotal, GST, TotalAmt as Amount, 0 as Void, PaidAmt, 1 as Uploaded, 1 as UploadedToPDA, CreditNote.CompanyName from CreditNote, Customer where CreditNote.CustNo = Customer.CustNo and (PaidAmt < TotalAmt or PaidAmt is Null) and (Void <> 1 or void is null) and Customer.SalesAgent in (Select CustAgentID from CustAgent where AgentID=" & objDo.SafeSQL(sAgentID) & ")", "CrNote")
            'dtup.UploadFile("SELECT CreditNoteNo as CrNoteNo, CreditDate as CrNoteDate, CreditNote.CustNo as CustomerID, GoodsReturnNo as RtnNo, SalesPersonCode as AgentID, Discount, Subtotal, GST, TotalAmt as Amount, 0 as Void, PaidAmt, 1 as Uploaded, 1 as UploadedToPDA, CreditNote.CompanyName from CreditNote, Customer where CreditNote.CustNo = Customer.CustNo and (DateDiff(d, CreditNote.CreditDate, " & objDo.SafeSQL(Format(DateAdd(DateInterval.Day, (-1) * dInvHistory, Date.Now.Date), "yyyyMMdd HH:mm:ss")) & ") <= 0)  and (Void <> 1 or void is null) and Customer.SalesAgent in (Select CustAgentID from CustAgent where AgentID=" & objDo.SafeSQL(sAgentID) & ")", "CrNote")
            'and SalesPersonCode = " & objDo.SafeSQL(sAgentID) & " 
            'dtup.UploadFile("SELECT CreditNoteDet.CreditNoteNo as CrNoteNo, CreditNoteDet.ItemNo AS itemid, CreditNoteDet.UOM, CreditNoteDet.Qty, CreditNoteDet.Price, CreditNoteDet.Amt AS amount, BaseUOM FROM CreditNote, CreditNoteDet, Customer where CreditNote.CustNo = Customer.CustNo and CreditNote.CreditNoteNo = CreditNoteDet.CreditNoteNo and (PaidAmt < TotalAmt or PaidAmt is Null) and Customer.SalesAgent = " & objDo.SafeSQL(sAgentID) & " and Void <> 1", "CrNoteDet")
            dtup.UploadFile("SELECT CreditNoteDet.CreditNoteNo as CrNoteNo, CreditNoteDet.ItemNo AS itemid, CreditNoteDet.UOM, CreditNoteDet.Qty, CreditNoteDet.Price, CreditNoteDet.Amt AS amount, BaseUOM, '' as InvNo FROM CreditNote, CreditNoteDet, Customer where CreditNote.CustNo = Customer.CustNo and CreditNote.CreditNoteNo = CreditNoteDet.CreditNoteNo and (PaidAmt < TotalAmt or PaidAmt is Null) and (Void <> 1 or void is null) and Customer.SalesAgent in (Select CustAgentID from CustAgent where AgentID=" & objDo.SafeSQL(sAgentID) & ")", "CrNoteDet")
            'dtup.UploadFile("SELECT * from Upload", "CrNote")
            'dtup.UploadFile("SELECT * from Upload", "CrNoteDet")
            dtup.UploadFile("SELECT * from Upload", "Exchange")
            dtup.UploadFile("SELECT * from Upload", "Exdet")
            dtup.UploadFile("SELECT * from Upload", "NewCust")
            dtup.UploadFile("SELECT * from Upload", "Rtn")
            dtup.UploadFile("SELECT * from Upload", "RtnDet")
            dtup.UploadFile("SELECT * from Upload", "StockTakeHdr")
            dtup.UploadFile("SELECT * from Upload", "StockTakeDet")
            dtup.UploadFile("SELECT * from Upload", "Exception")
            dtup.UploadFile("SELECT Distinct ItemNo, ItemName, UOM from SpecialItem", "SpecialItem")

            'dtup.UploadFile("SELECT RcptNo as ReceiptNo, RcptDt as ReceiptDt, Amount, PayMethod as PaymentMode, ChqNo as ExtDocNo, CurCode, CurExrate, CustID as CustomerID, DTG, AgentID, ChqDT, 0 as Void, BankName, Chqtype, CompanyName from Receipt where Void <> 1 and (ExportNav=0 or ExportNav is Null) and ChqType='PDC' and AgentID in (Select CustAgentID from CustAgent where AgentID=" & objDo.SafeSQL(sAgentID) & ")", "Receipts")
            dtup.UploadFile("SELECT RcptNo as ReceiptNo, RcptDt as ReceiptDt, Amount, PayMethod as PaymentMode, ChqNo as ExtDocNo, CurCode, CurExrate, CustID as CustomerID, DTG, AgentID, ChqDT, 0 as Void, BankName, Chqtype, CompanyName, CASE Remarks WHEN null Then '' Else Remarks END as Remarks from Receipt where Void <> 1 and (ExportNav=0 or ExportNav is Null) and ChqType='PDC' and AgentID in (Select CustAgentID from CustAgent where AgentID=" & objDo.SafeSQL(sAgentID) & ")", "Receipts")

            dtup.UploadFile("SELECT RcptItem.RcptNo as ReceiptNo, InvNo, AmtPaid as Amount from RcptItem, Receipt where RcptItem.RcptNo =Receipt.RcptNo and Void <> 1 and (Receipt.ExportNav=0 or Receipt.ExportNav is Null) and Receipt.ChqType='PDC' and AgentID in (Select CustAgentID from CustAgent where AgentID=" & objDo.SafeSQL(sAgentID) & ")", "ReceiptDet")
            dtup.UploadFile("SELECT Barcodes.* from Barcodes, Item where Item.ItemNo=Barcodes.ItemNo and Item.ToPDA=1", "Barcodes")
            dtup.UploadFile("SELECT * from Upload", "CustVisit")
            dtup.UploadFile("SELECT * from Upload", "Service")
            dtup.UploadFile("Select Code, Name as Description from Location", "Location")
            dtup.UploadFile("SELECT Code, Description, ReasonType from Reason", "Reason")
            dtup.UploadFile("Select * from ServiceMaster", "ServiceMaster")
            dtup.UploadFile("Select DailySchedule.CustNo as CustomerId, LastVisitedDate as LastVisited, CASE Missed WHEN 1 Then 1 Else 0 END as Missed, 0 as Visited, NextVisit, NoofMissed  from DailySchedule,RouteDet, RouteMaster where DailySchedule.CustNo = RouteDet.CustNo and RouteMaster.RouteNo = RouteDet.RouteNo and VehicleNo = " & objDo.SafeSQL(sVehicle), "DailySchedule")
            dtup.UploadFile("Select Code, Name, DisplayNo from Color order by DisplayNo", "Color")
            'dtup.UploadFile("SELECT * from Upload", "ConsignInvn")
            dtup.UploadFile("Select Code, Description from CustGroup", "CustGroup")
            dtup.UploadFile("Select * from Upload", "Orders")
            dtup.UploadFile("Select * from Upload", "OrderDet")
            dtup.UploadFile("Select AgentId, RemarksDate, Remarks from AgentRemarks where AgentID = " & objDo.SafeSQL(sAgentID), "AgentRemarks")
            dtup.UploadFile("Select AgentID, Month, IsNull(Target,0) as Target, IsNull(Actual,0) as Actual from SalesTargetByMonth", "SalesTargetByMonth")
            'Dim sPicFolder As String = GetPictureFolder()
            ''arkar
            'If Not Directory.Exists(sPicFolder) Then
            '    Try
            '        Directory.CreateDirectory(sPicFolder)
            '    Catch ex As Exception
            '    End Try
            'End If

            'Dim arrPic As String()
            'arrPic = Directory.GetFiles(sPicFolder)
            'For i As Integer = 0 To arrPic.Length - 1
            '    If System.IO.Path.GetFileName(arrPic(i)).ToLower.IndexOf(".jpg") > -1 Or System.IO.Path.GetFileName(arrPic(i)).ToLower.IndexOf(".jpeg") > -1 Or System.IO.Path.GetFileName(arrPic(i)).ToLower.IndexOf(".gif") > -1 Or System.IO.Path.GetFileName(arrPic(i)).ToLower.IndexOf(".bmp") > -1 Then
            '        If System.IO.File.Exists(sPicFolder & "\" & System.IO.Path.GetFileName(arrPic(i))) = True Then
            '            dtup.UploadItemPhoto(sPicFolder, System.IO.Path.GetFileName(arrPic(i)))
            '        Else
            '            MsgBox("Picture: " & sPicFolder & "\" & System.IO.Path.GetFileName(arrPic(i)) & " is not exists!!", MsgBoxStyle.Critical, "Upload Picture")
            '        End If
            '    End If
            'Next

        Catch ex As Exception
            MsgBox(ex.Message)
            bStatus = False
        End Try
        dtup.EndUpload()
        dtup.Disconnect()
        objDo.DisconnectDB()
        frm.Close()
        Return bStatus
    End Function

    Public Function GetNewTransNo() As String
        Dim dtr As SqlDataReader

        Dim MDTNo As String = ""
        dtr = objDo.ReadRecord("Select * from System, MDT Where System.MDTNo = MDT.MDTNo")
        If dtr.Read = True Then
            Dim sPre As String = dtr("PreITNo")
            Dim iLen As Int32 = dtr("LenITNo")
            Dim iOrdNo As Int32 = dtr("LastITNo") + 1
            GetNewTransNo = sPre & StrDup(iLen - Len(sPre) - Len(CStr(iOrdNo)), "0") & CStr(iOrdNo)
            MDTNo = dtr("MDTNo").ToString
        Else
            GetNewTransNo = ""
        End If
        dtr.Close()
        objDo.ExecuteSQL("Update MDT Set LastITNo = LastITNo + 1 where MDTNO = " & objDo.SafeSQL(MDTNo))
    End Function

    Private Sub PPCSync_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'TODO: This line of code loads data into the 'SalesDataSet.MDT' table. You can move, or remove it, as needed.
        'dtpScheduleDate.Value = DateAdd(DateInterval.Day, 1, Date.Now)
        Me.MDTTableAdapter.Fill(Me.SalesDataSet.MDT)
        'btnUpdate.Visible = False


        ' objDo.ConnectDB()
        'Dim rs As SqlDataReader
        'rs = objDo.ReadRecord("Select status from System")
        'If rs.Read = True Then
        '    If rs("status") = "Not Completed" Then
        '        MsgBox("Auto import is not successful. Please check on the Import/Export Screen!")
        '    End If

        'End If
        'objDo.DisconnectDB()
    End Sub

    Private Sub dgvSync_CellMouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvSync.CellMouseDown
        Try
            sAgentID = dgvSync.Item(2, e.RowIndex).Value
        Catch ex As Exception

        End Try
    End Sub

    Private Sub dgvSync_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSync.CellContentClick
        Try
            sAgentID = dgvSync.Item(2, e.RowIndex).Value
        Catch ex As Exception

        End Try

    End Sub

    Private Sub btnDownload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDownload.Click
        Download()
        MsgBox("Completed")
    End Sub

    Private Function Download() As Boolean
        If dtup.Connect(sConStr) = False Then
            MsgBox("Please connect device and try again.")
            dtup.Disconnect()
            Return False
        End If
        Dim path As String
        path = System.AppDomain.CurrentDomain.BaseDirectory & "download\"
        Dim frm As New status
        frm.Show()
        Dim diN As New IO.DirectoryInfo(path)
        Dim aryFiN As IO.FileInfo() = diN.GetFiles("*.*")
        Dim fiN As IO.FileInfo
        For Each fiN In aryFiN
            System.IO.File.Delete(path & fiN.Name)
        Next
        'dtup.Connect(sConStr)
        'Try
        Dim arr() As String = {"Orders", "OrderDet", "Invoices", "InvDet", "Receipts", "ReceiptDet", "System", "ItemTrans", "Exchange", "ExDet", "Rtn", "RtnDet", "CrNote", "CrNoteDet", "CustVisit", "NewCust", "Service", "Agent", "Exception", "Barcodes", "StockTakeHdr", "StockTakeDet"}
        dtup.arrDownload = arr
        dtup.Download()
        objDo.ConnectDB()
        frm.Refresh()
        If IO.File.Exists(path & "System.xml") = False Or IO.File.Exists(path & "System.xsd") = False Then
            objDo.DisconnectDB()
            dtup.Disconnect()
            frm.Close()
            Return True
        End If
        For Each str As String In arr
            If IO.File.Exists(path & str & ".xml") = False Or IO.File.Exists(path & str & ".xsd") = False Then
                MsgBox("Run application in the handheld then exit", MsgBoxStyle.Critical, "File not present")
                objDo.DisconnectDB()
                dtup.Disconnect()
                frm.Close()
                Return False
            End If
        Next
        frm.Refresh()

        DownloadMDT(path & "System.xml", path & "system.xsd")
        DownloadOrders(path & "Orders.xml", path & "Orders.xsd")
        DownloadOrdItem(path & "Orderdet.xml", path & "OrderDet.xsd")
        DownloadInvoices(path & "Invoices.xml", path & "Invoices.xsd")
        DownloadInvItem(path & "InvDet.xml", path & "InvDet.xsd")
        frm.Refresh()
        DownloadReceipt(path & "Receipts.xml", path & "Receipts.xsd")
        DownloadRcptItem(path & "ReceiptDet.xml", path & "ReceiptDet.xsd")
        DownloadItemTrans(path & "ItemTrans.xml", path & "ItemTrans.xsd")
        frm.Refresh()
        DownloadExchange(path & "Exchange.xml", path & "Exchange.xsd")
        DownloadExchangeItem(path & "ExDet.xml", path & "ExDet.xsd")
        DownloadReturn(path & "Rtn.xml", path & "Rtn.xsd")
        frm.Refresh()
        DownloadReturnItem(path & "RtnDet.xml", path & "RtnDet.xsd")
        DownloadCreditNote(path & "CrNote.xml", path & "CrNote.xsd")
        DownloadCreditNoteDet(path & "CrNoteDet.xml", path & "CrNoteDet.xsd")
        'DownloadStockTakeHdr(path & "StockTakeHdr.xml", path & "StockTakeHdr.xsd")
        'DownloadStockTakeDet(path & "StockTakeDet.xml", path & "StockTakeDet.xsd")
        DownloadGoodsInvn(path & "VanInventory.xml", path & "VanInventory.xsd")
        'DownloadConsignInvn(path & "ConsignInvn.xml", path & "ConsignInvn.xsd")
        DownloadCustVisit(path & "CustVisit.xml", path & "CustVisit.xsd")
        DownloadSalesTarget(path & "Agent.xml", path & "Agent.xsd")
        DownloadNewCust(path & "NewCust.xml", path & "NewCust.xsd")
        DownloadService(path & "Service.xml", path & "Service.xsd")
        DownloadException(path & "Exception.xml", path & "Exception.xsd")
        DownloadBarcodes(path & "Barcodes.xml", path & "Barcodes.xsd")

        'DownloadRtnReject(path & "RtnReject.xml", path & "RtnReject.xsd")
        Try


            'dtup.arrDownloadImages = arrImages

            'dtup.DownloadImages()

            'Dim sSignatureFolder As String = path.Replace("download", "Signature")
            'sSignatureFolder = sSignatureFolder.Replace("Download", "Signature")
            'If IO.Directory.Exists(sSignatureFolder) = False Then
            '    IO.Directory.CreateDirectory(sSignatureFolder)
            'End If

            'Dim diSign As New DirectoryInfo(path)

            'For Each fiSign As FileInfo In diSign.GetFiles("*.bmp")
            '    If File.Exists(sSignatureFolder & fiSign.Name) = False Then
            '        fiSign.MoveTo(sSignatureFolder & fiSign.Name)
            '    End If
            'Next
        Catch ex As Exception

        End Try
        'Dim aryFiSign As IO.FileInfo() = diSignature.GetFiles("*.*")
        'Dim fiSign As IO.FileInfo
        'For Each fiSign In aryFiSign
        '    System.IO.File.Move(path & fiSign.Name, sDirName & "\" & fiSign.Name)
        'Next


        Dim sDirName As String = path & "Backup\" & Format(Date.Now, "yyyyMMddHHmmss")
        System.IO.Directory.CreateDirectory(sDirName)
        Dim strFileSize As String = ""
        Dim di As New IO.DirectoryInfo(path)
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.*")
        Dim fi As IO.FileInfo
        For Each fi In aryFi
            System.IO.File.Move(path & fi.Name, sDirName & "\" & fi.Name)
        Next

        objDo.DisconnectDB()
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try
        dtup.Disconnect()
        frm.Close()
        Return True
    End Function

    Private Sub DownloadNewCust(ByVal FileName As String, ByVal XsdFileName As String)
        Dim ds As New DataSet
        Dim CommandSQL As String = ""
        Dim ValueSQL As String = ""
        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Exit Sub
        ds.ReadXmlSchema(XsdFileName)
        ds.ReadXml(FileName)
        arrDet = New ArrayList
        For Each row As DataRow In ds.Tables(0).Rows
            For Each column As DataColumn In ds.Tables(0).Columns
                If row(column) Is DBNull.Value Then
                    If column.DataType.ToString = "System.String" Then
                        row(column) = ""
                    ElseIf column.DataType.ToString = "System.DateTime" Then
                        row(column) = CDate("1/1/2000")
                    Else
                        row(column) = 0
                    End If
                End If
            Next
            If CheckAvailable("Select * from NewCust where CustID = " & objDo.SafeSQL(row("CustomerID").ToString)) = True Then
                objDo.ExecuteSQL("Delete from NewCust where CustID=" & objDo.SafeSQL(row("CustomerID").ToString))
                objDo.ExecuteSQL("Insert into NewCust (CustID, CustName, Contact, Address, Address2, City, Pin, Email, Phone, EditDate, AgentID, CompanyName) values (" & objDo.SafeSQL(row("CustomerID")) & "," & objDo.SafeSQL(row("CustomerName")) & "," & objDo.SafeSQL(row("Contact")) & "," & objDo.SafeSQL(row("Address")) & "," & objDo.SafeSQL(row("Address2")) & "," & objDo.SafeSQL(row("City")) & "," & objDo.SafeSQL(row("Pin")) & "," & objDo.SafeSQL(row("Email")) & "," & objDo.SafeSQL(row("Phone")) & "," & objDo.SafeSQL(Format(row("EditDate"), "yyyyMMdd HH:mm:ss")) & "," & objDo.SafeSQL(sAgentID) & "," & objDo.SafeSQL(row("CompanyName")) & ")")
            Else
                objDo.ExecuteSQL("Insert into NewCust (CustID, CustName, Contact, Address, Address2, City, Pin, Email, Phone, EditDate, AgentID, CompanyName) values (" & objDo.SafeSQL(row("CustomerID")) & "," & objDo.SafeSQL(row("CustomerName")) & "," & objDo.SafeSQL(row("Contact")) & "," & objDo.SafeSQL(row("Address")) & "," & objDo.SafeSQL(row("Address2")) & "," & objDo.SafeSQL(row("City")) & "," & objDo.SafeSQL(row("Pin")) & "," & objDo.SafeSQL(row("Email")) & "," & objDo.SafeSQL(row("Phone")) & "," & objDo.SafeSQL(Format(row("EditDate"), "yyyyMMdd HH:mm:ss")) & "," & objDo.SafeSQL(sAgentID) & "," & objDo.SafeSQL(row("CompanyName")) & ")")
            End If
        Next row
    End Sub
    Private Sub DownloadException(ByVal FileName As String, ByVal XsdFileName As String)
        Dim ds As New DataSet
        Dim CommandSQL As String = ""
        Dim ValueSQL As String = ""
        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Exit Sub
        ds.ReadXmlSchema(XsdFileName)
        ds.ReadXml(FileName)
        'arrDet.Clear()
        For Each row As DataRow In ds.Tables(0).Rows
            For Each column As DataColumn In ds.Tables(0).Columns
                If row(column) Is DBNull.Value Then
                    If column.DataType.ToString = "System.String" Then
                        row(column) = ""
                    ElseIf column.DataType.ToString = "System.DateTime" Then
                        row(column) = CDate("1/1/2000")
                    Else
                        row(column) = 0
                    End If
                End If
            Next
            If CheckAvailable("Select * from Exception where DocNo = " & objDo.SafeSQL(row("DocNo")) & " and DocType = " & objDo.SafeSQL(row("DocType")) & " and CustID = " & objDo.SafeSQL(row("CustID")) & " and ColType=" & objDo.SafeSQL(row("DocType")) & " and ItemID = " & objDo.SafeSQL(row("ItemID")) & " and AgentID = " & objDo.SafeSQL(row("AgentID"))) = True Then
                objDo.ExecuteSQL("Delete from Exception where DocNo = " & objDo.SafeSQL(row("DocNo")) & " and DocType = " & objDo.SafeSQL(row("DocType")) & " and CustID = " & objDo.SafeSQL(row("CustID")) & " and ColType=" & objDo.SafeSQL(row("DocType")) & " and ItemID = " & objDo.SafeSQL(row("ItemID")) & " and AgentID = " & objDo.SafeSQL(row("AgentID")))
                GoTo InsertData
            Else
InsertData:
                'Insert Command
                objDo.ExecuteSQL("Insert into Exception(CustID, DocNo, DocType, ItemId, ColType, AgentID, DocDate, Remarks) values (" & objDo.SafeSQL(row("CustID")) & "," & objDo.SafeSQL(row("DocNo")) & "," & objDo.SafeSQL(row("DocType")) & "," & objDo.SafeSQL(row("ItemID")) & "," & objDo.SafeSQL("") & "," & objDo.SafeSQL(row("AgentID")) & "," & objDo.SafeSQL(Format(row("DocDate"), "yyyyMMdd HH:mm:ss")) & "," & objDo.SafeSQL(row("Remarks")) & ")")
            End If
        Next row
    End Sub

    Private Sub DownloadBarcodes(ByVal FileName As String, ByVal XsdFileName As String)
        Dim ds As New DataSet
        Dim CommandSQL As String = ""
        Dim ValueSQL As String = ""
        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Exit Sub
        ds.ReadXmlSchema(XsdFileName)
        ds.ReadXml(FileName)
        'arrDet.Clear()
        For Each row As DataRow In ds.Tables(0).Rows
            For Each column As DataColumn In ds.Tables(0).Columns
                If row(column) Is DBNull.Value Then
                    If column.DataType.ToString = "System.String" Then
                        row(column) = ""
                    ElseIf column.DataType.ToString = "System.DateTime" Then
                        row(column) = CDate("1/1/2000")
                    Else
                        row(column) = 0
                    End If
                End If
            Next
            If CheckAvailable("Select * from Barcodes where ItemNo = " & objDo.SafeSQL(row("ItemNo")) & " and Barcode = " & objDo.SafeSQL(row("Barcode"))) = False Then
                objDo.ExecuteSQL("Insert into Barcodes(ItemNo, Uom, Barcode) values(" & objDo.SafeSQL(row("ItemNo")) & "," & objDo.SafeSQL(row("Uom")) & "," & objDo.SafeSQL(row("Barcode")) & ")")
            End If
        Next row
    End Sub

    Private Sub DownloadService(ByVal FileName As String, ByVal XsdFileName As String)
        Dim ds As New DataSet
        Dim CommandSQL As String = ""
        Dim ValueSQL As String = ""
        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Exit Sub
        ds.ReadXmlSchema(XsdFileName)
        ds.ReadXml(FileName)
        arrDet = New ArrayList
        For Each row As DataRow In ds.Tables(0).Rows
            For Each column As DataColumn In ds.Tables(0).Columns
                If row(column) Is DBNull.Value Then
                    If column.DataType.ToString = "System.String" Then
                        row(column) = ""
                    ElseIf column.DataType.ToString = "System.DateTime" Then
                        row(column) = CDate("1/1/2000")
                    Else
                        row(column) = 0
                    End If
                End If
            Next
            objDo.ExecuteSQL("Delete From Service where SerInvNo = " & objDo.SafeSQL(row("SerInvNo")))
            objDo.ExecuteSQL("Insert into Service (SerInvNo, ServiceID, Details, ServiceDt, CustomerID, AgentID, ReasonCode) values (" & objDo.SafeSQL(row("SerInvNo")) & "," & objDo.SafeSQL(row("ServiceID")) & "," & objDo.SafeSQL(row("Details")) & ",'" & Format(row("ServiceDt"), "yyyyMMdd HH:mm:ss") & "'," & objDo.SafeSQL(row("CustomerID")) & "," & objDo.SafeSQL(sAgentID) & "," & objDo.SafeSQL(row("ReasonCode")) & ")")
            'objDo.ExecuteSQL("Insert into Service (SerInvNo, ServiceID, Details, ServiceDt, CustomerID) values (" & objDo.SafeSQL(row("SerInvNo")) & "," & objDo.SafeSQL(row("ServiceID")) & "," & objDo.SafeSQL(row("Details")) & ",'" & Format(row("ServiceDt"), "yyyyMMdd HH:mm:ss") & "'," & objDo.SafeSQL(row("CustomerID")) ")")
        Next row
    End Sub

    Private Sub DownloadCustVisit(ByVal FileName As String, ByVal XsdFileName As String)
        Dim ds As New DataSet
        Dim CommandSQL As String = ""
        Dim ValueSQL As String = ""
        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Exit Sub

        Dim iWeek As Integer = Date.Now.Day
        Dim sWeek As String = ""
        If iWeek <= 7 Then
            sWeek = "Week1"
        ElseIf iWeek > 7 And iWeek <= 14 Then
            sWeek = "Week2"
        ElseIf iWeek > 14 And iWeek <= 21 Then
            sWeek = "Week3"
        ElseIf iWeek > 21 And iWeek <= 28 Then
            sWeek = "Week4"
        Else
            sWeek = "Week1"
        End If
        objDo.ExecuteSQL("Update DailySchedule set Missed = 1, Visited = 0 from DailySchedule,RouteDet, RouteMaster where DailySchedule.CustNo = RouteDet.CustNo and RouteMaster.RouteNo = RouteDet.RouteNo and RouteWeek = " & objDo.SafeSQL(sWeek) & " and VehicleNo = " & objDo.SafeSQL(sVehicle))

        ds.ReadXmlSchema(XsdFileName)
        ds.ReadXml(FileName)
        arrDet = New ArrayList
        For Each row As DataRow In ds.Tables(0).Rows
            For Each column As DataColumn In ds.Tables(0).Columns
                If row(column) Is DBNull.Value Then
                    If column.DataType.ToString = "System.String" Then
                        row(column) = ""
                    ElseIf column.DataType.ToString = "System.DateTime" Then
                        row(column) = CDate("1/1/2000")
                    Else
                        row(column) = 0
                    End If
                End If
            Next
            objDo.ExecuteSQL("Insert into CustVisit (CustID, TransNo, TransType,  TransDate, AgentID) values (" & objDo.SafeSQL(row("CustID")) & "," & objDo.SafeSQL(row("TransNo")) & "," & objDo.SafeSQL(row("TransType")) & ",'" & Format(row("TransDate"), "yyyyMMdd HH:mm:ss") & "'," & objDo.SafeSQL(row("AgentID")) & ")")
            If CheckAvailable("Select * from DailySchedule where CustNo = " & objDo.SafeSQL(row("CustID"))) = True Then
                objDo.ExecuteSQL("Update DailySchedule set LastVisitedDate ='" & Format(row("TransDate"), "yyyyMMdd HH:mm:ss") & "', NextVisit = '" & Format(GetNextSchedule(row("CustID").ToString, row("TransDate")), "yyyyMMdd") & "', Visited = 1, Missed = 0, NoOfMissed = 0  where CustNo=" & objDo.SafeSQL(row("CustID")))
            Else
                objDo.ExecuteSQL("Insert into DailySchedule values(" & objDo.SafeSQL(row("CustID")) & "," & objDo.SafeSQL(Format(row("TransDate"), "yyyyMMdd HH:mm:ss")) & ",0,1," & objDo.SafeSQL(Format(GetNextSchedule(row("CustID").ToString, row("TransDate")), "yyyyMMdd")) & ",0)")
            End If
        Next row

        Dim dtrMissed As SqlDataReader
        dtrMissed = objDo.ReadRecordAnother("Select DailySchedule.* from DailySchedule,RouteDet, RouteMaster where DailySchedule.CustNo = RouteDet.CustNo and RouteMaster.RouteNo = RouteDet.RouteNo and RouteWeek = " & objDo.SafeSQL(sWeek) & " and VehicleNo = " & objDo.SafeSQL(sVehicle) & " and Missed = 1")
        While dtrMissed.Read = True
            If DateDiff(DateInterval.Day, Date.Now, GetNextSchedule(dtrMissed("CustNo").ToString, dtrMissed("NextVisit"))) > 0 Then
                objDo.ExecuteSQL("Update DailySchedule set NextVisit = '" & Format(GetNextSchedule(dtrMissed("CustNo").ToString, dtrMissed("NextVisit")), "yyyyMMdd") & "', Visited = 1, Missed = 0, NoOfMissed = isnull(NoOfMissed,0) + 1 where CustNo=" & objDo.SafeSQL(dtrMissed("CustNo").ToString))
            End If
        End While
        dtrMissed.Close()
    End Sub

    Private Sub DownloadSalesTarget(ByVal FileName As String, ByVal XsdFileName As String)
        Dim ds As New DataSet
        Dim CommandSQL As String = ""
        Dim ValueSQL As String = ""
        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Exit Sub
        ds.ReadXmlSchema(XsdFileName)
        ds.ReadXml(FileName)
        arrDet = New ArrayList
        For Each row As DataRow In ds.Tables(0).Rows
            For Each column As DataColumn In ds.Tables(0).Columns
                If row(column) Is DBNull.Value Then
                    If column.DataType.ToString = "System.String" Then
                        row(column) = ""
                    ElseIf column.DataType.ToString = "System.DateTime" Then
                        row(column) = CDate("1/1/2000")
                    Else
                        row(column) = 0
                    End If
                End If
            Next
            If sAgt = row("AgentNo") Then
                'CurMonTarget=" & CStr(row("CurMonTarget")) & ", 
                objDo.ExecuteSQL("Update SalesAgent set CurPDCChqCollection=" & CStr(row("CurPDCChqCollection")) & ", CurChqCollection=" & CStr(row("CurChqCollection")) & ", CurCashCollection=" & CStr(row("CurCashCollection")) & ", CurCreditNote=" & CStr(row("CurCreditNote")) & ",CurMonTarget=" & CStr(row("CurMonTarget")) & ",LastInvDate ='" & Format(row("LastInvDt"), "yyyyMMdd HH:mm:ss") & "' where Code=" & objDo.SafeSQL(row("AgentNo")))
            End If
        Next row
    End Sub


    Private Function GetNextSchedule(ByVal CustNo As String, ByVal VisitDate As Date) As Date
        Dim dtr As SqlDataReader
        Dim iD As Integer = 0
        Dim dtNext As Date = DateAdd(DateInterval.Year, 1, Now)
        Dim iW As Integer = 0
        Dim iWeek As Integer = VisitDate.Day
        Dim sWeek As Integer = 1
        If iWeek <= 7 Then
            sWeek = 1
        ElseIf iWeek > 7 And iWeek <= 14 Then
            sWeek = 2
        ElseIf iWeek > 14 And iWeek <= 21 Then
            sWeek = 3
        ElseIf iWeek > 21 And iWeek <= 28 Then
            sWeek = 4
        Else
            sWeek = 1
        End If

        dtr = objDo.ReadRecord("Select RouteDay - " & VisitDate.DayOfWeek & " as RD, right(RouteWeek,1) - " & sWeek & " as WeekNo from RouteMaster, RouteDet where RouteMaster.RouteNo = RouteDet.RouteNo and CustNo = " & objDo.SafeSQL(CustNo) & " order by RouteDay, RouteWeek")
        While dtr.Read
            iD = dtr("RD")
            If iD < 0 Then iD = iD * -1
            iW = dtr("WeekNo")
            If iW <= 0 Then iW = iW + 4
            If DateAdd(DateInterval.Day, iD + 7 * iW, VisitDate) < dtNext Then dtNext = DateAdd(DateInterval.Day, iD + 7 * iW, VisitDate)
        End While
        dtr.Close()
        Return dtNext
    End Function

    'Private Function GetNextSchedule(ByVal CustNo As String, ByVal VisitDate As Date) As Date
    '    Dim dtr As SqlDataReader
    '    Dim iD As Integer = 0
    '    Dim dtNext As Date = DateAdd(DateInterval.Year, 1, Now)
    '    dtr = objDo.ReadRecord("Select RouteDay - " & VisitDate.DayOfWeek & " as RD from RouteMaster, RouteDet where RouteMaster.RouteNo = RouteDet.RouteNo and CustNo = " & objDo.SafeSQL(CustNo) & " order by RouteDay")
    '    While dtr.Read
    '        iD = dtr("RD")
    '        If iD <= 0 Then iD = iD + 7
    '        If DateAdd(DateInterval.Day, iD, VisitDate) < dtNext Then dtNext = DateAdd(DateInterval.Day, iD, VisitDate)
    '    End While
    '    dtr.Close()
    '    Return dtNext
    'End Function

    Private Sub DownloadOrders(ByVal FileName As String, ByVal XsdFileName As String)
        arrOrdNo = New ArrayList
        Dim ds As New DataSet
        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Exit Sub
        ds.ReadXmlSchema(XsdFileName)
        ds.ReadXml(FileName)
        arrDet = New ArrayList
        Dim cmd As SqlCommand = New SqlCommand("", objDo.MyConnection)
        For Each row As DataRow In ds.Tables(0).Rows
            For Each column As DataColumn In ds.Tables(0).Columns
                If row(column) Is DBNull.Value Then
                    If column.DataType.ToString = "System.String" Then
                        row(column) = ""
                    ElseIf column.DataType.ToString = "System.DateTime" Then
                        row(column) = CDate("1/1/2000")
                    Else
                        row(column) = 0
                    End If
                End If
            Next

            If cmd.Parameters.Count > 0 Then cmd.Parameters.Clear()

            ImgCount += 1
            ReDim Preserve arrImages(ImgCount)
            arrImages(ImgCount) = row("OrderNo") & ".bmp"

            dtup.DownloadOrderSign(row("OrderNo").ToString)

            Dim bImageExists As Boolean = True
            Dim sImgLocation As String = ""
            Dim sImgFolder As String
            sImgFolder = System.AppDomain.CurrentDomain.BaseDirectory & "Signature\"
            If IO.Directory.Exists(sImgFolder) Then
                If IO.File.Exists(sImgFolder & arrImages(ImgCount).ToString) Then
                    sImgLocation = sImgFolder & arrImages(ImgCount).ToString
                Else
                    bImageExists = False
                End If
            Else
                IO.Directory.CreateDirectory(sImgFolder)
                bImageExists = False
            End If

            If bImageExists = True Then
                Try
                    Dim imageProduct As Image = Image.FromFile(sImgLocation)

                    Dim bm As New Bitmap(imageProduct)
                    Dim ms As New MemoryStream

                    bm.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp)
                    Dim bytImageData(ms.Length() - 1) As Byte
                    bytImageData = ms.ToArray

                    Dim IMG_DATA As New SqlParameter("@Signature", SqlDbType.Image, _
                    bytImageData.Length, ParameterDirection.Input, False, _
                    0, 0, Nothing, DataRowVersion.Current, bytImageData)
                    cmd.Parameters.Add(IMG_DATA)

                    imageProduct.Dispose()
                    bm.Dispose()
                    ms.Close()
                    ms.Dispose()
                Catch ex As Exception
                    bImageExists = False
                End Try
            End If

            If IsDBNull(row("Delivered")) = False And IsDBNull(row("Consignment")) = False Then
                If CBool(row("Delivered")) = True And CBool(row("Consignment")) = True Then
                    ImgCount += 1
                    ReDim Preserve arrImages(ImgCount)
                    arrImages(ImgCount) = row("OrderNo") & "_Deli" & ".bmp"
                End If
            End If


            If CheckAvailable("Select * from OrderHdr where OrdNo = " & objDo.SafeSQL(row("OrderNo").ToString)) = True Then
                If CheckAvailable("Select * from OrderHdr where OrdNo = " & objDo.SafeSQL(row("OrderNo").ToString) & " and CustId=" & objDo.SafeSQL(row("CustomerID").ToString)) = True Then
                    If bImageExists = True Then
                        cmd.CommandText = "Update OrderHdr set void = " & IIf(row("Void"), 1, 0) & ", Delivered = " & IIf(CBool(row("Delivered")) = True, 1, 0) & ",CustSign = @Signature where OrderHdr.OrdNo = " & objDo.SafeSQL(row("OrderNo").ToString)
                    Else
                        cmd.CommandText = "Update OrderHdr set void = " & IIf(row("Void"), 1, 0) & ", Delivered = " & IIf(CBool(row("Delivered")) = True, 1, 0) & " where OrderHdr.OrdNo = " & objDo.SafeSQL(row("OrderNo").ToString)
                    End If
                    cmd.ExecuteNonQuery()
                Else
                    arrDet.Add(row("OrderNo").ToString)
                    arrOrdNo.Add(row("OrderNo").ToString)
                    objDo.ExecuteSQL("Delete from OrderHdr where OrdNo = " & objDo.SafeSQL(row("OrderNo").ToString))
                    'objDo.ExecuteSQL("Update OrderHdr set void = " & IIf(row("Void"), 1, 0) & " where OrderHdr.OrdNo = " & objDo.SafeSQL(row("OrderNo").ToString))
                    objDo.ExecuteSQL("Delete from OrdItem where OrdNo = " & objDo.SafeSQL(row("OrderNo").ToString))
                    'Update Command
                    GoTo InsertData
                End If
            Else

                arrDet.Add(row("OrderNo").ToString)
InsertData:
                'Insert Command
                Dim iConsign As Integer = 0
                iConsign = IIf(IIf(IsDBNull(row("Consignment")) = True, False, CBool(row("Consignment"))) = True, 1, 0)

                If iConsign = 1 Then
                    cmd.CommandText = "Insert into OrderHdr(OrdNo, OrdDt, PONo, Discount, SubTotal, GstAmt, TotalAmt, PayTerms, TermDays, Delivered, CurCode, CurExRate, CustId, DTG, AgentID, Exported, ExportMYOB, Void,VoidDate,DeliveryDate,AcBillRef,DisPer,CompanyName,Remarks, ShipName, ShipAdd, ShipAdd2, ShipAdd3, ShipAdd4, ShipPin, ShipCity, MDTNo, Consignment, CustSign) values (" & objDo.SafeSQL(row("OrderNo")) & ", " & objDo.SafeSQL(Format(row("OrdDt"), "yyyyMMdd HH:mm:ss")) & ", " & objDo.SafeSQL(row("CustRef")) & ", " & CStr(row("discount")) & ", " & CStr(row("netamt")) & "," & CStr(row("gst")) & "," & CStr(row("amount")) & "," & objDo.SafeSQL(row("Terms")) & "," & CStr(row("TermDays")) & "," & IIf(CBool(row("Delivered")) = True, 1, 0) & "," & objDo.SafeSQL(row("CurCode")) & "," & CStr(row("CurExRate")) & "," & objDo.SafeSQL(row("CustomerID")) & "," & objDo.SafeSQL(Format(row("DTG"), "yyyyMMdd HH:mm:ss")) & "," & objDo.SafeSQL(row("AgentId")) & ", 1,0," & IIf(row("Void"), 1, 0) & "," & objDo.SafeSQL(Format(row("VoidDate"), "yyyyMMdd HH:mm:ss")) & "," & objDo.SafeSQL(Format(row("DeliDate"), "yyyyMMdd HH:mm:ss")) & ", " & objDo.SafeSQL(row("AcBillRef")) & "," & row("DiscountPer").ToString & "," & objDo.SafeSQL(row("CompanyName")) & "," & objDo.SafeSQL(row("Remarks")) & "," & objDo.SafeSQL(row("ShipName").ToString) & "," & objDo.SafeSQL(row("ShipAdd").ToString) & "," & objDo.SafeSQL(row("ShipAdd2").ToString) & "," & objDo.SafeSQL(row("ShipAdd3").ToString) & "," & objDo.SafeSQL(row("ShipAdd4").ToString) & "," & objDo.SafeSQL(row("ShipPin").ToString) & "," & objDo.SafeSQL(row("ShipCity").ToString) & "," & objDo.SafeSQL(sMDTNo) & "," & iConsign & ",@Signature)"
                Else
                    cmd.CommandText = "Insert into OrderHdr(OrdNo, OrdDt, PONo, Discount, SubTotal, GstAmt, TotalAmt, PayTerms, TermDays, Delivered, CurCode, CurExRate, CustId, DTG, AgentID, Exported, ExportMYOB, Void,VoidDate,DeliveryDate,AcBillRef,DisPer,CompanyName,Remarks, ShipName, ShipAdd, ShipAdd2, ShipAdd3, ShipAdd4, ShipPin, ShipCity, MDTNo, Consignment, CustSign) values (" & objDo.SafeSQL(row("OrderNo")) & ", " & objDo.SafeSQL(Format(row("OrdDt"), "yyyyMMdd HH:mm:ss")) & ", " & objDo.SafeSQL(row("CustRef")) & ", " & CStr(row("discount")) & ", " & CStr(row("netamt")) & "," & CStr(row("gst")) & "," & CStr(row("amount")) & "," & objDo.SafeSQL(row("Terms")) & "," & CStr(row("TermDays")) & "," & IIf(CBool(row("Delivered")) = True, 1, 0) & "," & objDo.SafeSQL(row("CurCode")) & "," & CStr(row("CurExRate")) & "," & objDo.SafeSQL(row("CustomerID")) & "," & objDo.SafeSQL(Format(row("DTG"), "yyyyMMdd HH:mm:ss")) & "," & objDo.SafeSQL(row("AgentId")) & ", 0,0," & IIf(row("Void"), 1, 0) & "," & objDo.SafeSQL(Format(row("VoidDate"), "yyyyMMdd HH:mm:ss")) & "," & objDo.SafeSQL(Format(row("DeliDate"), "yyyyMMdd HH:mm:ss")) & ", " & objDo.SafeSQL(row("AcBillRef")) & "," & row("DiscountPer").ToString & "," & objDo.SafeSQL(row("CompanyName")) & "," & objDo.SafeSQL(row("Remarks")) & "," & objDo.SafeSQL(row("ShipName").ToString) & "," & objDo.SafeSQL(row("ShipAdd").ToString) & "," & objDo.SafeSQL(row("ShipAdd2").ToString) & "," & objDo.SafeSQL(row("ShipAdd3").ToString) & "," & objDo.SafeSQL(row("ShipAdd4").ToString) & "," & objDo.SafeSQL(row("ShipPin").ToString) & "," & objDo.SafeSQL(row("ShipCity").ToString) & "," & objDo.SafeSQL(sMDTNo) & "," & iConsign & ",@Signature)"
                End If

                If bImageExists = False Then
                    cmd.CommandText = cmd.CommandText.Replace("@Signature", "null")
                End If

                cmd.ExecuteNonQuery()

                'If iConsign = 1 Then
                '    objDo.ExecuteSQL("Insert into OrderHdr(OrdNo, OrdDt, PONo, Discount, SubTotal, GstAmt, TotalAmt, PayTerms, TermDays, Delivered, CurCode, CurExRate, CustId, DTG, AgentID, Exported, ExportMYOB, Void,VoidDate,DeliveryDate,AcBillRef,DisPer,CompanyName,Remarks, ShipName, ShipAdd, ShipAdd2, ShipAdd3, ShipAdd4, ShipPin, ShipCity, MDTNo, Consignment) values (" & objDo.SafeSQL(row("OrderNo")) & ", " & objDo.SafeSQL(Format(row("OrdDt"), "yyyyMMdd HH:mm:ss")) & ", " & objDo.SafeSQL(row("CustRef")) & ", " & CStr(row("discount")) & ", " & CStr(row("netamt")) & "," & CStr(row("gst")) & "," & CStr(row("amount")) & "," & objDo.SafeSQL(row("Terms")) & "," & CStr(row("TermDays")) & "," & IIf(CBool(row("Delivered")) = True, 1, 0) & "," & objDo.SafeSQL(row("CurCode")) & "," & CStr(row("CurExRate")) & "," & objDo.SafeSQL(row("CustomerID")) & "," & objDo.SafeSQL(Format(row("DTG"), "yyyyMMdd HH:mm:ss")) & "," & objDo.SafeSQL(row("AgentId")) & ", 1,0," & IIf(row("Void"), 1, 0) & "," & objDo.SafeSQL(Format(row("VoidDate"), "yyyyMMdd HH:mm:ss")) & "," & objDo.SafeSQL(Format(row("DeliDate"), "yyyyMMdd HH:mm:ss")) & ", " & objDo.SafeSQL(row("AcBillRef")) & "," & row("DiscountPer").ToString & "," & objDo.SafeSQL(row("CompanyName")) & "," & objDo.SafeSQL(row("Remarks")) & "," & objDo.SafeSQL(row("ShipName").ToString) & "," & objDo.SafeSQL(row("ShipAdd").ToString) & "," & objDo.SafeSQL(row("ShipAdd2").ToString) & "," & objDo.SafeSQL(row("ShipAdd3").ToString) & "," & objDo.SafeSQL(row("ShipAdd4").ToString) & "," & objDo.SafeSQL(row("ShipPin").ToString) & "," & objDo.SafeSQL(row("ShipCity").ToString) & "," & objDo.SafeSQL(sMDTNo) & "," & iConsign & ")")
                'Else
                '    objDo.ExecuteSQL("Insert into OrderHdr(OrdNo, OrdDt, PONo, Discount, SubTotal, GstAmt, TotalAmt, PayTerms, TermDays, Delivered, CurCode, CurExRate, CustId, DTG, AgentID, Exported, ExportMYOB, Void,VoidDate,DeliveryDate,AcBillRef,DisPer,CompanyName,Remarks, ShipName, ShipAdd, ShipAdd2, ShipAdd3, ShipAdd4, ShipPin, ShipCity, MDTNo, Consignment) values (" & objDo.SafeSQL(row("OrderNo")) & ", " & objDo.SafeSQL(Format(row("OrdDt"), "yyyyMMdd HH:mm:ss")) & ", " & objDo.SafeSQL(row("CustRef")) & ", " & CStr(row("discount")) & ", " & CStr(row("netamt")) & "," & CStr(row("gst")) & "," & CStr(row("amount")) & "," & objDo.SafeSQL(row("Terms")) & "," & CStr(row("TermDays")) & "," & IIf(CBool(row("Delivered")) = True, 1, 0) & "," & objDo.SafeSQL(row("CurCode")) & "," & CStr(row("CurExRate")) & "," & objDo.SafeSQL(row("CustomerID")) & "," & objDo.SafeSQL(Format(row("DTG"), "yyyyMMdd HH:mm:ss")) & "," & objDo.SafeSQL(row("AgentId")) & ", 0,0," & IIf(row("Void"), 1, 0) & "," & objDo.SafeSQL(Format(row("VoidDate"), "yyyyMMdd HH:mm:ss")) & "," & objDo.SafeSQL(Format(row("DeliDate"), "yyyyMMdd HH:mm:ss")) & ", " & objDo.SafeSQL(row("AcBillRef")) & "," & row("DiscountPer").ToString & "," & objDo.SafeSQL(row("CompanyName")) & "," & objDo.SafeSQL(row("Remarks")) & "," & objDo.SafeSQL(row("ShipName").ToString) & "," & objDo.SafeSQL(row("ShipAdd").ToString) & "," & objDo.SafeSQL(row("ShipAdd2").ToString) & "," & objDo.SafeSQL(row("ShipAdd3").ToString) & "," & objDo.SafeSQL(row("ShipAdd4").ToString) & "," & objDo.SafeSQL(row("ShipPin").ToString) & "," & objDo.SafeSQL(row("ShipCity").ToString) & "," & objDo.SafeSQL(sMDTNo) & "," & iConsign & ")")
                'End If

            End If
        Next row
        cmd.Dispose()
    End Sub

    Private Sub DownloadRtnReject(ByVal FileName As String, ByVal XsdFileName As String)
        Dim ds As New DataSet
        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Exit Sub
        ds.ReadXmlSchema(XsdFileName)
        ds.ReadXml(FileName)
        For Each row As DataRow In ds.Tables(0).Rows
            For Each column As DataColumn In ds.Tables(0).Columns
                If row(column) Is DBNull.Value Then
                    If column.DataType.ToString = "System.String" Then
                        row(column) = ""
                    ElseIf column.DataType.ToString = "System.DateTime" Then
                        row(column) = CDate("1/1/2000")
                    Else
                        row(column) = 0
                    End If
                End If
            Next

            If IsDBNull(row("Accepted")) = False Then
                If CBool(row("Accepted")) = True Then
                    ImgCount += 1
                    ReDim Preserve arrImages(ImgCount)
                    arrImages(ImgCount) = row("RtnRejectNo") & ".bmp"
                End If
            End If

            'If CheckAvailable("Select * from RtnReject where RtnRejectNo = " & objDo.SafeSQL(row("RtnRejectNo").ToString) & " and CustId=" & objDo.SafeSQL(row("CustomerID").ToString)) = True Then
            objDo.ExecuteSQL("Update RtnReject Set Accepted = " & objDo.SafeSQL(row("Accepted").ToString) & " where RtnRejectNo = " & objDo.SafeSQL(row("RtnRejectNo").ToString) & " and CustId=" & objDo.SafeSQL(row("CustomerID").ToString))
            'Else
            'objDo.ExecuteSQL("Insert into RtnReject(RtnRejectNo, RtnRejectDate, RtnNo, CustID, AgentID, Accepted) values (" & objDo.SafeSQL(row("RtnRejectNo")) & ", " & objDo.SafeSQL(Format(row("RtnRejectDate"), "yyyyMMdd HH:mm:ss")) & ", " & objDo.SafeSQL(row("RtnNo")) & ", " & objDo.SafeSQL(row("CustomerID")) & "," & objDo.SafeSQL(row("AgentID")) & "," & objDo.SafeSQL(row("Accepted")) & ")")
            'End If
        Next row
    End Sub

    Private Sub DownloadOrdItem(ByVal FileName As String, ByVal XsdFileName As String)
        Dim ds As New DataSet
        Dim CommandSQL As String = ""
        Dim ValueSQL As String = ""
        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Exit Sub
        ds.ReadXmlSchema(XsdFileName)
        ds.ReadXml(FileName)
        For Each row As DataRow In ds.Tables(0).Rows
            For Each str As String In arrDet
                If row("OrderNo").ToString = str Then
                    'Insert Command
                    For Each column As DataColumn In ds.Tables(0).Columns
                        If row(column) Is DBNull.Value Then
                            If column.DataType.ToString = "System.String" Then
                                row(column) = ""
                            ElseIf column.DataType.ToString = "System.DateTime" Then
                                row(column) = CDate("1/1/2000")
                            Else
                                row(column) = 0
                            End If
                        End If
                    Next
                    objDo.ExecuteSQL("Insert into OrdItem(OrdNo, ItemNo, UOM, Qty, Price, SubAmt, Remarks, Description, DisPer, DisPr, PromoID, PromoOffer, ActPrice, ReasonCode, ExpiryDate, ColorRemarks) values (" & IIf(row("OrderNo") Is DBNull.Value, "Null", objDo.SafeSQL(row("OrderNo"))) & "," & IIf(row("ItemId") Is DBNull.Value, "Null", objDo.SafeSQL(row("Itemid"))) & "," & IIf(row("UOM") Is DBNull.Value, "Null", objDo.SafeSQL(row("uom"))) & "," & IIf(row("Qty") Is DBNull.Value, "Null", CStr(row("qty"))) & "," & IIf(row("Price") Is DBNull.Value, "Null", CStr(row("Price"))) & "," & IIf(row("Amount") Is DBNull.Value, "Null", CStr(row("amount"))) & "," & objDo.SafeSQL(row("Remarks")) & "," & objDo.SafeSQL(row("ItemName")) & "," & IIf(row("DisPer") Is DBNull.Value, "Null", CStr(row("DisPer"))) & "," & IIf(row("DisPr") Is DBNull.Value, "Null", CStr(row("DisPr"))) & "," & objDo.SafeSQL(row("PromoID").ToString) & "," & objDo.SafeSQL(row("PromoOffer").ToString) & "," & IIf(row("ActPrice") Is DBNull.Value, "Null", CStr(row("ActPrice"))) & "," & IIf(row("ReasonCode") Is DBNull.Value, "Null", objDo.SafeSQL(row("ReasonCode")) & "," & objDo.SafeSQL(Format(row("ExpDate"), "yyyyMMdd HH:mm:ss")) & "," & objDo.SafeSQL(row("ColorRemarks"))) & ")")
                    Exit For
                End If
            Next
        Next row
        Dim i As Integer
        For i = 0 To arrOrdNo.Count - 1
            objDo.ExecuteSQL("Update OrdItem Set OrdNo=" & objDo.SafeSQL(arrOrdNo(i).ToString & "T") & " where OrdNo = " & objDo.SafeSQL(arrOrdNo(i).ToString))
            objDo.ExecuteSQL("Update OrderHdr Set OrdNo=" & objDo.SafeSQL(arrOrdNo(i).ToString & "T") & " where OrdNo = " & objDo.SafeSQL(arrOrdNo(i).ToString))
        Next
    End Sub

    Private Sub DownloadInvoices(ByVal FileName As String, ByVal XsdFileName As String)
        Dim ds As New DataSet
        Dim CommandSQL As String = ""
        Dim ValueSQL As String = ""
        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Exit Sub
        ds.ReadXmlSchema(XsdFileName)
        ds.ReadXml(FileName)
        arrDet = New ArrayList
        For Each row As DataRow In ds.Tables(0).Rows
            For Each column As DataColumn In ds.Tables(0).Columns
                If row(column) Is DBNull.Value Then
                    If column.DataType.ToString = "System.String" Then
                        row(column) = ""
                    ElseIf column.DataType.ToString = "System.DateTime" Then
                        row(column) = CDate("1/1/2000")
                    Else
                        row(column) = 0
                    End If
                End If
            Next
            If CheckAvailable("Select * from Invoice where InvNo = " & objDo.SafeSQL(row("InvNo").ToString)) = True Then
                Dim dPaidAmt As Double = 0
                If IsInvoiceExists(row("InvNo").ToString, dPaidAmt) Then
                    'If CheckAvailable("Select * from Invoice where InvNo = " & objDo.SafeSQL(row("InvNo").ToString) & " and DATEDIFF(s, DTG, " & objDo.SafeSQL(Format(row("DTG"), "yyyyMMdd HH:mm:ss")) & ") > 0") = True Then
                    'arrDet.Add(row("InvNo").ToString)
                    'objDo.ExecuteSQL("Delete from InvItem where InvNo = " & objDo.SafeSQL(row("InvNo").ToString))
                    'objDo.ExecuteSQL("Delete from Invoice where InvNo = " & objDo.SafeSQL(row("InvNo").ToString))
                    If dPaidAmt < row("PaidAmt") Then
                        objDo.ExecuteSQL("Update Invoice set Void = " & IIf(row("Void"), 1, 0) & ", PaidAmt = " & row("PaidAmt") & " where InvNo = " & objDo.SafeSQL(row("InvNo").ToString))
                    End If
                    'GoTo InsertData
                End If
            Else
                arrDet.Add(row("InvNo").ToString)
InsertData:
                'Insert Command
                objDo.ExecuteSQL("Insert into Invoice(InvNo, InvDt, DoNo, DoDt, OrdNo, CustId, AgentId, Discount, SubTotal, GstAmt, TotalAmt, PayTerms, TermDays, CurCode, CurExRate, PaidAmt, DTG, PrintNo, Exported, Void, CustRefNo,GST, AcBillRef, DisPer, CompanyName) values (" & objDo.SafeSQL(row("InvNo")) & ", '" & Format(row("InvDt"), "yyyyMMdd HH:mm:ss") & "'," & objDo.SafeSQL(row("InvNo")) & ", '" & Format(row("InvDt"), "yyyyMMdd HH:mm:ss") & "'," & objDo.SafeSQL(row("OrderNo")) & "," & objDo.SafeSQL(row("CustomerID")) & "," & objDo.SafeSQL(row("AgentId")) & "," & row("Discount") & "," & row("netamt") & "," & row("gst") & "," & row("amount") & "," & objDo.SafeSQL(row("Terms")) & "," & row("TermDays") & "," & objDo.SafeSQL(row("CurCode")) & "," & row("CurExRate") & "," & row("PaidAmt") & "," & objDo.SafeSQL(Format(row("DTG"), "yyyyMMdd HH:mm:ss")) & "," & row("PrintNo") & ",0," & IIf(row("Void"), 1, 0) & ",''," & dGST & ", " & objDo.SafeSQL(row("AcBillRef")) & "," & row("DiscountPer").ToString & "," & objDo.SafeSQL(row("CompanyName")) & " )")
                objDo.ExecuteSQL("Update Invoice set CustRefNo = OrderHdr.PONo from OrderHdr where OrderHdr.OrdNo = Invoice.OrdNo")
            End If
        Next row
    End Sub

    Private Function IsInvoiceExists(ByVal sInvNo As String, ByRef dInvAmt As Double) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean = False
        dtr = objDo.ReadRecord("Select InvNo, PaidAmt from Invoice where InvNo = " & objDo.SafeSQL(sInvNo))
        If dtr.Read Then
            bAns = True
            dInvAmt = dtr("PaidAmt")
        End If
        dtr.Close()
        Return bAns
    End Function

    Private Function IsCrNoteExists(ByVal sInvNo As String, ByRef dInvAmt As Double) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean = False
        dtr = objDo.ReadRecord("Select CreditNoteNo, PaidAmt from CreditNote where CreditNoteNo = " & objDo.SafeSQL(sInvNo))
        If dtr.Read Then
            bAns = True
            dInvAmt = dtr("PaidAmt")
        End If
        dtr.Close()
        Return bAns
    End Function

    Private Sub DownloadInvItem(ByVal FileName As String, ByVal XsdFileName As String)
        Dim ds As New DataSet
        Dim CommandSQL As String = ""
        Dim ValueSQL As String = ""
        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Exit Sub
        ds.ReadXmlSchema(XsdFileName)
        ds.ReadXml(FileName)
        For Each row As DataRow In ds.Tables(0).Rows
            For Each str As String In arrDet
                If row("InvNo").ToString = str Then
                    'Insert Command
                    For Each column As DataColumn In ds.Tables(0).Columns
                        If row(column) Is DBNull.Value Then
                            If column.DataType.ToString = "System.String" Then
                                row(column) = ""
                            ElseIf column.DataType.ToString = "System.DateTime" Then
                                row(column) = CDate("1/1/2000")
                            Else
                                row(column) = 0
                            End If
                        End If
                    Next
                    objDo.ExecuteSQL("Insert into InvItem (InvNo, ItemNo, UOM, Qty, Price, SubAmt, DisPr, DisPer, Remarks, Description, ReasonCode,ColorRemarks) Values (" & objDo.SafeSQL(row("InvNo")) & "," & objDo.SafeSQL(row("Itemid")) & "," & objDo.SafeSQL(row("uom")) & "," & row("qty") & "," & row("price") & "," & row("amount") & ", 0, 0" & "," & objDo.SafeSQL(row("Remarks")) & "," & objDo.SafeSQL(row("ItemName")) & "," & IIf(row("ReasonCode") Is DBNull.Value, "Null", objDo.SafeSQL(row("ReasonCode")) & "," & objDo.SafeSQL(row("ColorRemarks"))) & ")")
                    Exit For
                End If
            Next
        Next row
    End Sub

    Private Sub DownloadReceipt(ByVal FileName As String, ByVal XsdFileName As String)
        Dim ds As New DataSet
        Dim CommandSQL As String = ""
        Dim ValueSQL As String = ""
        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Exit Sub
        ds.ReadXmlSchema(XsdFileName)
        ds.ReadXml(FileName)
        arrDet = New ArrayList
        For Each row As DataRow In ds.Tables(0).Rows
            For Each column As DataColumn In ds.Tables(0).Columns
                If row(column) Is DBNull.Value Then
                    If column.DataType.ToString = "System.String" Then
                        row(column) = ""
                    ElseIf column.DataType.ToString = "System.DateTime" Then
                        row(column) = CDate("1/1/2000")
                    Else
                        row(column) = 0
                    End If
                End If
            Next
            If CheckAvailable("Select * from Receipt where RcptNo = " & objDo.SafeSQL(row("ReceiptNo").ToString)) = True Then
                If CheckAvailable("Select * from Receipt where RcptNo = " & objDo.SafeSQL(row("ReceiptNo").ToString) & " and DATEDIFF(s, DTG, " & objDo.SafeSQL(Format(row("DTG"), "yyyyMMdd HH:mm:ss")) & ") > 0") = True Then
                    arrDet.Add(row("ReceiptNo").ToString)
                    objDo.ExecuteSQL("Delete from RcptItem where RcptNo = " & objDo.SafeSQL(row("ReceiptNo").ToString))
                    objDo.ExecuteSQL("Delete from Receipt where RcptNo = " & objDo.SafeSQL(row("ReceiptNo").ToString))
                    'Update Command
                    GoTo InsertData
                End If
            Else
                arrDet.Add(row("ReceiptNo").ToString)
InsertData:
                Dim sBank As String = ""
                If Trim(row("ExtDocNo").ToString) <> "" Then
                    sBank = row("BankName").ToString
                End If
                If Trim(row("PaymentMode").ToString) = "CASH" Then
                    objDo.ExecuteSQL("Insert into Receipt (RcptNo, RcptDt, Amount, PayMethod, CurCode, CurExRate, CustID, AgentId, ChqNo, ChqDt, BankName, DTG, Exported, Void, CompanyName, ChqType, ExportNav, Remarks) values (" & objDo.SafeSQL(row("ReceiptNo")) & ", '" & Format(row("receiptDt"), "yyyyMMdd HH:mm:ss") & "'," & row("amount") & "," & objDo.SafeSQL(row("PaymentMode")) & "," & objDo.SafeSQL(row("CurCode")) & "," & row("CurExRate") & "," & objDo.SafeSQL(row("CustomerID")) & "," & objDo.SafeSQL(row("AgentID")) & "," & objDo.SafeSQL(row("ExtDocNo")) & "," & objDo.SafeSQL(Format(row("ChqDt"), "yyyyMMdd HH:mm:ss")) & "," & objDo.SafeSQL(sBank) & "," & objDo.SafeSQL(Format(row("DTG"), "yyyyMMdd HH:mm:ss")) & ",0," & IIf(row("Void"), 1, 0) & "," & objDo.SafeSQL(row("CompanyName")) & "," & objDo.SafeSQL(row("ChqType")) & ",0," & objDo.SafeSQL(IIf(IsDBNull(row("Remarks")), "", row("Remarks"))) & ")") 'ChqType
                Else
                    objDo.ExecuteSQL("Insert into Receipt (RcptNo, RcptDt, Amount, PayMethod, CurCode, CurExRate, CustID, AgentId, ChqNo, ChqDt, BankName, DTG, Exported, Void, CompanyName, ChqType, ExportNav, Remarks) values (" & objDo.SafeSQL(row("ReceiptNo")) & ", '" & Format(row("receiptDt"), "yyyyMMdd HH:mm:ss") & "'," & row("amount") & "," & objDo.SafeSQL(row("PaymentMode")) & "," & objDo.SafeSQL(row("CurCode")) & "," & row("CurExRate") & "," & objDo.SafeSQL(row("CustomerID")) & "," & objDo.SafeSQL(row("AgentID")) & "," & objDo.SafeSQL(row("ExtDocNo")) & "," & objDo.SafeSQL(Format(row("ChqDt"), "yyyyMMdd HH:mm:ss")) & "," & objDo.SafeSQL(sBank) & "," & objDo.SafeSQL(Format(row("DTG"), "yyyyMMdd HH:mm:ss")) & ",1," & IIf(row("Void"), 1, 0) & "," & objDo.SafeSQL(row("CompanyName")) & "," & objDo.SafeSQL(row("ChqType")) & ",0," & objDo.SafeSQL(IIf(IsDBNull(row("Remarks")), "", row("Remarks"))) & ")") 'ChqType
                End If

                'UIC Change on 2/3/2007
                'objDo.ExecuteSQL("Insert into Receipt (RcptNo, RcptDt, Amount, PayMethod, CurCode, CurExRate, CustID, AgentId, ChqNo, ChqDt, BankName, DTG, Exported, Void, CompanyName, ChqType) values (" & objDo.SafeSQL(row("ReceiptNo")) & ", '" & Format(row("receiptDt"), "yyyyMMdd HH:mm:ss") & "'," & row("amount") & "," & objDo.SafeSQL(row("PaymentMode")) & "," & objDo.SafeSQL(row("CurCode")) & "," & row("CurExRate") & "," & objDo.SafeSQL(row("CustomerID")) & "," & objDo.SafeSQL(row("AgentID")) & "," & objDo.SafeSQL(row("ExtDocNo")) & "," & objDo.SafeSQL(Format(row("ChqDt"), "yyyyMMdd HH:mm:ss")) & "," & objDo.SafeSQL(sBank) & "," & objDo.SafeSQL(Format(row("DTG"), "yyyyMMdd HH:mm:ss")) & ",0," & IIf(row("Void"), 1, 0) & "," & objDo.SafeSQL(row("CompanyName")) & "," & objDo.SafeSQL(row("ChqType")) & ")") 'ChqType
            End If
        Next row
    End Sub
    Private Sub DownloadExchange(ByVal FileName As String, ByVal XsdFileName As String)
        Dim ds As New DataSet
        Dim CommandSQL As String = ""
        Dim ValueSQL As String = ""
        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Exit Sub
        ds.ReadXmlSchema(XsdFileName)
        ds.ReadXml(FileName)
        arrDet = New ArrayList
        For Each row As DataRow In ds.Tables(0).Rows
            For Each column As DataColumn In ds.Tables(0).Columns
                If row(column) Is DBNull.Value Then
                    If column.DataType.ToString = "System.String" Then
                        row(column) = ""
                    ElseIf column.DataType.ToString = "System.DateTime" Then
                        row(column) = CDate("1/1/2000")
                    Else
                        row(column) = 0
                    End If
                End If
            Next
            If CheckAvailable("Select * from GoodsExchange where ExchangeNo = " & objDo.SafeSQL(row("ExNo").ToString)) = True Then
                '   If CheckAvailable("Select * from Receipt where RcptNo = " & objDo.SafeSQL(row("ReceiptNo").ToString) & " and DATEDIFF(s, DTG, " & objDo.SafeSQL(Format(row("DTG"), "yyyyMMdd HH:mm:ss")) & ") > 0") = True Then
                arrDet.Add(row("ExNo").ToString)
                objDo.ExecuteSQL("Delete from GoodsExchange where ExchangeNo = " & objDo.SafeSQL(row("ExNo").ToString))
                objDo.ExecuteSQL("Delete from GoodsExchangeItem where ExchangeNo = " & objDo.SafeSQL(row("ExNo").ToString))
                'Update Command
                GoTo InsertData
                'End If
            Else
                arrDet.Add(row("ExNo").ToString)
InsertData:
                'Insert Command
                objDo.ExecuteSQL("Insert into GoodsExchange (ExchangeNo, ExchangeDate, CustNo, SalesPersonCode, Exported, CompanyName) values (" & objDo.SafeSQL(row("ExNo")) & ", '" & Format(row("ExDate"), "yyyyMMdd HH:mm:ss") & "'," & objDo.SafeSQL(row("CustomerID")) & "," & objDo.SafeSQL(row("AgentID")) & "," & 0 & "," & objDo.SafeSQL(row("CompanyName")) & ")")
            End If
        Next row
    End Sub
    Private Sub DownloadGoodsInvn(ByVal FileName As String, ByVal XsdFileName As String)
        Dim ds As New DataSet
        Dim CommandSQL As String = ""
        Dim ValueSQL As String = ""
        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Exit Sub
        ds.ReadXmlSchema(XsdFileName)
        ds.ReadXml(FileName)
        arrDet = New ArrayList
        objDo.ExecuteSQL("Delete from GoodsInvn Where Location = " & objDo.SafeSQL(sLocation))
        For Each row As DataRow In ds.Tables(0).Rows
            For Each column As DataColumn In ds.Tables(0).Columns
                If row(column) Is DBNull.Value Then
                    If column.DataType.ToString = "System.String" Then
                        row(column) = ""
                    ElseIf column.DataType.ToString = "System.DateTime" Then
                        row(column) = CDate("1/1/2000")
                    Else
                        row(column) = 0
                    End If
                End If
            Next
            objDo.ExecuteSQL("Insert into GoodsInvn (Location, ItemNo, UOM, Qty, BadQty) values (" & objDo.SafeSQL(sLocation) & ", " & objDo.SafeSQL(row("ItemID")) & "," & objDo.SafeSQL(row("UOM")) & "," & row("Qty").ToString & "," & row("BadQty").ToString & ")")
        Next row
    End Sub

    Private Sub DownloadConsignInvn(ByVal FileName As String, ByVal XsdFileName As String)
        Dim ds As New DataSet
        Dim CommandSQL As String = ""
        Dim ValueSQL As String = ""
        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Exit Sub
        ds.ReadXmlSchema(XsdFileName)
        ds.ReadXml(FileName)
        arrDet = New ArrayList

        For Each row As DataRow In ds.Tables(0).Rows
            For Each column As DataColumn In ds.Tables(0).Columns
                If row(column) Is DBNull.Value Then
                    If column.DataType.ToString = "System.String" Then
                        row(column) = ""
                    ElseIf column.DataType.ToString = "System.DateTime" Then
                        row(column) = CDate("1/1/2000")
                    Else
                        row(column) = 0
                    End If
                End If
            Next
            If CheckAvailable("Select * from GoodsInvn where Location = " & objDo.SafeSQL(sLocation) & " and ItemNo = " & objDo.SafeSQL(row("ItemID"))) = True Then
                objDo.ExecuteSQL("Update GoodsInvn Set Qty = Qty - " & row("Qty").ToString & ", BadQty = BadQty + " & row("BadQty").ToString & " where Location = " & objDo.SafeSQL(sLocation) & " and ItemNo = " & objDo.SafeSQL(row("ItemID")))
            Else
                objDo.ExecuteSQL("Insert into GoodsInvn (Location, ItemNo, UOM, Qty, BadQty) values (" & objDo.SafeSQL(sLocation) & ", " & objDo.SafeSQL(row("ItemID")) & "," & objDo.SafeSQL(row("UOM")) & "," & row("Qty").ToString & "," & row("BadQty").ToString & ")")
            End If
            If CheckAvailable("Select * from GoodsInvn where Location = " & objDo.SafeSQL(row("Location").ToString) & " and ItemNo = " & objDo.SafeSQL(row("ItemID"))) = True Then
                objDo.ExecuteSQL("Update GoodsInvn Set Qty = Qty + " & row("Qty").ToString & ", BadQty = BadQty - " & row("BadQty").ToString & " where Location = " & objDo.SafeSQL(row("Location").ToString) & " and ItemNo = " & objDo.SafeSQL(row("ItemID")))
            Else
                objDo.ExecuteSQL("Insert into GoodsInvn (Location, ItemNo, UOM, Qty, BadQty) values (" & objDo.SafeSQL(row("Location").ToString) & ", " & objDo.SafeSQL(row("ItemID")) & "," & objDo.SafeSQL(row("UOM")) & "," & row("Qty").ToString & "," & row("BadQty").ToString & ")")
            End If
        Next row
    End Sub

    Private Sub DownloadExchangeItem(ByVal FileName As String, ByVal XsdFileName As String)
        Dim ds As New DataSet
        Dim CommandSQL As String = ""
        Dim ValueSQL As String = ""
        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Exit Sub
        ds.ReadXmlSchema(XsdFileName)
        ds.ReadXml(FileName)
        For Each row As DataRow In ds.Tables(0).Rows
            For Each str As String In arrDet
                If row("ExNo").ToString = str Then
                    'Insert Command
                    For Each column As DataColumn In ds.Tables(0).Columns
                        If row(column) Is DBNull.Value Then
                            If column.DataType.ToString = "System.String" Then
                                row(column) = ""
                            ElseIf column.DataType.ToString = "System.DateTime" Then
                                row(column) = CDate("1/1/2000")
                            Else
                                row(column) = 0
                            End If
                        End If
                    Next
                    objDo.ExecuteSQL("Insert into GoodsExchangeItem( ExchangeNo, ItemNo, UOM, Quantity, Remarks) Values (" & objDo.SafeSQL(row("ExNo")) & "," & objDo.SafeSQL(row("ItemId")) & "," & objDo.SafeSQL(row("uom")) & "," & row("qty") & "," & objDo.SafeSQL(row("Remarks")) & ")")
                    Exit For
                End If
            Next
        Next row
    End Sub
    '    Private Sub DownloadReturn(ByVal FileName As String, ByVal XsdFileName As String)
    '        Dim ds As New DataSet
    '        Dim CommandSQL As String = ""
    '        Dim ValueSQL As String = ""
    '        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Exit Sub
    '        ds.ReadXmlSchema(XsdFileName)
    '        ds.ReadXml(FileName)
    '        arrDet = New ArrayList
    '        For Each row As DataRow In ds.Tables(0).Rows
    '            For Each column As DataColumn In ds.Tables(0).Columns
    '                If row(column) Is DBNull.Value Then
    '                    If column.DataType.ToString = "System.String" Then
    '                        row(column) = ""
    '                    ElseIf column.DataType.ToString = "System.DateTime" Then
    '                        row(column) = CDate("1/1/2000")
    '                    Else
    '                        row(column) = 0
    '                    End If
    '                End If
    '            Next
    '            If CheckAvailable("Select * from GoodsReturn where ReturnNo = " & objDo.SafeSQL(row("RtnNo").ToString)) = True Then
    '                '   If CheckAvailable("Select * from Receipt where RcptNo = " & objDo.SafeSQL(row("ReceiptNo").ToString) & " and DATEDIFF(s, DTG, " & objDo.SafeSQL(Format(row("DTG"), "yyyyMMdd HH:mm:ss")) & ") > 0") = True Then
    '                arrDet.Add(row("RtnNo").ToString)
    '                objDo.ExecuteSQL("Delete from GoodsReturn where ReturnNo = " & objDo.SafeSQL(row("RtnNo").ToString))
    '                objDo.ExecuteSQL("Delete from GoodsReturnItem where ReturnNo = " & objDo.SafeSQL(row("RtnNo").ToString))
    '                'Update Command
    '                GoTo InsertData
    '                'End If
    '            Else
    '                arrDet.Add(row("RtnNo").ToString)
    'InsertData:
    '                'Insert Command
    '                objDo.ExecuteSQL("Insert into GoodsReturn (ReturnNo, ReturnDate, CustNo, SalesPersonCode, Exported) values (" & objDo.SafeSQL(row("RtnNo")) & ", '" & Format(row("RtnDate"), "yyyyMMdd HH:mm:ss") & "'," & objDo.SafeSQL(row("CustomerID")) & "," & objDo.SafeSQL(row("AgentID")) & "," & 0 & ")")
    '            End If
    '        Next row
    '    End Sub
    '    Private Sub DownloadReturnItem(ByVal FileName As String, ByVal XsdFileName As String)
    '        Dim ds As New DataSet
    '        Dim CommandSQL As String = ""
    '        Dim ValueSQL As String = ""
    '        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Exit Sub
    '        ds.ReadXmlSchema(XsdFileName)
    '        ds.ReadXml(FileName)
    '        For Each row As DataRow In ds.Tables(0).Rows
    '            For Each str As String In arrDet
    '                If row("RtnNo").ToString = str Then
    '                    'Insert Command
    '                    For Each column As DataColumn In ds.Tables(0).Columns
    '                        If row(column) Is DBNull.Value Then
    '                            If column.DataType.ToString = "System.String" Then
    '                                row(column) = ""
    '                            ElseIf column.DataType.ToString = "System.DateTime" Then
    '                                row(column) = CDate("1/1/2000")
    '                            Else
    '                                row(column) = 0
    '                            End If
    '                        End If
    '                    Next
    '                    objDo.ExecuteSQL("Insert into GoodsReturnItem( ReturnNo, ItemNo, UOM, Quantity, Remarks) Values (" & objDo.SafeSQL(row("RtnNo")) & "," & objDo.SafeSQL(row("ItemId")) & "," & objDo.SafeSQL(row("uom")) & "," & row("qty") & "," & objDo.SafeSQL(row("Remarks")) & ")")
    '                    Exit For
    '                End If
    '            Next
    '        Next row
    '    End Sub

    Private Sub DownloadReturn(ByVal FileName As String, ByVal XsdFileName As String)
        Dim ds As New DataSet
        Dim CommandSQL As String = ""
        Dim ValueSQL As String = ""
        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Exit Sub
        ds.ReadXmlSchema(XsdFileName)
        ds.ReadXml(FileName)
        arrDet = New ArrayList
        For Each row As DataRow In ds.Tables(0).Rows
            For Each column As DataColumn In ds.Tables(0).Columns
                If row(column) Is DBNull.Value Then
                    If column.DataType.ToString = "System.String" Then
                        row(column) = ""
                    ElseIf column.DataType.ToString = "System.DateTime" Then
                        row(column) = CDate("1/1/2000")
                    Else
                        row(column) = 0
                    End If
                End If
            Next
            If CheckAvailable("Select * from GoodsReturn where ReturnNo = " & objDo.SafeSQL(row("RtnNo").ToString)) = True Then
                '   If CheckAvailable("Select * from Receipt where RcptNo = " & objDo.SafeSQL(row("ReceiptNo").ToString) & " and DATEDIFF(s, DTG, " & objDo.SafeSQL(Format(row("DTG"), "yyyyMMdd HH:mm:ss")) & ") > 0") = True Then
                arrDet.Add(row("RtnNo").ToString)
                objDo.ExecuteSQL("Delete from GoodsReturn where ReturnNo = " & objDo.SafeSQL(row("RtnNo").ToString))
                objDo.ExecuteSQL("Delete from GoodsReturnItem where ReturnNo = " & objDo.SafeSQL(row("RtnNo").ToString))
                'Update Command
                GoTo InsertData
                'End If
            Else
                arrDet.Add(row("RtnNo").ToString)
InsertData:
                'Insert Command
                objDo.ExecuteSQL("Insert into GoodsReturn (ReturnNo, ReturnDate, CustNo, SalesPersonCode, Discount, SubTotal, GST, TotalAmt, Void, VoidDate, Exported, CompanyName) values (" & objDo.SafeSQL(row("RtnNo")) & ", '" & Format(row("RtnDate"), "yyyyMMdd HH:mm:ss") & "'," & objDo.SafeSQL(row("CustomerID")) & "," & objDo.SafeSQL(row("AgentID")) & "," & objDo.SafeSQL(row("Discount")) & "," & objDo.SafeSQL(row("SubTotal")) & "," & objDo.SafeSQL(row("GST")) & "," & objDo.SafeSQL(row("Amount")) & "," & IIf(row("Void"), 1, 0) & "," & objDo.SafeSQL(Format(row("VoidDate"), "yyyyMMdd HH:mm:ss")) & "," & 0 & "," & objDo.SafeSQL(row("CompanyName")) & ")")
            End If
        Next row
    End Sub

    Private Sub DownloadReturnItem(ByVal FileName As String, ByVal XsdFileName As String)
        Dim ds As New DataSet
        Dim CommandSQL As String = ""
        Dim ValueSQL As String = ""
        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Exit Sub
        ds.ReadXmlSchema(XsdFileName)
        ds.ReadXml(FileName)
        For Each row As DataRow In ds.Tables(0).Rows
            For Each str As String In arrDet
                If row("RtnNo").ToString = str Then
                    'Insert Command
                    For Each column As DataColumn In ds.Tables(0).Columns
                        If row(column) Is DBNull.Value Then
                            If column.DataType.ToString = "System.String" Then
                                row(column) = ""
                            ElseIf column.DataType.ToString = "System.DateTime" Then
                                row(column) = CDate("1/1/2000")
                            Else
                                row(column) = 0
                            End If
                        End If
                    Next
                    objDo.ExecuteSQL("Insert into GoodsReturnItem(ReturnNo, ItemNo, UOM, Quantity, ReasonCode, Remarks, Price, Amt, InvNo) Values (" & objDo.SafeSQL(row("RtnNo")) & "," & objDo.SafeSQL(row("ItemId")) & "," & objDo.SafeSQL(row("uom")) & "," & row("qty") & "," & objDo.SafeSQL(row("ReasonCode")) & "," & objDo.SafeSQL(row("Remarks")) & "," & CStr(row("Price")) & "," & CStr(row("amount")) & "," & objDo.SafeSQL(row("InvNo")) & ")")
                    Exit For
                End If
            Next
        Next row
    End Sub

    Private Sub DownloadStockTakeHdr(ByVal FileName As String, ByVal XsdFileName As String)
        Dim ds As New DataSet
        Dim CommandSQL As String = ""
        Dim ValueSQL As String = ""
        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Exit Sub
        ds.ReadXmlSchema(XsdFileName)
        ds.ReadXml(FileName)
        arrDet = New ArrayList
        For Each row As DataRow In ds.Tables(0).Rows
            For Each column As DataColumn In ds.Tables(0).Columns
                If row(column) Is DBNull.Value Then
                    If column.DataType.ToString = "System.String" Then
                        row(column) = ""
                    ElseIf column.DataType.ToString = "System.DateTime" Then
                        row(column) = CDate("1/1/2000")
                    Else
                        row(column) = 0
                    End If
                End If
            Next
            If CheckAvailable("Select * from StockTakeHdr where StockTakeNo = " & objDo.SafeSQL(row("StockTakeNo").ToString)) = True Then
                objDo.ExecuteSQL("Delete from StockTakeHdr where StockTakeNo = " & objDo.SafeSQL(row("StockTakeNo").ToString))
                objDo.ExecuteSQL("Delete from StockTakeDet where StockTakeNo = " & objDo.SafeSQL(row("StockTakeNo").ToString))
                GoTo InsertData
            Else
InsertData:
                objDo.ExecuteSQL("Insert into StockTakeHdr (StockTakeNo, StockTakeDate, LocationCode, AgentID) values (" & objDo.SafeSQL(row("StockTakeNo")) & ", '" & Format(row("StockTakeDate"), "yyyyMMdd HH:mm:ss") & "'," & objDo.SafeSQL(row("LocationCode")) & "," & objDo.SafeSQL(row("AgentID")) & ")")
            End If
        Next row
    End Sub

    Private Sub DownloadStockTakeDet(ByVal FileName As String, ByVal XsdFileName As String)
        Dim ds As New DataSet
        Dim CommandSQL As String = ""
        Dim ValueSQL As String = ""
        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Exit Sub
        ds.ReadXmlSchema(XsdFileName)
        ds.ReadXml(FileName)
        For Each row As DataRow In ds.Tables(0).Rows
            For Each column As DataColumn In ds.Tables(0).Columns
                If row(column) Is DBNull.Value Then
                    If column.DataType.ToString = "System.String" Then
                        row(column) = ""
                    ElseIf column.DataType.ToString = "System.DateTime" Then
                        row(column) = CDate("1/1/2000")
                    Else
                        row(column) = 0
                    End If
                End If
            Next
            objDo.ExecuteSQL("Insert into StockTakeDet(StockTakeNo, ItemId, UOM, Qty, DiffQty, BinNo, VariantCode, Remarks) Values (" & objDo.SafeSQL(row("StockTakeNo")) & "," & objDo.SafeSQL(row("ItemId")) & "," & objDo.SafeSQL(row("Uom")) & "," & row("Qty") & "," & row("DiffQty") & "," & objDo.SafeSQL(row("BinNo")) & "," & objDo.SafeSQL(row("VariantCode")) & "," & objDo.SafeSQL(row("Remarks")) & ")")
            'objDo.ExecuteSQL("Insert into StockTakeDet(StockTakeNo, ItemId, UOM, Qty, BinNo, VariantCode, Remarks) Values (" & objDo.SafeSQL(row("StockTakeNo")) & "," & objDo.SafeSQL(row("ItemId")) & "," & objDo.SafeSQL(row("Uom")) & "," & row("Qty") & "," & objDo.SafeSQL(row("BinNo")) & "," & objDo.SafeSQL(row("VariantCode")) & "," & objDo.SafeSQL(row("Remarks")) & ")")
        Next row
    End Sub

    Private Sub DownloadCreditNote(ByVal FileName As String, ByVal XsdFileName As String)
        Dim ds As New DataSet
        Dim CommandSQL As String = ""
        Dim ValueSQL As String = ""
        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Exit Sub
        ds.ReadXmlSchema(XsdFileName)
        ds.ReadXml(FileName)
        arrDet = New ArrayList
        For Each row As DataRow In ds.Tables(0).Rows
            For Each column As DataColumn In ds.Tables(0).Columns
                If row(column) Is DBNull.Value Then
                    If column.DataType.ToString = "System.String" Then
                        row(column) = ""
                    ElseIf column.DataType.ToString = "System.DateTime" Then
                        row(column) = CDate("1/1/2000")
                    Else
                        row(column) = 0
                    End If
                End If
            Next
            If CheckAvailable("Select * from CreditNote where CreditNoteNo = " & objDo.SafeSQL(row("CrNoteNo").ToString)) = True Then
                '   If CheckAvailable("Select * from Receipt where RcptNo = " & objDo.SafeSQL(row("ReceiptNo").ToString) & " and DATEDIFF(s, DTG, " & objDo.SafeSQL(Format(row("DTG"), "yyyyMMdd HH:mm:ss")) & ") > 0") = True Then
                'arrDet.Add(row("CrNoteNo").ToString)
                'objDo.ExecuteSQL("Delete from CreditNote where CreditNoteNo = " & objDo.SafeSQL(row("CrNoteNo").ToString))
                'objDo.ExecuteSQL("Delete from CreditNoteDet where CreditNoteNo = " & objDo.SafeSQL(row("CrNoteNo").ToString))
                'Update Command
                'GoTo InsertData
                'End If
                Dim dPaidAmt As Double = 0
                If IsCrNoteExists(row("CrNoteNo").ToString, dPaidAmt) Then
                    If dPaidAmt < row("PaidAmt") Then
                        objDo.ExecuteSQL("Update CreditNote Set PaidAmt = " & row("PaidAmt") & " where CreditNoteNo = " & objDo.SafeSQL(row("CrNoteNo")))
                    End If
                End If
            Else
                arrDet.Add(row("CrNoteNo").ToString)
InsertData:
                'Insert Command
                objDo.ExecuteSQL("Insert into CreditNote ( CreditNoteNo, CreditDate, CustNo, GoodsReturnNo, SalesPersonCode, Discount, SubTotal, GST, TotalAmt,DTG, Void, VoidDate, Exported, PaidAmt, PONo, CompanyName) values (" & objDo.SafeSQL(row("CrNoteNo")) & ", '" & Format(row("CrNoteDate"), "yyyyMMdd HH:mm:ss") & "'," & objDo.SafeSQL(row("CustomerID")) & "," & objDo.SafeSQL(row("RtnNo")) & "," & objDo.SafeSQL(row("AgentID")) & "," & objDo.SafeSQL(row("Discount")) & "," & objDo.SafeSQL(row("SubTotal")) & "," & objDo.SafeSQL(row("GST")) & "," & objDo.SafeSQL(row("Amount")) & "," & objDo.SafeSQL(Format(row("DTG"), "yyyyMMdd HH:mm:ss")) & "," & IIf(row("Void"), 1, 0) & "," & objDo.SafeSQL(Format(row("VoidDate"), "yyyyMMdd HH:mm:ss")) & ",0," & row("PaidAmt") & "," & objDo.SafeSQL(row("CustRtnNote").ToString) & "," & objDo.SafeSQL(row("CompanyName").ToString) & ")")
            End If
        Next row
    End Sub



    Private Sub DownloadCreditNoteDet(ByVal FileName As String, ByVal XsdFileName As String)
        Dim ds As New DataSet
        Dim CommandSQL As String = ""
        Dim ValueSQL As String = ""
        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Exit Sub
        ds.ReadXmlSchema(XsdFileName)
        ds.ReadXml(FileName)
        For Each row As DataRow In ds.Tables(0).Rows
            For Each str As String In arrDet
                If row("CrNoteNo").ToString = str Then
                    'Insert Command
                    For Each column As DataColumn In ds.Tables(0).Columns
                        If row(column) Is DBNull.Value Then
                            If column.DataType.ToString = "System.String" Then
                                row(column) = ""
                            ElseIf column.DataType.ToString = "System.DateTime" Then
                                row(column) = CDate("1/1/2000")
                            Else
                                row(column) = 0
                            End If
                        End If
                    Next
                    objDo.ExecuteSQL("Insert into CreditNoteDet(CreditNoteNo, ItemNo, Uom, Baseuom, Price, Qty, Amt, GoodQty, ReasonCode, Remarks, InvNo) values (" & objDo.SafeSQL(row("CrNoteNo")) & "," & objDo.SafeSQL(row("Itemid")) & "," & objDo.SafeSQL(row("uom")) & "," & objDo.SafeSQL(row("Baseuom")) & "," & CStr(row("Price")) & "," & CStr(row("qty")) & "," & CStr(row("amount")) & "," & CStr(row("GoodQty")) & "," & objDo.SafeSQL(row("ReasonCode")) & "," & objDo.SafeSQL(row("Remarks")) & "," & objDo.SafeSQL(row("InvNo")) & ")")
                    Exit For
                End If
            Next
        Next row
    End Sub

    Private Sub DownloadMDT(ByVal FileName As String, ByVal XsdFileName As String)
        Dim ds As New DataSet
        Dim CommandSQL As String = ""
        Dim ValueSQL As String = ""
        sLocation = ""
        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Exit Sub
        ds.ReadXmlSchema(XsdFileName)
        ds.ReadXml(FileName)
        For Each row As DataRow In ds.Tables(0).Rows
            'If CheckAvailable("Select * from MDT where RcptNo = " & objDo.SafeSQL(row("PDAID").ToString)) = True Then
            'Update Command
            sMDTNo = row("PDAID").ToString
            dGST = row("GST")
            sLocation = row("LOCATION").ToString
            sAgt = row("AgentID").ToString
            sAgentID = row("AgentID").ToString
            objDo.ExecuteSQL("Update MDT Set LastCustNo=" & row("LastCustNo") & "," & "LastOrdNo=" & row("LastOrdNo") & "," & "LastInvNo=" & row("LastInvNo") & "," & "LastRcptNo=" & row("LastRcptNo") & "," & "LastRtnNo=" & row("LastRetNo") & "," & "LastExNo=" & row("LastExNo") & "," & "LastITNo=" & row("LastITNo") & "," & "LastCRNo=" & row("LastCRNo") & "," & "LastSerNo=" & row("LastSerNo") & "," & "Printer=" & row("Printer") & "," & "PrintPort=" & row("PrintPort") & "," & "PrintBaud=" & row("PrintBaud") & "," & "DotPrintPort=" & row("DotPrintPort") & "," & "DotPrintBaud=" & row("DotPrintBaud") & " where MDTNo=" & objDo.SafeSQL(row("PDAID")) & " and CompanyName = " & objDo.SafeSQL(row("CompanyName")) & "")
            Dim dtrMDT As SqlDataReader
            dtrMDT = objDo.ReadRecord("Select AgentID, Location, VehicleId from MDT where MDTNo=" & objDo.SafeSQL(row("PDAID")))
            If dtrMDT.Read = True Then
                sVehicle = dtrMDT("VehicleId")
            End If
            dtrMDT.Close()
            'End If
        Next row

    End Sub

    Private Function GetMDTName(ByVal FileName As String, ByVal XsdFileName As String) As String
        Dim ds As New DataSet
        Dim CommandSQL As String = ""
        Dim ValueSQL As String = ""
        Dim sMDT As String = ""
        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Return ""
        ds.ReadXmlSchema(XsdFileName)
        ds.ReadXml(FileName)
        For Each row As DataRow In ds.Tables(0).Rows
            sMDT = Trim(row("PDAID"))
        Next row
        Return sMDT
    End Function


    Private Sub DownloadRcptItem(ByVal FileName As String, ByVal XsdFileName As String)
        Dim ds As New DataSet
        Dim CommandSQL As String = ""
        Dim ValueSQL As String = ""
        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Exit Sub
        ds.ReadXmlSchema(XsdFileName)
        ds.ReadXml(FileName)

        For Each row As DataRow In ds.Tables(0).Rows
            For Each str As String In arrDet
                If row("ReceiptNo").ToString = str Then
                    'Insert Command
                    For Each column As DataColumn In ds.Tables(0).Columns
                        If row(column) Is DBNull.Value Then
                            If column.DataType.ToString = "System.String" Then
                                row(column) = ""
                            ElseIf column.DataType.ToString = "System.DateTime" Then
                                row(column) = CDate("1/1/2000")
                            Else
                                row(column) = 0
                            End If
                        End If
                    Next
                    objDo.ExecuteSQL("Insert into RcptItem(RcptNo, InvNo, AmtPaid) values (" & objDo.SafeSQL(row("receiptNo")) & "," & objDo.SafeSQL(row("InvNo")) & "," & row("amount") & ")")
                    Exit For
                End If
            Next
        Next row
    End Sub

    Private Sub DownloadItemTrans(ByVal FileName As String, ByVal XsdFileName As String)
        Dim ds As New DataSet
        Dim CommandSQL As String = ""
        Dim ValueSQL As String = ""
        If System.IO.File.Exists(FileName) = False Or System.IO.File.Exists(XsdFileName) = False Then Exit Sub
        ds.ReadXmlSchema(XsdFileName)
        ds.ReadXml(FileName)
        'arrDet.Clear()
        For Each row As DataRow In ds.Tables(0).Rows
            'If CheckAvailable("Select * from ItemTrans where DocNo = " & objDo.SafeSQL(row("DocNo").ToString)) = True Then

            '    arrDet.Add(row("RcptNo").ToString)
            '    objDo.ExecuteSQL("Delete from ItemTrans where DocNo = " & objDo.SafeSQL(row("DocNo").ToString))
            '    'Update Command

            'Else
            '    arrDet.Add(row("RcptNo").ToString)
            '    'Insert Command

            For Each column As DataColumn In ds.Tables(0).Columns
                If row(column) Is DBNull.Value Then
                    If column.DataType.ToString = "System.String" Then
                        row(column) = ""
                    ElseIf column.DataType.ToString = "System.DateTime" Then
                        row(column) = CDate("1/1/2000")
                    Else
                        row(column) = 0
                    End If
                End If
            Next
            If CheckAvailable("Select * from ItemTrans where DocNo = " & objDo.SafeSQL(row("DocNo")) & " and DocType = " & objDo.SafeSQL(row("DocType")) & " and Location = " & objDo.SafeSQL(row("Location")) & " and ItemID = " & objDo.SafeSQL(row("ItemID")) & " and UOM = " & objDo.SafeSQL(row("UOM")) & " and Remarks = " & objDo.SafeSQL(row("Remarks"))) = True Then

            Else
                objDo.ExecuteSQL("Insert into ItemTrans( DocNo, DocDt, DocType, Location,ItemId, Uom, Qty, Remarks, Exported, DownloadDate) values (" & objDo.SafeSQL(row("DocNo")) & ", '" & Format(row("DocDate"), "yyyyMMdd HH:mm:ss") & "'," & objDo.SafeSQL(row("DocType")) & "," & objDo.SafeSQL(row("Location")) & "," & objDo.SafeSQL(row("ItemID")) & "," & objDo.SafeSQL(row("Uom")) & "," & row("qty") & "," & objDo.SafeSQL(row("Remarks")) & ",0," & Format(Date.Now, "yyyy-MM-dd") & ")")
                ''Updating Goods inventory for the van on the backend system
                ''VANINVN
                'If row("DocType") = "GIN" Or row("DocType") = "GVAR" Then
                '    objDo.ExecuteSQL("Update GoodsInvn Set Qty = Qty + " & row("Qty").ToString & " where Location = " & objDo.SafeSQL(row("Location")) & " and ItemNo = " & objDo.SafeSQL(row("ItemID").ToString))
                'ElseIf row("DocType") = "GOUT" Or row("DocType") = "DELI" Or row("DocType") = "RTN" Or row("DocType") = "EX" Then
                '    objDo.ExecuteSQL("Update GoodsInvn Set Qty = Qty - " & row("Qty").ToString & " where Location = " & objDo.SafeSQL(row("Location")) & " and ItemNo = " & objDo.SafeSQL(row("ItemID").ToString))
                'ElseIf row("DocType") = "STIN" Then
                '    objDo.ExecuteSQL("Update GoodsInvn Set Qty = Qty + " & row("Qty").ToString & " where Location = " & objDo.SafeSQL(row("Location")) & " and ItemNo = " & objDo.SafeSQL(row("ItemID").ToString))
                'ElseIf row("DocType") = "STOUT" Then
                '    objDo.ExecuteSQL("Update GoodsInvn Set Qty = Qty - " & row("Qty").ToString & " where Location = " & objDo.SafeSQL(row("Location")) & " and ItemNo = " & objDo.SafeSQL(row("ItemID").ToString))
                'End If
            End If

            'End If
        Next row
    End Sub
    Private Function CheckAvailable(ByVal SQL As String) As Boolean
        Dim dr As SqlClient.SqlDataReader
        Dim bCheck As Boolean = False
        dr = objDo.ReadRecord(SQL)
        bCheck = dr.Read
        dr.Close()
        Return bCheck
    End Function

    Private Sub btnSync_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSync.Click
        btnSync.Enabled = False
        If Download() = True Then
            If Upload() = True Then
                MsgBox("Completed")
            End If
        End If
        btnSync.Enabled = True
    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

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

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Dim frm As New Password()
        Dim dlgResult As DialogResult = frm.ShowDialog()
        If dlgResult = Windows.Forms.DialogResult.Cancel Then
            Exit Sub
        End If

        btnSync.Enabled = False
        btnUpload.Enabled = False
        If MsgBox("""Update"" will overwrite the data in the handheld. Must do the Synchronization First before Click ""Change User"". Do you want to continue?", MsgBoxStyle.YesNo, "Update?") = MsgBoxResult.Yes Then
            IsChangeUser = True
            If Upload() = True Then
                MsgBox("Completed")
            End If
            IsChangeUser = False
        End If
        btnSync.Enabled = True
        btnUpload.Enabled = True
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        objDo.ConnectDB()
        objDo.DisconnectAnotherDB()
        Dim dtr As SqlDataReader
        dtr = objDo.ReadRecord("Select ItemId, UOM, Qty , DocType, Location from ItemTrans Where doctype <> 'VANINVN'")
        While dtr.Read
            If dtr("DocType") = "GIN" Then
                objDo.ExecuteSQLAnother("Update GoodsInvn Set Qty = Qty + " & dtr("qty") & " where Location = " & objDo.SafeSQL(dtr("Location")) & " and ItemNo = " & objDo.SafeSQL(dtr("ItemId")))
            ElseIf dtr("DocType") = "GOUT" Or dtr("DocType") = "DELI" Then
                objDo.ExecuteSQLAnother("Update GoodsInvn Set Qty = Qty - " & dtr("qty") & " where Location = " & objDo.SafeSQL(dtr("Location")) & " and ItemNo = " & objDo.SafeSQL(dtr("ItemId")))
            ElseIf dtr("DocType") = "RTN" Then
                objDo.ExecuteSQLAnother("Update GoodsInvn Set BadQty = BadQty + " & dtr("qty") & " where Location = " & objDo.SafeSQL(dtr("Location")) & " and ItemNo = " & objDo.SafeSQL(dtr("ItemId")))
            End If
        End While
        dtr.Close()
        objDo.DisconnectAnotherDB()
        objDo.DisconnectDB()
        MsgBox("Completed")
    End Sub

    Private Sub btnUploadInvn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUploadInvn.Click

    End Sub

    Private Function UploadInventory() As Boolean
        sLocation = ""

        Dim bStatus As Boolean = True
        If dtup.Connect(sConStr) = False Then
            MsgBox("Please connect device and try again.")
            dtup.Disconnect()
            Return False
        End If
        objDo.ConnectDB()
        Dim arr As New ArrayList
        'objDo.ConnectAnotherDB()
        Dim frm As New status
        frm.Show()
        frm.Refresh()
        dtup.StartUpload()
        'New Code for Calculating Balance
        '''''
        '   Dim dtr As SqlDataReader
        dtup.StartUpload()
        Try
            frm.Refresh()
            Dim sMDT As String = ""
            If dtup.GetMDTFile = False Or IsChangeUser = True Then
NewMDT:
                frm.Visible = False
                frm.Refresh()
                'If MsgBox("Do you want to upload data for MDT " & dgvSync.Item(0, dgvSync.CurrentCell.RowIndex).Value, MsgBoxStyle.YesNo, "Download Data?") = MsgBoxResult.Yes Then
                sMDT = dgvSync.Item(0, dgvSync.CurrentCell.RowIndex).Value
                'sAgentID = dgvSync.Item(2, dgvSync.CurrentCell.RowIndex).Value
                'sLocation = dgvSync.Item(3, dgvSync.CurrentCell.RowIndex).Value
                Dim dtrMDT As SqlDataReader
                dtrMDT = objDo.ReadRecord("Select AgentID, Location, VehicleId, InvHistory from MDT where MDTNo=" & objDo.SafeSQL(sMDT))
                If dtrMDT.Read = True Then
                    sAgentID = dtrMDT("AgentID")
                    sLocation = dtrMDT("Location")
                    sVehicle = dtrMDT("VehicleId")
                    If IsDBNull(dtrMDT("InvHistory")) = False Then dInvHistory = dtrMDT("InvHistory")
                End If
                dtrMDT.Close()
                frm.Visible = True
                frm.Refresh()
                'Else
                '    dtup.EndUpload()
                '    dtup.Disconnect()
                '    objDo.DisconnectDB()
                '    frm.Close()
                '    Return False
                'End If
            Else
                Dim path As String
                path = System.AppDomain.CurrentDomain.BaseDirectory & "download\"
                sMDT = GetMDTName(path & "System.xml", path & "system.xsd")
                If sMDT = "" Then
                    GoTo NewMDT
                End If
                Dim dtrMDT As SqlDataReader
                dtrMDT = objDo.ReadRecord("Select AgentID, Location, VehicleId, InvHistory from MDT where MDTNo=" & objDo.SafeSQL(sMDT))
                If dtrMDT.Read = True Then
                    sAgentID = dtrMDT("AgentID")
                    sLocation = dtrMDT("Location")
                    sVehicle = dtrMDT("VehicleId")
                    If IsDBNull(dtrMDT("InvHistory")) = False Then dInvHistory = dtrMDT("InvHistory")
                End If
                dtrMDT.Close()
            End If
            objDo.ExecuteSQL("INSERT into SyncHistory(SyncTime, AgentId, Location) values (" & objDo.SafeSQL(Format(DateTime.Now, "yyyyMMdd HH:mm:ss")) & "," & objDo.SafeSQL(sAgentID) & "," & objDo.SafeSQL(sLocation) & ")")
            dtup.UploadFile("Select Item.ItemNo as itemid, uom, qty, 0 as BadQty from GoodsInvn, Item where GoodsInvn.ItemNo=Item.ItemNo and toPDA=1 and Location=" & objDo.SafeSQL(sLocation) & " union SELECT Item.ItemNo, Item.BaseUOM, 0 as Qty, 0 as BadQty FROM Item Where ToPDA=1 and (Item.ItemNo NOT IN (SELECT ItemNo FROM GoodsInvn WHERE Location =" & objDo.SafeSQL(sLocation) & "))", "VanInventory")
        Catch ex As Exception
            MsgBox(ex.Message)
            bStatus = False
        End Try
        dtup.EndUpload()
        dtup.Disconnect()
        objDo.DisconnectDB()
        frm.Close()
        Return bStatus
    End Function

 End Class
