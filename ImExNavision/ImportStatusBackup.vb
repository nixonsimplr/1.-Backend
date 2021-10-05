Imports System
Imports System.Resources
Imports System.Globalization
Imports System.Threading
Imports System.Reflection
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Imports System.Text.RegularExpressions
'Imports System.Data.Odbc

Public Class ImportStatusBackup
    Implements ISalesBase

    Private aAgent As New ArrayList()
    Dim arRem As New ArrayList
    Dim strPay As String = " "
    Dim strDate As String = " "
    Private sNavCompanyName As String = ""
    Private arrOrdNo As New ArrayList()
    Private arrRecNo As New ArrayList()
    Private arrCrNo As New ArrayList()
    Private arrNewCustNo As New ArrayList()
    Public sCustPostGroup, sGenPostGroup, sGSTPostGroup, sGSTProdGroup, sGenJournalTemplate, sGenJournalBatch, sWorkSheetTemplate, sJournalBatch, sItemJnlTemplate, sItemJnlBatch, sFocBatch, sExBatch, sBadLoc, sItemReclassTemplate, sItemReclassBatch As String
    Private Structure DelCust
        Dim CustID As String
        Dim PrGroup As String
    End Structure
    Dim i, igCnt As Integer
    'Dim cnt As Integer = 0
    Private NavCompanyName As String = GetCompanyName()
    Private Structure ArrRemarks
        Dim sOrdNo As String
        Dim sRemarks As String
    End Structure
    Private Structure ArrItemPrice
        Dim ItemCode As String
        Dim SalesCode As String
        Dim MaxDate As Date
        Dim SalesType As Integer
        Dim MinQty As Double
        Dim VariantCode As String
        Dim sUOM As String
    End Structure
    Private Structure ArrPurchaseOrder
        Dim PONo As String
        Dim PODt As Date
        Dim VendorId As String
        Dim AgentID As String
        Dim Discount As String
        Dim SubTotal As String
        Dim GSTAmt As String
        Dim TotalAmt As String
        Dim Void As String
        Dim PrintNo As String
        Dim PayTerms As String
        Dim CurCode As String
        Dim CurExRate As String
        Dim Exported As String
        Dim DTG As Date
        Dim DeliveryDate As Date
        Dim LocationCode As String
        Dim ExternalDocNo As String
        Dim Remarks As String
        Dim POType As String
        Dim ContainerNo As String
        Dim Department As String
        Dim ManufacturerCode As String
        Dim CompanyName As String
        Dim DocEntry As String
    End Structure
    Private Structure ArrDiscGroup
        Dim ItemCode As String
        Dim SalesCode As String
        Dim MaxDate As Date
        Dim SalesType As String
        Dim MinQty As Double
        Dim VariantCode As String
        Dim sUOM As String
    End Structure
    Private Structure ArrInvoice
        Dim InvNo As String
        Dim InvDate As Date
        Dim OrdNo As String
        Dim CustID As String
        Dim AgentID As String
        Dim Discount As Double
        Dim Subtotal As Double
        Dim GST As Double
        Dim GSTAmt As Double
        Dim TotalAmt As Double
        Dim PaidAmt As Double
        Dim payterms As String
        Dim TermDays As String
        Dim CurExRate As String
        Dim CurCode As String
    End Structure

    Private Structure ArrCrNote
        Dim CrNo As String
        Dim CrDate As Date
        Dim OrdNo As String
        Dim CustID As String
        Dim AgentID As String
        Dim Discount As Double
        Dim Subtotal As Double
        Dim GST As Double
        Dim GSTAmt As Double
        Dim TotalAmt As Double
        Dim PaidAmt As Double
        Dim payterms As String
        Dim TermDays As String
        Dim CurExRate As String
        Dim CurCode As String
    End Structure
    Private Structure ArrPrCrNote
        Dim DocEntry As String
        Dim CreditNoteNo As String
        Dim ItemNo As String
        Dim UOM As String
        Dim Qty As String
        Dim Price As String
        Dim Amt As String
        Dim VariantCode As String
        Dim Description As String
        Dim FromLocation As String
        Dim FromBin As String
        Dim DeliQty As String
        Dim Remarks As String

    End Structure

    Private Structure Arrord
        Dim Ordno As String
        Dim DocEntry As String
    End Structure
    Private Structure ArrTrans
        Dim transno As String
        Dim DocEntry As String
    End Structure
    Private Sub ImportStatus_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub
    Private Function GetCompanyName() As String
        Dim ds As New DataSet
        Dim dataDirectory As String
        dataDirectory = Windows.Forms.Application.StartupPath
        ds.ReadXml(dataDirectory & "\Simplr.xml")
        Dim table As DataTable
        For Each table In ds.Tables
            Dim row As DataRow
            If table.TableName = "CompanyName" Then
                For Each row In table.Rows
                    'MsgBox(row("CompanyName").ToString())
                    Return row("Value").ToString()
                Next row
            End If
        Next table
        Return ""
    End Function

    Private Sub ImportStatus_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ConnectDB()
        Dim ArrList As New ArrayList
        '  cnt = 0
        ArrList.Add("Customer")
        ArrList.Add("Category")
        ArrList.Add("Product")
        ArrList.Add("Item Price")
        ArrList.Add("Location")
        ArrList.Add("Agent")
        ArrList.Add("Payment Terms")
        ArrList.Add("Invoice")
        ArrList.Add("CreditNote")
        ArrList.Add("UOM")
        ArrList.Add("Inventory")
        ArrList.Add("Vendor")
        ArrList.Add("PurchaseOrder")
        ArrList.Add("Purchasecreditnote")
        ArrList.Add("Salesorder")
        ArrList.Add("TransferOrder")
        dgvStatus.Rows.Clear()
        Dim dtr As SqlDataReader
        For i = 0 To ArrList.Count - 1
            dtr = ReadRecord("Select Status from ImportStatus where TableName = " & SafeSQL(ArrList.Item(i)))
            If dtr.Read = True Then
                dgvStatus.Rows.Add(ArrList.Item(i).ToString, CBool(dtr("Status")))
            Else
                dgvStatus.Rows.Add(ArrList.Item(i).ToString, True)
            End If
            dtr.Close()
        Next
        loadCombo()
        ' LoadExportCombo()
        chkOrders.Checked = True
        'btnIm_Click(Me, Nothing)
        'Me.Close()
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

    Private Sub btnIm_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnIm.Click
        Try

            btnIm.Enabled = False
            btnEx.Enabled = False
            btnDelete.Enabled = False

            'Dim frmStatus As New status
            'frmStatus.Show()
            'sNavCompanyName = GetCompanyName()

            System.IO.File.AppendAllText(Application.StartupPath & "\ImportLog.txt", "Start Import: " & Date.Now.ToString)
            ExecuteSQL("Update System set Status='Not Completed'")
            ConnectNavDB()
            ConnectAnotherDB()
            ImportCustomer()
            ImportProduct()
            ImportErpStockRequest()
            'ImportItemPrice()
            'UpdateGPSCoordinates()
            'ImportSalesAgent()
            DisConnect()
            MsgBox("Import completed", vbInformation, "Information")
        Catch ex As Exception
            btnIm.Enabled = True
            btnEx.Enabled = True
            btnDelete.Enabled = True
            ' MsgBox(ex.Message)
            'System.IO.File.AppendAllText(Application.StartupPath & "\ErrorLog.txt", Date.Now & vbCrLf & ex.Message & vbCrLf)
            DisConnect()
        End Try

    End Sub

    Public Sub DisConnect()
        DisconnectNavDB()
        DisconnectAnotherDB()
        btnIm.Enabled = True
        btnEx.Enabled = True
        btnDelete.Enabled = True
        ExecuteSQL("Update System set LastImExDate = " & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")))
    End Sub



    Public Sub ImportCustProd()
        Dim dtr As SqlDataReader
        ExecuteSQL("Delete from CustProd")
        dtr = ReadRecord("SELECT Distinct Invoice.CustId, InvItem.ItemNo, Item.Description, Item.ItemName, Item.BaseUOM as UOM, Item.UnitPrice, Customer.PriceGroup FROM Invoice, InvItem, Item, Customer where Invoice.InvNo = InvItem.InvNo and InvItem.ItemNo = Item.ItemNo and Invoice.CustID=Customer.CustNo and (DateDiff(d, InvDt, " & SafeSQL(Format(DateAdd(DateInterval.Month, (-1) * 3, Date.Now.Date), "yyyyMMdd HH:mm:ss")) & ") <= 0)")
        'dtr = ReadRecord("Select Distinct Invoice.CustID, Customer.PriceGroup, InvItem.ItemNo, ShortDesc, Item.BaseUOM as UOM, Item.ItemName from InvItem, Invoice, Item, Customer where Customer.CustNo = Invoice.CustId and Invoice.InvNo = InvItem.InvNo and InvItem.ItemNo = Item.ItemNo")

        While dtr.Read
            Dim dPr As Double = 0
            Dim dItemPr As Double
            dItemPr = GetPrice(dtr("ItemNo").ToString, dtr("CustId").ToString, dtr("PriceGroup").ToString, dtr("UOM").ToString)
            'If dtr("UOM").ToString = "CTN" Then
            '    If IsCustProdExists(dtr("CustId").ToString, dtr("ItemNo").ToString, dtr("UOM").ToString) = False Then
            '        ExecuteSQLAnother("Insert into CustProd (CustId, ItemNo, Description, ItemName, Uom, Price) Values (" & SafeSQL(dtr("CustID")) & " , " & SafeSQL(dtr("ItemNo")) & ", " & SafeSQL(dtr("ShortDesc")) & "," & SafeSQL(dtr("ItemName").ToString) & ", " & SafeSQL(dtr("Uom")) & ", " & dItemPr & " )")
            '    End If
            'Else
            ExecuteSQLAnother("Insert into CustProd (CustId, ItemNo, Description, ItemName, Uom, Price) Values (" & SafeSQL(dtr("CustID")) & " , " & SafeSQL(dtr("ItemNo")) & ", " & SafeSQL(dtr("Description")) & "," & SafeSQL(dtr("ItemName").ToString) & ", " & SafeSQL(dtr("Uom")) & ", " & dItemPr & " )")
            'End If
        End While
        dtr.Close()
        dtr.Dispose()
        dtr = Nothing
    End Sub

    Public Function GetPrice(ByVal sItemId As String, ByVal sCustNo As String, ByVal sPrGroup As String, ByVal sUOM As String) As Double
        Dim dtr As SqlDataReader
        Dim dPr1 As Double = 0
        Dim dPr2 As Double = 0
        Dim dPr3 As Double = 0
        Dim dPr As Double = 0
        dtr = ReadRecordAnother("Select UnitPrice from ItemPr where UOM= " & SafeSQL(sUOM) & " and ItemNo = '" & sItemId & "' and PriceGroup = '" & sPrGroup & "' and SalesType = 'Customer Price Group' and MinQty=0")
        If dtr.Read Then
            dPr1 = dtr("UnitPrice")
            dPr = dPr1
        End If
        dtr.Close()
        dtr = ReadRecordAnother("Select UnitPrice from ItemPr where UOM= " & SafeSQL(sUOM) & " and  ItemNo = '" & sItemId & "' and PriceGroup = '" & sCustNo & "'and SalesType = 'Customer' and MinQty=0")
        If dtr.Read Then
            dPr2 = dtr("UnitPrice")
            If dPr2 < dPr And dPr2 > 0 Then
                dPr = dPr2
            End If
        End If
        dtr.Close()
        dtr = ReadRecordAnother("Select UnitPrice from ItemPr where UOM= " & SafeSQL(sUOM) & " and ItemNo = '" & sItemId & "' and SalesType = 'All Customers' and MinQty=0")
        If dtr.Read Then
            dPr3 = dtr("UnitPrice")
            If dPr3 < dPr And dPr3 > 0 Then
                dPr = dPr3
            End If
        End If
        dtr.Close()
        If dPr = 0 Then
            dtr = ReadRecordAnother("Select UnitPrice from Item where ItemNo = '" & sItemId & "'")
            If dtr.Read Then
                dPr = dtr("UnitPrice")
            End If
            dtr.Close()
        End If
        Return dPr
    End Function

    Private Function IsExists(ByVal strSql As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = ReadRecord(strSql)
        bAns = dtr.Read
        dtr.Close()
        Return bAns
    End Function

    Public Sub ImportShipToAddress()
        Dim dt As DateTime
        Dim dtr As SqlDataReader
        dt = Date.Now
        ExecuteSQL("Delete CustomerBill")
        ExecuteSQL("Update Customer Set BillMultiple = 0 ")
        Dim icnt As Integer = 1
        dtr = ReadNavRecord("Select * from """ & sNavCompanyName & "Ship-to Address""")
        While dtr.Read
            If IsExists("Select * from CustomerBill where AcBillRef = " & SafeSQL(dtr("Code").ToString) & " and CustNo = " & SafeSQL(dtr("Customer No_").ToString)) = True Then
                'Update
                ExecuteSQL("Update CustomerBill Set CustName = " & SafeSQL(dtr("Name").ToString) & ",Address = " & SafeSQL(dtr("Address").ToString) & " , Address2 = " & SafeSQL(dtr("Address 2").ToString) _
                        & " , City = " & SafeSQL(dtr("City").ToString) & ",ContactPerson = " & SafeSQL(dtr("Contact").ToString) & " ,Phone = " & SafeSQL(dtr("Phone No_").ToString) _
                        & " , PostCode = " & SafeSQL(dtr("Post Code").ToString) _
                        & " where AcBillRef = " & SafeSQL(dtr("Code").ToString) & " and CustNo = " & SafeSQL(dtr("Customer No_").ToString))
            Else
                'Insert
                ExecuteSQL("Insert into CustomerBill(AcBillRef, CompanyName, CustNo, CustName, Address, Address2, City, ContactPerson, Phone, PostCode) VALUES " _
                        & "(" & SafeSQL(dtr("Code").ToString) & "," & SafeSQL("STD") & "," & SafeSQL(dtr("Customer No_").ToString) & "," & SafeSQL(dtr("Name").ToString) _
                        & "," & SafeSQL(dtr("Address").ToString) & "," & SafeSQL(dtr("Address 2").ToString) & "," & SafeSQL(dtr("City").ToString) & "," & SafeSQL(dtr("Contact").ToString) _
                        & "," & SafeSQL(dtr("Phone No_").ToString) & "," & SafeSQL(dtr("Post Code").ToString) & ")")

            End If
            ExecuteSQL("Update Customer Set BillMultiple = 1 where CustNo = " & SafeSQL(dtr("Customer No_").ToString))
        End While
        dtr.Close()
    End Sub

    Private Sub UpdateLastTimeStamp(ByVal TableName As String, ByVal iValue As Date)
        ExecuteSQL("Update SyncTimeStamp Set LastTimeStamp=1111, LastImportDate = " & SafeSQL(Format(Date.Now, "yyyyMMdd")) & " where TableName = " & SafeSQL(TableName))
    End Sub

    Private Sub InsertLastTimeStamp(ByVal TableName As String, ByVal iValue As Date)
        ExecuteSQL("Insert into SyncTimeStamp (LastTimeStamp, TableName, LastImportDate) values (1111," & SafeSQL(TableName) & "," & SafeSQL(Format(Date.Now, "yyyyMMdd")) & ")")
    End Sub


    Private Function GetLastTimeStamp(ByVal TableName As String, ByRef iValue As Date, ByRef dNewRecord As Int16) As Boolean
        Dim bFound As Boolean = False
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select LastImportDate, LastTimeStamp from SyncTimeStamp where TableName = " & SafeSQL(TableName))
        If dtr.Read Then
            iValue = dtr("LastImportDate")
            dNewRecord = dtr("LastTimeStamp")
            bFound = True
        End If
        dtr.Close()
        Return bFound
    End Function

    Public Sub ImportCustomer()

        Dim icnt As Integer = 1
        Dim iValue As Date = Date.Now
        Dim dNewRecord As Int16 = 0

        Dim arrInvNo = New ArrayList
        Dim sQry As String = ""
        Dim dTransport As Double = 0
        Dim dExRate As Double = 0
        Dim dtr As SqlDataReader
        Dim dtr1 As SqlDataReader
        Dim sExtDocNo As String = ""
        Dim sSql As String
        Try

            ExecuteSQL("Update Customer set Active = 0")
            dtr = ReadNavRecord("Select * from Customer")


            While dtr.Read
                If dtr("CustNo").ToString <> "" Then
                    If IsCustomerExists(dtr("CustNo").ToString) Then
                        sQry = "UPDATE [dbo].[Customer]" & _
                               " SET [CustName] = " & SafeSQL(dtr("CustName").ToString) & _
                                 " ,[ChineseName] = " & SafeSQL(dtr("CustName").ToString) & _
                                  ",[SearchName] = " & SafeSQL(dtr("CustName").ToString) & _
                                  ",[Address] = " & SafeSQL(dtr("Address").ToString) & _
                                  ",[Address2] = " & SafeSQL(dtr("House_Number").ToString) & _
                                  ",[Address3] = " & SafeSQL(dtr("Alley").ToString) & _
                                  ",[Address4] = " & SafeSQL(dtr("Amphoe").ToString) & _
                                  ",[Tambon] = " & SafeSQL(dtr("Tambon").ToString) & _
                                  ",[Province] = " & SafeSQL(dtr("Province").ToString) & _
                                  ",[PostCode] = " & SafeSQL(dtr("PostCode").ToString) & _
                                  ",[CountryCode] = " & SafeSQL(dtr("CountryCode").ToString) & _
                                  ",[Phone] = " & SafeSQL(dtr("Phone").ToString) & _
                                  ",[ContactPerson] = " & SafeSQL(dtr("ContactPerson").ToString) & _
                                  ",[Balance] = " & SafeSQL(dtr("Balance").ToString) & _
                                  ",[CreditLimit] = " & SafeSQL(dtr("CreditLimit").ToString) & _
                                  ",[ZoneCode] = " & SafeSQL(dtr("ZoneCode").ToString) & _
                                  ",[FaxNo] = " & SafeSQL(dtr("FaxNo").ToString) & _
                                  ",[PaymentTerms] = " & SafeSQL(dtr("PaymentTerms").ToString) & _
                                  ",[ShipAgent] = " & SafeSQL(dtr("ShipAgent").ToString) & _
                                  ",[Bill-toNo] = " & SafeSQL(dtr("Bill-toNo").ToString) & _
                                  ",[Active] = " & SafeSQL(dtr("Active").ToString) & _
                                  ",[ShipName] = " & SafeSQL(dtr("ShipName").ToString) & _
                                  ",[ShipAddr] = " & SafeSQL(dtr("DeliveryAddress").ToString) & _
                                  ",[ShipAddr2] = " & SafeSQL(dtr("Ship_House_Number").ToString) & _
                                  ",[ShipAddr3] = " & SafeSQL(dtr("Ship_Alley").ToString) & _
                                  ",[ShipAddr4] = " & SafeSQL(dtr("Ship_Amphoe").ToString) & _
                                  ",[ShipCountryCode] = " & SafeSQL(dtr("Ship_Tambon").ToString) & _
                                  ",[ShipCity] = " & SafeSQL(dtr("Ship_Province").ToString) & _
                                  ",[ShipPost] = " & SafeSQL(dtr("ShipPostCode").ToString) & _
                                  ",[GSTType] = " & SafeSQL(dtr("GSTType").ToString) & _
                                  ",[Remarks] = " & SafeSQL(dtr("Remarks").ToString) & _
                                  ",[Dimension1] = " & SafeSQL(dtr("Reference").ToString) & _
                                  ",[DiscountGroup] = " & SafeSQL(dtr("DiscountGroup").ToString) & _
                                  ",[GSTNO] = " & SafeSQL(dtr("GSTNO").ToString) & _
                                  ",[Channel] = " & SafeSQL(dtr("Branch").ToString) & _
                                  ",[CustType] = " & SafeSQL(dtr("Buss_Type").ToString) & _
                                  ",[Dimension2] = " & SafeSQL(dtr("Buss_TypeOther").ToString) & _
                                  ",[Remarks2] = " & SafeSQL(dtr("Credit_Time").ToString) & _
                                  ",[CustGrade] = " & SafeSQL(dtr("Institution_Type").ToString) & _
                                  ",[CompanyName] = " & SafeSQL(dtr("Institution_Other").ToString) & _
                                  ",[Latitude] = " & SafeSQL(dtr("Latitude").ToString) & _
                                  ",[Longitude] = " & SafeSQL(dtr("Longitude").ToString) & _
                                  ",[Location] = " & SafeSQL(dtr("Location").ToString) & _
                                  ",[Location_Other] = " & SafeSQL(dtr("Location_Other").ToString) & _
                                  ",[GSTProdGroup] = " & SafeSQL(dtr("Product_List").ToString) & _
                                  ",[Shop_Type] = " & SafeSQL(dtr("Shop_Type").ToString) & _
                                  ",[Tax_Id] = " & SafeSQL(dtr("Tax_Id").ToString) & _
                                  ",[SalesAgent] = " & SafeSQL(dtr("Area").ToString) & _
                                  ",[Area] = " & SafeSQL(dtr("Area").ToString) & _
                                  ",[PaymentMethod] = 'Cash' " & _
                                  " WHERE CustNo = " & SafeSQL(dtr("CustNo").ToString)

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Import Customer - " & dtr("CustNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteSQL(sQry)

                    Else
                        sQry = "Insert into Customer ([CustNo],	[CustName],	[ChineseName],	[SearchName],	[Address],	Address2,	Address3,	Address4,	[Tambon],	[Province],	[PostCode],                    [CountryCode],	[Phone],	[ContactPerson],	[Balance],	[CreditLimit],	[ZoneCode],	[FaxNo],	[PaymentTerms],	[ShipAgent],	[Bill-toNo],[Active],	[ShipName],	ShipAddr,	ShipAddr2,	ShipAddr3,	ShipAddr4,	ShipCountryCode,	ShipCity,	[ShipPost],	[GSTType],[Remarks],	Dimension1,	[DiscountGroup],	[GSTNO],	Channel,	CustType,	Dimension2,	Remarks2,	CustGrade,	CompanyName,[Latitude],	[Longitude],	[Location],	[Location_Other],	GSTProdGroup,	[Shop_Type],	[Tax_Id],	[SalesAgent], [Area], PaymentMethod)" & _
                                "Values (" & SafeSQL(dtr("CustNo").ToString) & "," & SafeSQL(dtr("CustName").ToString) & "," & SafeSQL(dtr("CustName").ToString) & "," & SafeSQL(dtr("CustName").ToString) & "," & SafeSQL(dtr("Address").ToString) & "," & SafeSQL(dtr("House_Number").ToString) & "," & SafeSQL(dtr("Alley").ToString) & "," & SafeSQL(dtr("Amphoe").ToString) & "," & SafeSQL(dtr("Tambon").ToString) & "," & SafeSQL(dtr("Province").ToString) & "," & SafeSQL(dtr("PostCode").ToString) & "," & SafeSQL(dtr("CountryCode").ToString) & "," & SafeSQL(dtr("Phone").ToString) & "," & SafeSQL(dtr("ContactPerson").ToString) & "," & SafeSQL(dtr("Balance").ToString) & "," & SafeSQL(dtr("CreditLimit").ToString) & "," & SafeSQL(dtr("ZoneCode").ToString) & "," & SafeSQL(dtr("FaxNo").ToString) & "," & SafeSQL(dtr("PaymentTerms").ToString) & "," & SafeSQL(dtr("ShipAgent").ToString) & "," & SafeSQL(dtr("Bill-toNo").ToString) & "," & SafeSQL(dtr("Active").ToString) & "," & SafeSQL(dtr("ShipName").ToString) & "," & SafeSQL(dtr("DeliveryAddress").ToString) & "," & SafeSQL(dtr("Ship_House_Number").ToString) & "," & SafeSQL(dtr("Ship_Alley").ToString) & "," & SafeSQL(dtr("Ship_Amphoe").ToString) & "," & SafeSQL(dtr("Ship_Tambon").ToString) & "," & SafeSQL(dtr("Ship_Province").ToString) & "," & SafeSQL(dtr("ShipPostCode").ToString) & "," & SafeSQL(dtr("GSTType").ToString) & "," & SafeSQL(dtr("Remarks").ToString) & "," & SafeSQL(dtr("Reference").ToString) & "," & SafeSQL(dtr("DiscountGroup").ToString) & "," & SafeSQL(dtr("GSTNO").ToString) & "," & SafeSQL(dtr("Branch").ToString) & "," & SafeSQL(dtr("Buss_Type").ToString) & "," & SafeSQL(dtr("Buss_TypeOther").ToString) & "," & SafeSQL(dtr("Credit_Time").ToString) & "," & SafeSQL(dtr("Institution_Type").ToString) & "," & SafeSQL(dtr("Institution_Other").ToString) & "," & SafeSQL(dtr("Latitude").ToString) & "," & SafeSQL(dtr("Longitude").ToString) & "," & SafeSQL(dtr("Location").ToString) & "," & SafeSQL(dtr("Location_Other").ToString) & "," & SafeSQL(dtr("Product_List").ToString) & "," & SafeSQL(dtr("Shop_Type").ToString) & "," & SafeSQL(dtr("Tax_Id").ToString) & "," & SafeSQL(dtr("Area").ToString) & "," & SafeSQL(dtr("Area").ToString) & ",'Cash')"


                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Import Customer - " & dtr("CustNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteSQL(sQry)
                    End If

                End If

            End While
            dtr.Close()
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export Invoices - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try
    End Sub

    Private Sub ImportErpStockRequest()
        Dim icnt As Integer = 1
        Dim iValue As Date = Date.Now
        Dim dNewRecord As Int16 = 0

        Dim arrInvNo = New ArrayList
        Dim sQry As String = ""
        Dim dTransport As Double = 0
        Dim dExRate As Double = 0
        Dim dtr As SqlDataReader
        Dim dtr1 As SqlDataReader
        Dim sExtDocNo As String = ""
        Dim sSql As String
        Try

            'ExecuteSQL("Update Customer set Active = 0")
            dtr = ReadNavRecord("Select * from ErpStockRequest where IsNull(IsRead,0) = 0 Order by StockInNo")


            While dtr.Read
                If dtr("CustNo").ToString <> "" Then
                    If IsStockInNoExists(dtr("StockInNo").ToString) Then

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Import ErpStockReq - " & dtr("CustNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                    Else
                        sQry = "Insert into DeliveryOrderHdr (OrdNo, OrdDt,AgentID)" & _
                                "Values (" & SafeSQL(dtr("StockInNo").ToString) & "," & SafeSQL(Format(dtr("TransDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("AgentId").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Import ErpStockReq - " & dtr("StockInNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteSQL(sQry)

                        sQry = "Insert into DeliveryOrdItem (OrdNo, ItemNo,UOM,Qty,Location,Remarks,ReasonCode)" & _
                               "Values (" & SafeSQL(dtr("StockInNo").ToString) & "," & SafeSQL(dtr("ItemNo").ToString) & "," & SafeSQL(dtr("UOM").ToString) & "," & SafeSQL(dtr("Qty").ToString) & "," & SafeSQL(dtr("Location").ToString) & "," & SafeSQL(dtr("Remarks").ToString) & "," & SafeSQL(dtr("Reason").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Import ErpStockReq - " & dtr("StockInNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteSQL(sQry)

                    End If
                End If
            End While
            dtr.Close()
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export Invoices - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try
    End Sub


    Public Sub ImportVendor()
        Dim dtr As SqlDataReader
        Dim icnt As Integer = 1
        Dim iValue As Date = Date.Now
        Dim ivalueLast30 As Date = iValue.AddDays(-600)
        Dim dNewRecord As Int16 = 0
        Dim bSync As Boolean = GetLastTimeStamp("Vendor", iValue, dNewRecord)
        If dNewRecord = 0 Then
            dtr = ReadNavRecord("SELECT * FROM OCRD WHERE (CardType = 'S') and CardCode <> 'WILLY MOTOR CO'")
        Else
            ' dtr = ReadNavRecord("SELECT * FROM OCRD WHERE (CardType = 'S') and CardCode <> 'WILLY MOTOR CO' and (UpdateDate >= " & SafeSQL(Format(ivalueLast30, "yyyyMMdd")) & " or UpdateDate is null) order by UpdateDate")
            dtr = ReadNavRecord("SELECT * FROM OCRD WHERE (CardType = 'S') and CardCode <> 'WILLY MOTOR CO'")
        End If
        Dim sdVenNo As String = "", Str As String = ""
        While dtr.Read
            sdVenNo = dtr("CardCode").ToString
            If dtr("CardCode").ToString <> "" And IsVendorExists(sdVenNo) = True Then
                Str = "Update Vendor Set VendorName  = " & SafeSQL(dtr("CardName").ToString.Trim) & ", Address = " & SafeSQL(dtr("Address").ToString.Trim) & ", Address2 = " & SafeSQL("") & ", Address3 = '', Address4 = '', PostCode = " & SafeSQL(dtr("ZipCode").ToString.Trim) & ", City = '', CountryCode = " & SafeSQL(dtr("Country").ToString.Trim) & ", PhoneNo = " & SafeSQL(dtr("Phone1").ToString.Trim) & ", Contact =" & SafeSQL(dtr("CntctPrsn").ToString.Trim) & ", FaxNo = " & SafeSQL(dtr("Fax").ToString.Trim) & ", email = " & SafeSQL(dtr("E_Mail").ToString.Trim) & ",CurrencyCode = " & SafeSQL("").ToString & ",Active = 1,  DTG = " & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")) & ",LocationCode = '',PayTerms = '' Where VendorNo = " & SafeSQL(dtr("CardCode").ToString.Trim)
                ' MsgBox(Str)
                ExecuteSQLAnother(Str)
            Else
                Str = "Insert into Vendor(VendorNo, VendorName, ChineseName, Address, Address2, Address3, Address4, PostCode,City,CountryCode,PhoneNo,Contact,FaxNo,Email,Website,Active,DTG,LocationCode,PayTerms) Values " & "(" & SafeSQL(dtr("CardCode").ToString) & "," & SafeSQL(dtr("CardName").ToString) & "," & SafeSQL(dtr("CardName").ToString) & "," & SafeSQL(dtr("Address").ToString) & "," & SafeSQL("") & "," & SafeSQL("") & "," & SafeSQL("") & "," & SafeSQL(dtr("ZipCode").ToString.Trim) & ",''," & SafeSQL(dtr("Country").ToString.Trim) & "," & SafeSQL(dtr("Phone1").ToString.Trim) & "," & SafeSQL(dtr("Phone1").ToString.Trim) & "," & SafeSQL(dtr("Fax").ToString.Trim) & ",'', '',1," & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")) & " ,'','')"
                'MsgBox(Str)
                ExecuteSQLAnother(Str)
            End If
            icnt += 1
        End While
        dtr.Close()
        dtr.Dispose()
        dtr = Nothing
        If bSync = True Then
            UpdateLastTimeStamp("Vendor", iValue)
        Else
            InsertLastTimeStamp("Vendor", iValue)
        End If
        'dtr = ReadNavRecord("Select * from """ & sNavCompanyName & "Default Dimension"" where ""Table Name""='Vendor'")
        'While dtr.Read
        '    If dtr("Dimension Code").ToString = "REGION" Then
        '        ExecuteSQLAnother("Update Customer Set ZoneCode=" & SafeSQL(dtr("Dimension Value Code").ToString) & ", Dimension2=" & SafeSQL(dtr("Dimension Value Code").ToString) & " where CustNo=" & SafeSQL(dtr("No_")))
        '    End If
        'End While
        'dtr.Close()

        'ExecuteSQL("Update Customer Set PriceGroup ='ALL' Where (PriceGroup='' or PriceGroup is Null)")
        'ExecuteSQL("Update Customer Set AcBillRef ='' Where AcBillRef is Null")
        ' ExecuteSQL("Update Customer Set ShipName = CustName where (ShipName='' or ShipName is Null)")
        ' ExecuteSQL("Update Customer Set Active=0 where CustName like '%Closed%'")
        ExecuteSQL("Update ImportStatus Set Status = 0 where TableName='Vendor'")
    End Sub

    Public Sub ImportCustGroup()
        Dim dtr As SqlDataReader
        ExecuteSQL("Delete From CustGroup")
        dtr = ReadNavRecord("Select * from OITB")
        While dtr.Read
            If dtr("ItmsGrpCod").ToString <> "" Then
                If IsExists("Select Code from CustGroup where code=" & SafeSQL(dtr("ItmsGrpCod"))) = False Then
                    ExecuteSQLAnother("Insert into CustGroup(Code , Description) Values (" & SafeSQL(dtr("ItmsGrpCod").ToString) & "," & SafeSQL(IIf(dtr("ItmsGrpNam").ToString = "", dtr("ItmsGrpCod").ToString, dtr("ItmsGrpNam").ToString)) & ")")
                Else
                    ExecuteSQLAnother("Update CustGroup set Description=" & SafeSQL(IIf(dtr("ItmsGrpNam").ToString = "", dtr("GroupCode").ToString, dtr("ItmsGrpNam").ToString)) & " Where Code=" & SafeSQL(dtr("ItmsGrpCod").ToString))
                End If
            End If
        End While
        dtr.Close()
        dtr.Dispose()
        dtr = Nothing
    End Sub

    Public Sub ImportCategory()
        Dim dtr As SqlDataReader
        'ExecuteSQL("Delete From Category")
        dtr = ReadNavRecord("Select * from OITB")
        While dtr.Read
            If dtr("ItmsGrpCod").ToString <> "" Then
                If IsExists("Select Code from Category where code=" & SafeSQL(dtr("ItmsGrpCod"))) = False Then
                    ExecuteSQLAnother("Insert into Category(Code , Description) Values (" & SafeSQL(dtr("ItmsGrpCod").ToString) & "," & SafeSQL(IIf(dtr("ItmsGrpNam").ToString = "", dtr("ItmsGrpCod").ToString, dtr("ItmsGrpNam").ToString)) & ")")
                Else
                    ExecuteSQLAnother("Update Category set Description=" & SafeSQL(IIf(dtr("ItmsGrpNam").ToString = "", dtr("ItmsGrpCod").ToString, dtr("ItmsGrpNam").ToString)) & " Where Code=" & SafeSQL(dtr("ItmsGrpCod").ToString))
                End If
            End If
        End While
        dtr.Close()
        dtr.Dispose()
        dtr = Nothing

    End Sub

    Public Sub ImportBrand()
        Dim dtr As SqlDataReader
        ExecuteSQL("Delete From Brand")
        dtr = ReadNavRecord("Select * from OMRC")
        While dtr.Read
            If dtr("FirmCode").ToString <> "" Then
                If IsExists("Select Code from Brand where code=" & SafeSQL(dtr("FirmCode"))) = False Then
                    ExecuteSQLAnother("Insert into Brand(Code , Description) Values (" & SafeSQL(dtr("FirmCode").ToString) & "," & SafeSQL(IIf(dtr("FirmName").ToString = "", dtr("FirmCode").ToString, dtr("FirmName").ToString)) & ")")
                Else
                    ExecuteSQLAnother("Update Brand set Description=" & SafeSQL(IIf(dtr("FirmName").ToString = "", dtr("FirmCode").ToString, dtr("FirmName").ToString)) & " Where Code=" & SafeSQL(dtr("FirmCode").ToString))
                End If
            End If
        End While
        dtr.Close()
        dtr.Dispose()
        dtr = Nothing
    End Sub

    Public Sub ImportProduct()
        Dim dtr As SqlDataReader

        ExecuteSQL("Update Item Set Active = 0")


        Dim icnt As Integer = 1
        Dim iValue As Date = Date.Now
        Dim dNewRecord As Int16 = 0

        Dim arrInvNo = New ArrayList
        Dim sQry As String = ""
        Dim dTransport As Double = 0
        Dim dExRate As Double = 0
        Dim dtr1 As SqlDataReader
        Dim sExtDocNo As String = ""
        Dim sSql As String
        Try


            dtr = ReadNavRecord("Select * from Item")

            While dtr.Read
                If dtr("ItemNo").ToString <> "" Then
                    If IsItemExists(dtr("ItemNo")) Then
                        sQry = "Update Item Set ItemNo = " & SafeSQL(dtr("ItemNo").ToString) & _
                                ", Description = " & SafeSQL(dtr("Item_Name").ToString) & _
                                ", ItemName = " & SafeSQL(dtr("Item_Name").ToString) & _
                                ", ShortDesc= " & SafeSQL(dtr("Item_Name").ToString) & _
                                ", ChineseDesc = " & SafeSQL(dtr("Item_Name").ToString) & _
                                ", BaseUOM = " & SafeSQL(dtr("LooseUOm").ToString) & _
                                ", UnitPrice = " & SafeSQL(dtr("Price1").ToString) & _
                                ", Active = " & SafeSQL(dtr("Active").ToString) & _
                                ", Category = " & SafeSQL(dtr("Category").ToString) & _
                                ", Brand = " & SafeSQL(dtr("Brand").ToString) & _
                                ", SubCategory = " & SafeSQL(dtr("Sub_Category").ToString) & _
                                ", ToPDA = 1" & _
                                ", SubBrand = " & SafeSQL(dtr("Sub_Brand").ToString) & _
                                ", Favourite = " & SafeSQL(dtr("Favourite").ToString) & _
                                ", BulkUOM = " & SafeSQL(dtr("BulkUOM").ToString) & _
                                ", BulkQty = " & SafeSQL(dtr("BulkQty").ToString) & _
                                ", LooseUOM = " & SafeSQL(dtr("LooseUOM").ToString) & _
                                ", LooseQty = " & SafeSQL(dtr("LooseQty").ToString) & _
                                ", MaxQty1 = " & If(IsDBNull(dtr("MaxQty1")), 0, dtr("MaxQty1")) & _
                                ", MaxQty2 = " & If(IsDBNull(dtr("MaxQty2")), 0, dtr("MaxQty2")) & _
                                ", MaxQty3 = " & If(IsDBNull(dtr("MaxQty3")), 0, dtr("MaxQty3")) & _
                                ", MaxQty4 = " & If(IsDBNull(dtr("MaxQty4")), 0, dtr("MaxQty4")) & _
                                ", MaxQty5 = " & If(IsDBNull(dtr("MaxQty5")), 0, dtr("MaxQty5")) & _
                                ", Price1 = " & SafeSQL(dtr("Price1").ToString) & _
                                ", Price2 = " & SafeSQL(dtr("Price2").ToString) & _
                                ", Price3 = " & SafeSQL(dtr("Price3").ToString) & _
                                ", Price4 = " & SafeSQL(dtr("Price4").ToString) & _
                                ", Price5 = " & SafeSQL(dtr("Price5").ToString) & _
                                ", PackType = " & SafeSQL(dtr("PackType").ToString) & _
                                ", PriceGroup = " & SafeSQL(dtr("PriceGroup").ToString) & _
                                ", [Return] = " & SafeSQL(dtr("Return").ToString) & _
                                ", Sale = " & SafeSQL(dtr("Sale").ToString) & _
                                ", VAT = " & SafeSQL(dtr("VAT").ToString) & _
                                ", [Trading_Discount] = " & SafeSQL(dtr("Trading_Discount").ToString) & _
                                ", [Size_Unit] = " & SafeSQL(dtr("Size_Unit").ToString) & _
                                ", Size  = " & SafeSQL(dtr("Size").ToString) & _
                                " Where ItemNo = " & SafeSQL(dtr("ItemNo").ToString)

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export Item - " & dtr("ItemNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteSQL(sQry)

                    Else
                        sQry = "Insert into Item (ItemNo, Description, ItemName, ShortDesc, ChineseDesc, BaseUOM, UnitPrice, Active, Category, Brand, SubCategory, ToPDA, SubBrand, Favourite,  BulkUOM, BulkQty, LooseUOM, LooseQty, MaxQty1, MaxQty2, MaxQty3, MaxQty4, MaxQty5, Price1, Price2, Price3, Price4, Price5, PackType, PriceGroup, [Return], Sale, VAT, [Trading_Discount], [Size_Unit], Size)" & _
                            "Values (" & SafeSQL(dtr("ItemNo").ToString) & "," & SafeSQL(dtr("Item_Name").ToString) & "," & SafeSQL(dtr("Item_Name").ToString) & "," & SafeSQL(dtr("Item_Name").ToString) & "," & SafeSQL(dtr("Item_Name").ToString) & "," & SafeSQL(dtr("LooseUOM").ToString) & "," & SafeSQL(dtr("Price1").ToString) & "," & SafeSQL(dtr("Active").ToString) & "," & SafeSQL(dtr("Category").ToString) & _
                            "," & SafeSQL(dtr("Brand").ToString) & "," & SafeSQL(dtr("Sub_Category").ToString) & ",1," & SafeSQL(dtr("Sub_Brand").ToString) & "," & SafeSQL(dtr("Favourite").ToString) & "," & SafeSQL(dtr("BulkUOM").ToString) & "," & SafeSQL(dtr("BulkQty").ToString) & "," & SafeSQL(dtr("LooseUOM").ToString) & "," & SafeSQL(dtr("LooseQty").ToString) & "," & If(IsDBNull(dtr("MaxQty1")), 0, dtr("MaxQty1")) & _
                            "," & If(IsDBNull(dtr("MaxQty2")), 0, dtr("MaxQty2")) & "," & If(IsDBNull(dtr("MaxQty3")), 0, dtr("MaxQty3")) & "," & If(IsDBNull(dtr("MaxQty4")), 0, dtr("MaxQty4")) & "," & If(IsDBNull(dtr("MaxQty5")), 0, dtr("MaxQty5")) & "," & SafeSQL(dtr("Price1").ToString) & "," & SafeSQL(dtr("Price2").ToString) & "," & SafeSQL(dtr("Price3").ToString) & "," & SafeSQL(dtr("Price4").ToString) & "," & SafeSQL(dtr("Price5").ToString) & _
                            "," & SafeSQL(dtr("PackType").ToString) & "," & SafeSQL(dtr("PriceGroup").ToString) & "," & SafeSQL(dtr("Return").ToString) & "," & SafeSQL(dtr("Sale").ToString) & "," & SafeSQL(dtr("VAT").ToString) & "," & SafeSQL(dtr("Trading_Discount").ToString) & "," & SafeSQL(dtr("Size_Unit").ToString) & "," & SafeSQL(dtr("Size").ToString) & ")"


                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export Item - " & dtr("ItemNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteSQL(sQry)
                    End If
                End If
            End While
            dtr.Close()

            ExecuteSQL("Delete from UOM")
            ExecuteSQL("Insert into UOM (ItemNo, UOM, BaseQTY, DTG, PackingSize, Dimension, M3, CompanyNo, Length, Width, Height, Weight, Cubage) Select ItemNo, BaseUOM,1 , GetDate(),'',0,0,'',0,0,0,0,0 from Item")
            ExecuteSQL("Insert into UOM (ItemNo, UOM, BaseQTY, DTG, PackingSize, Dimension, M3, CompanyNo, Length, Width, Height, Weight, Cubage) Select ItemNo, BulkUOM,BulkQty , GetDate(),'',0,0,'',0,0,0,0,0 from Item where BulkQty>1 and BaseUOM <> BulkUOM")

            ExecuteSQL("Delete from Category")
            ExecuteSQL("Insert into Category (Code, Description, DTG, DisplayNo) Select Distinct Category, Category, GETDATE(),1 from Item where isnull(Category,'') <>''")

            ExecuteSQL("Delete from Brand")
            ExecuteSQL("Insert into Brand (Code, Description,  DisplayNo) Select Distinct Brand, Brand,1 from Item where isnull(Brand,'') <>''")

            ExecuteSQL("Delete from StockTakeProductTeam")
            ExecuteSQL("Insert into StockTakeProductTeam (Team, ItemNo) Select Distinct Team, ItemNo from StockTakeProductTeam")

        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export Item - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try
    End Sub


    Public Sub ImportItemPrice()
        Dim dtr As SqlDataReader
        Dim sType As String = ""
        Dim sPriceGroup As String = ""
        Dim sItemNo As String = ""
        Dim sUOm As String = ""
        Dim dQty As Double = 0
        ExecuteSQL("Delete from ItemPr")
        dtr = ReadRecord("Select ItemNo, isnull(MaxQty1,0) as MaxQty1, isnull(MaxQty2,0) as MaxQty2, isnull(MaxQty3,0) as MaxQty3, isnull(MaxQty4,0) as MaxQty4, isnull(MaxQty5,0) as MaxQty5, isnull(Price1,0) as Price1, isnull(Price2,0) as Price2, isnull(Price3,0) as Price3, isnull(Price4,0) as Price4, isnull(Price5,0) as Price5, BulkUOM   from Item order by ItemNo")
        While dtr.Read
            sType = "Customer Price Group"
            If dtr("Price1") <> 0 Then
                ExecuteSQLAnother("Insert into ItemPr(PriceGroup, ItemNo, UnitPrice, SalesType, Minqty, VariantCode, UOM, FromDate, ToDate) Values (" & SafeSQL("STANDARD") & "," & SafeSQL(dtr("ItemNo").ToString.Trim) & "," & IIf(IsDBNull(dtr("Price1")), 0, dtr("Price1")) & "," & SafeSQL(sType) & ",1," & SafeSQL("") & "," & SafeSQL(dtr("BulkUOM").ToString) & "," & SafeSQL(Format(Date.Now.AddDays(-1), "yyyyMMdd 00:00:00")) & "," & SafeSQL(Format(Date.Now.AddYears(1), "yyyyMMdd 23:59:59")) & ")")
            End If
            If dtr("Price2") <> 0 Then
                ExecuteSQLAnother("Insert into ItemPr(PriceGroup, ItemNo, UnitPrice, SalesType, Minqty, VariantCode, UOM, FromDate, ToDate) Values (" & SafeSQL("STANDARD") & "," & SafeSQL(dtr("ItemNo").ToString.Trim) & "," & IIf(IsDBNull(dtr("Price2")), 0, dtr("Price2")) & "," & SafeSQL(sType) & "," & dtr("MaxQty1") & "," & SafeSQL("") & "," & SafeSQL(dtr("BulkUOM").ToString) & "," & SafeSQL(Format(Date.Now.AddDays(-1), "yyyyMMdd 00:00:00")) & "," & SafeSQL(Format(Date.Now.AddYears(1), "yyyyMMdd 23:59:59")) & ")")
            End If
            If dtr("Price3") <> 0 Then
                ExecuteSQLAnother("Insert into ItemPr(PriceGroup, ItemNo, UnitPrice, SalesType, Minqty, VariantCode, UOM, FromDate, ToDate) Values (" & SafeSQL("STANDARD") & "," & SafeSQL(dtr("ItemNo").ToString.Trim) & "," & IIf(IsDBNull(dtr("Price3")), 0, dtr("Price3")) & "," & SafeSQL(sType) & "," & dtr("MaxQty2") & "," & SafeSQL("") & "," & SafeSQL(dtr("BulkUOM").ToString) & "," & SafeSQL(Format(Date.Now.AddDays(-1), "yyyyMMdd 00:00:00")) & "," & SafeSQL(Format(Date.Now.AddYears(1), "yyyyMMdd 23:59:59")) & ")")
            End If
            If dtr("Price4") <> 0 Then
                ExecuteSQLAnother("Insert into ItemPr(PriceGroup, ItemNo, UnitPrice, SalesType, Minqty, VariantCode, UOM, FromDate, ToDate) Values (" & SafeSQL("STANDARD") & "," & SafeSQL(dtr("ItemNo").ToString.Trim) & "," & IIf(IsDBNull(dtr("Price4")), 0, dtr("Price4")) & "," & SafeSQL(sType) & "," & dtr("MaxQty3") & "," & SafeSQL("") & "," & SafeSQL(dtr("BulkUOM").ToString) & "," & SafeSQL(Format(Date.Now.AddDays(-1), "yyyyMMdd 00:00:00")) & "," & SafeSQL(Format(Date.Now.AddYears(1), "yyyyMMdd 23:59:59")) & ")")
            End If

            If dtr("Price5") <> 0 Then
                ExecuteSQLAnother("Insert into ItemPr(PriceGroup, ItemNo, UnitPrice, SalesType, Minqty, VariantCode, UOM, FromDate, ToDate) Values (" & SafeSQL("STANDARD") & "," & SafeSQL(dtr("ItemNo").ToString.Trim) & "," & IIf(IsDBNull(dtr("Price5")), 0, dtr("Price5")) & "," & SafeSQL(sType) & "," & dtr("MaxQty4") & "," & SafeSQL("") & "," & SafeSQL(dtr("BulkUOM").ToString) & "," & SafeSQL(Format(Date.Now.AddDays(-1), "yyyyMMdd 00:00:00")) & "," & SafeSQL(Format(Date.Now.AddYears(1), "yyyyMMdd 23:59:59")) & ")")
            End If
        End While
        dtr.Close()
        dtr.Dispose()
        dtr = Nothing
        ExecuteSQL("Delete from ItemPr where UnitPrice = 0")
        ExecuteSQL("Update ItemPr Set MinPrice = 0 where MinPrice is Null")
        '  ExecuteSQL("Update ItemPr set UOM = Item.BaseUOM from Item where Item.ItemNo = ItemPr.ItemNo and ItemPr.UOM = ''")
    End Sub

    Public Sub ImportLocation()
        Dim dtr As SqlDataReader
        dtr = ReadNavRecord("Select * from OWHS")
        While dtr.Read
            If IsExists("Select Code from Location where code=" & SafeSQL(dtr("WhsCode"))) = False Then
                ExecuteSQLAnother("Insert into Location(Code, Name, DTG) Values (" & SafeSQL(dtr("WhsCode").ToString) & "," & SafeSQL(dtr("WhsName").ToString) & ", getdate())")
            Else
                ExecuteSQLAnother("Update Location set Name=" & SafeSQL(IIf(dtr("WhsName").ToString = "", dtr("WhsCode").ToString, dtr("WhsName").ToString)) & " Where Code=" & SafeSQL(dtr("WhsCode").ToString))
            End If
        End While

        dtr.Close()
        dtr.Dispose()
        dtr = Nothing

    End Sub

    Public Sub ImportCustAgent()
        Dim dtr As SqlDataReader
        Dim icnt As Integer = 1
        Dim iValue As Date = Date.Now
        Dim dNewRecord As Int16 = 0

        Dim arrInvNo = New ArrayList
        Dim sQry As String = ""
        Dim dTransport As Double = 0
        Dim dExRate As Double = 0
        Dim dtr1 As SqlDataReader
        Dim sExtDocNo As String = ""
        Dim sSql As String
        'ExecuteSQL("Update SalesAgent Set Active = 0")
        dtr = ReadRecord("Select Distinct AgentID as SalesAgent from CustAgent  order by AgentID")

        While dtr.Read
            If dtr("SalesAgent").ToString <> "" Then
                If IsMDTExists(dtr("SalesAgent").ToString) = False Then
                    ExecuteSQLAnother("Insert into MDT(MDTNo, Description, AgentId, Location, RouteNo, VehicleID, SolutionName) Values (" & SafeSQL(dtr("SalesAgent").ToString) & "," & SafeSQL(dtr("SalesAgent").ToString) & "," & SafeSQL("") & "," & SafeSQL(dtr("SalesAgent").ToString) & "," & SafeSQL("") & "," & SafeSQL("") & ", 'SALES')")
                Else
                    '     ExecuteSQLAnother("Update SalesAgent Set Name = " & SafeSQL(dtr("SalesAgent").ToString) & ", Active = 1 Where Code = " & SafeSQL(dtr("SalesAgent").ToString))
                End If

                If IsNoSeriesExists(dtr("SalesAgent").ToString) = False Then
                    ExecuteSQLAnother("Insert into NoSeries (MDTNo, DocType, ConditionMaster, ConditionType, ConditionValue, Prefix, LastNumber, NoLength, StartDate, EndDate) Select " & SafeSQL(dtr("SalesAgent").ToString) & ", DocType, ConditionMaster, ConditionType, ConditionValue, CASE WHEN DocType='ITEMTRANS' THEN " & SafeSQL(dtr("SalesAgent").ToString) & " +SubString(DocType,1,2) ELSE " & SafeSQL(dtr("SalesAgent").ToString) & " +SubString(DocType,1,1) END, 0 as LastNumber, NoLength, StartDate, EndDate from NoSeries where MDTNo='M1'")
                Else
                    '     ExecuteSQLAnother("Update SalesAgent Set Name = " & SafeSQL(dtr("SalesAgent").ToString) & ", Active = 1 Where Code = " & SafeSQL(dtr("SalesAgent").ToString))
                End If
            End If
        End While
        dtr.Close()
        dtr.Dispose()
        dtr = Nothing


        ExecuteSQL("Delete from CustAgent")
        dtr = ReadNavRecord(" Select Distinct SalesUnit, Area  from Area")

        While dtr.Read
            If dtr("SalesUnit").ToString <> "" Then

                sQry = "Insert into CustAgent (AgentID, CustAgentID, Position)" & _
                    "Values (" & SafeSQL(dtr("SalesUnit").ToString) & "," & SafeSQL(dtr("Area").ToString) & ",1)"


                ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export CustAgent - " & dtr("SalesUnit").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                ExecuteSQL(sQry)

            End If
        End While
        dtr.Close()


        ExecuteSQL("Delete from Team")
        dtr = ReadNavRecord(" Select Distinct Team, SalesUnit  from Team")

        While dtr.Read
            If dtr("Team").ToString <> "" Then

                sQry = "Insert into Team (Team, SalesUnit)" & _
                    "Values (" & SafeSQL(dtr("Team").ToString) & "," & SafeSQL(dtr("SalesUnit").ToString) & ")"


                ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export Team - " & dtr("Team").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                ExecuteSQL(sQry)

            End If
        End While
        dtr.Close()


        ExecuteSQL("Delete from ProdTeam")
        dtr = ReadNavRecord("Select Distinct Team, ItemNo  from ProductTeam")

        While dtr.Read
            If dtr("Team").ToString <> "" Then

                sQry = "Insert into ProdTeam (Team, ItemNo)" & _
                    "Values (" & SafeSQL(dtr("Team").ToString) & "," & SafeSQL(dtr("ItemNo").ToString) & ")"


                ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export ProdTeam - " & dtr("Team").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                ExecuteSQL(sQry)

            End If
        End While
        dtr.Close()

    End Sub

    Private Sub ExportInvoices()
        Dim sCurCode As String = ""
        Dim arrInvNo = New ArrayList
        Dim sQry As String = ""
        Dim dTransport As Double = 0
        sCurCode = "" 'dtr1("LCY Code")
        Dim dExRate As Double = 0
        Dim dtr As SqlDataReader
        Dim dtr1 As SqlDataReader
        Dim sExtDocNo As String = ""
        Dim sSql As String
        Try
            dtr = ReadRecord("Select  InvNo from Invoice where isnull(Exported,0) = 0 Order by InvNo")
            While dtr.Read
                If arrInvNo.Contains(dtr("InvNo").ToString) = False Then arrInvNo.Add(dtr("InvNo").ToString)
            End While
            dtr.Close()
            dtr.Dispose()


            For i = 0 To arrInvNo.Count - 1
                Try

                    dtr = ReadRecord("Select * from Invoice where InvNo =  " & SafeSQL(arrInvNo(i)) & " Order by Invoice.InvNo")

                    While dtr.Read
                        ExecuteNavSQL("Delete from Invoice where InvNo=" & SafeSQL(dtr("InvNo").ToString))
                        ExecuteNavSQL("Delete from InvItem where InvNo=" & SafeSQL(dtr("InvNo").ToString))

                        sQry = "Insert into Invoice (InvNo,InvDt,OrdNo,CustId,AgentId,SubTotal,GstAmt,TotalAmt,PaidAmt,PayTerms,Void,GST) Values (" & SafeSQL(dtr("InvNo").ToString) & _
                            "," & SafeSQL(Format(dtr("InvDt"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("OrdNo").ToString) & "," & SafeSQL(dtr("CustId").ToString) & "," & SafeSQL(dtr("AgentId").ToString) & _
                            "," & SafeSQL(dtr("SubTotal").ToString) & "," & SafeSQL(dtr("GstAmt").ToString) & "," & SafeSQL(dtr("TotalAmt").ToString) & "," & SafeSQL(dtr("PaidAmt").ToString) & _
                            "," & SafeSQL(dtr("PayTerms").ToString) & "," & SafeSQL(dtr("Void").ToString) & "," & SafeSQL(dtr("GST").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export Invoices - " & dtr("InvNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteNavSQL(sQry)

                    End While
                    dtr.Close()

                Catch ex As Exception
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export Invoices - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try

                Try
                    dtr1 = ReadRecord("Select * From InvItem  Where InvNo =  " & SafeSQL(arrInvNo(i)) & " Order by InvNo")
                    While dtr1.Read
                        sSql = "Insert into InvItem (InvNo,ItemNo,UOM,Qty,Price,Discount,SubAmt,SalesType,ReasonCode,[LineNo]) Values (" & SafeSQL(dtr1("InvNo").ToString) & _
                             "," & SafeSQL(dtr1("ItemNo").ToString) & "," & SafeSQL(dtr1("UOM").ToString) & "," & SafeSQL(dtr1("Qty").ToString) & "," & SafeSQL(dtr1("Price").ToString) & _
                             "," & SafeSQL(dtr1("Discount").ToString) & "," & SafeSQL(dtr1("SubAmt").ToString) & "," & SafeSQL(dtr1("SalesType").ToString) & "," & SafeSQL(dtr1("ReasonCode").ToString) & _
                             "," & SafeSQL(dtr1("LineNo").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export InvItem - " & dtr1("InvNo").ToString) & SafeSQL(arrInvNo(i)) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sSql) & ")")

                        ExecuteNavSQL(sSql)
                    End While
                    dtr1.Close()
                    ExecuteSQLAnother("Update Invoice Set Exported = 1 Where InvNo = " & SafeSQL(arrInvNo(i)))
                Catch ex As Exception
                    dtr1.Close()
                    System.IO.File.AppendAllText(Application.StartupPath & "\ErrorLog.txt", "Export Inv Item Error" & "   " & Date.Now.ToString & " " & ex.Message & vbCrLf)
                    Exit Sub
                End Try

            Next
        Catch ex As Exception
            dtr.Close()
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export Invoice'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try
    End Sub
    Private Sub ExportStockInItem()
        Dim sCurCode As String = ""
        Dim arrInvNo = New ArrayList
        Dim sQry As String = ""
        Dim dTransport As Double = 0
        sCurCode = "" 'dtr1("LCY Code")
        Dim dExRate As Double = 0
        Dim dtr As SqlDataReader
        Dim sExtDocNo As String = ""

        Try
            dtr = ReadRecord("Select  StockInNo from StockInItem where approved =1 and (exported is NULL or Exported = 0) Order by StockInNo")
            While dtr.Read
                If arrInvNo.Contains(dtr("StockInNo").ToString) = False Then arrInvNo.Add(dtr("StockInNo").ToString)
            End While
            dtr.Close()
            dtr.Dispose()


            For i = 0 To arrInvNo.Count - 1
                Try

                    dtr = ReadRecord("Select H.StockInNo, H.TransDate, H.AgentID, H.Location, ItemNo, UOM, Qty-TransitQty as Qty , H.Remarks, Reason,[LineNo] " & _
                        " from StockinHdr H INner Join  Stockinitem D on H.StockINNo=D.StockINNo where H.StockInNo = " & SafeSQL(arrInvNo(i)) & " and approved =1 and (exported is NULL or Exported = 0) Order by H.StockInNo")
                    While dtr.Read
                        ExecuteNavSQL("Delete from VanStockRequest where StockInNo=" & SafeSQL(dtr("StockInNo")) & "and ItemNo = " & SafeSQL(dtr("ItemNo")))

                        sQry = "Insert into VanStockRequest (StockInNo, TransDate, AgentID, Location, ItemNo, UOM, Qty , Remarks, Reason) Values (" & SafeSQL(dtr("StockInNo").ToString) & _
                            "," & SafeSQL(Format(dtr("TransDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("AgentID").ToString) & "," & SafeSQL(dtr("Location").ToString) & "," & SafeSQL(dtr("ItemNo").ToString) & _
                            "," & SafeSQL(dtr("UOM").ToString) & "," & dtr("Qty").ToString & "," & SafeSQL(dtr("Remarks").ToString) & "," & SafeSQL(dtr("Reason").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export StockInItem - " & dtr("StockInNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteNavSQL(sQry)
                        ExecuteSQLAnother("Update StockInItem Set Exported = 1 Where StockInNo = " & SafeSQL(arrInvNo(i)) & " and [LineNo] = " & SafeSQL(dtr("LineNo").ToString))
                    End While
                    dtr.Close()

                Catch ex As Exception
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export StockInItem - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try

            Next
        Catch ex As Exception
            dtr.Close()
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export StockInItem'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try
    End Sub
    Private Sub ExportSalesOrder()
        Dim sCurCode As String = ""
        Dim arrInvNo = New ArrayList
        Dim sQry As String = ""
        Dim dTransport As Double = 0
        sCurCode = "" 'dtr1("LCY Code")
        Dim dExRate As Double = 0
        Dim dtr As SqlDataReader
        Dim dtr1 As SqlDataReader
        Dim sExtDocNo As String = ""
        Dim sSql As String
        Try
            dtr = ReadRecord("Select  OrdNo from OrderHdr where isnull(Exported,0) = 0 Order by OrdNo")
            While dtr.Read
                If arrInvNo.Contains(dtr("OrdNo").ToString) = False Then arrInvNo.Add(dtr("OrdNo").ToString)
            End While
            dtr.Close()
            dtr.Dispose()


            For i = 0 To arrInvNo.Count - 1
                Try

                    dtr = ReadRecord("Select * from OrderHdr where OrdNo =  " & SafeSQL(arrInvNo(i)) & " Order by OrdNo")

                    While dtr.Read
                        ExecuteNavSQL("Delete from OrderHdr where OrdNo=" & SafeSQL(dtr("OrdNo").ToString))
                        ExecuteNavSQL("Delete from OrdItem where OrdNo=" & SafeSQL(dtr("OrdNo").ToString))

                        sQry = "Insert into OrderHdr (OrdNo,OrdDt,CustId,AgentId,SubTotal,GstAmt,TotalAmt,PayTerms,DeliveryDate) Values (" & SafeSQL(dtr("OrdNo").ToString) & _
                            "," & SafeSQL(Format(dtr("OrdDt"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("CustId").ToString) & "," & SafeSQL(dtr("AgentId").ToString) & "," & SafeSQL(dtr("SubTotal").ToString) & _
                            "," & SafeSQL(dtr("GstAmt").ToString) & "," & SafeSQL(dtr("TotalAmt").ToString) & "," & SafeSQL(dtr("PayTerms").ToString) & "," & SafeSQL(Format(dtr("DeliveryDate"), "yyyyMMdd HH:mm:ss")) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export Invoices - " & dtr("OrdNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteNavSQL(sQry)

                    End While
                    dtr.Close()

                Catch ex As Exception
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export OrderHdr - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try

                Try
                    dtr1 = ReadRecord("Select * From OrdItem  Where OrdNo =  " & SafeSQL(arrInvNo(i)) & " Order by OrdNo")
                    While dtr1.Read
                        sSql = "Insert into OrdItem (OrdNo,ItemNo,UOM,Qty,Price,Discount,SubAmt,SalesType,[LineNo]) Values (" & SafeSQL(dtr1("OrdNo").ToString) & _
                             "," & SafeSQL(dtr1("ItemNo").ToString) & "," & SafeSQL(dtr1("UOM").ToString) & "," & SafeSQL(dtr1("Qty").ToString) & "," & SafeSQL(dtr1("Price").ToString) & _
                             "," & SafeSQL(dtr1("Discount").ToString) & "," & SafeSQL(dtr1("SubAmt").ToString) & "," & SafeSQL(dtr1("SalesType").ToString) & "," & SafeSQL(dtr1("LineNo").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export OrdItem - " & SafeSQL(arrInvNo(i))) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sSql) & ")")


                        ExecuteNavSQL(sSql)
                    End While
                    dtr1.Close()
                    ExecuteSQLAnother("Update OrderHdr Set Exported = 1 Where OrdNo = " & SafeSQL(arrInvNo(i)))
                Catch ex As Exception
                    dtr1.Close()
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export OrdItem'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                    System.IO.File.AppendAllText(Application.StartupPath & "\ErrorLog.txt", "Export Inv Item Error" & "   " & Date.Now.ToString & " " & ex.Message & vbCrLf)
                    Exit Sub
                End Try
            Next
        Catch ex As Exception
            dtr.Close()
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export OrderHdr'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try
    End Sub
    Private Sub ExportCreditMemo()
        Dim sCurCode As String = ""
        Dim arrInvNo = New ArrayList
        Dim sQry As String = ""
        Dim dTransport As Double = 0
        sCurCode = "" 'dtr1("LCY Code")
        Dim dExRate As Double = 0
        Dim dtr As SqlDataReader
        Dim dtr1 As SqlDataReader
        Dim sExtDocNo As String = ""
        Dim sSql As String
        Try
            dtr = ReadRecord("Select  CreditNoteNo from CreditNote where isnull(Exported,0) = 0 Order by CreditNoteNo")
            While dtr.Read
                If arrInvNo.Contains(dtr("CreditNoteNo").ToString) = False Then arrInvNo.Add(dtr("CreditNoteNo").ToString)
            End While
            dtr.Close()
            dtr.Dispose()

            For i = 0 To arrInvNo.Count - 1
                Try

                    dtr = ReadRecord("Select * from CreditNote where CreditNoteNo =  " & SafeSQL(arrInvNo(i)) & " Order by CreditNoteNo")

                    While dtr.Read
                        ExecuteNavSQL("Delete from CreditNote where CreditNoteNo=" & SafeSQL(dtr("CreditNoteNo").ToString))
                        ExecuteNavSQL("Delete from CreditNoteDet where CreditNoteNo=" & SafeSQL(dtr("CreditNoteNo").ToString))

                        sQry = "Insert into CreditNote (CreditNoteNo,CreditDate,CustNo,GoodsReturnNo,SalesPersonCode,SubTotal,Gst,TotalAmt,PaidAmt) Values (" & SafeSQL(dtr("CreditNoteNo").ToString) & _
                            "," & SafeSQL(Format(dtr("CreditDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("CustNo").ToString) & "," & SafeSQL(dtr("GoodsReturnNo").ToString) & "," & SafeSQL(dtr("SalesPersonCode").ToString) & _
                            "," & SafeSQL(dtr("SubTotal").ToString) & "," & SafeSQL(dtr("Gst").ToString) & "," & SafeSQL(dtr("TotalAmt").ToString) & "," & SafeSQL(dtr("PaidAmt").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export CreditNote - " & dtr("CreditNoteNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteNavSQL(sQry)

                    End While
                    dtr.Close()

                Catch ex As Exception
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export CreditNote - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try

                Try
                    dtr1 = ReadRecord("Select * From CreditNoteDet  Where CreditNoteNo =  " & SafeSQL(arrInvNo(i)) & " Order by CreditNoteNo")
                    While dtr1.Read
                        sSql = "Insert into CreditNoteDet (CreditNoteNo,ItemNo,UOM,BaseUOM,Price,Qty,Amt,[LineNo]) Values (" & SafeSQL(dtr1("CreditNoteNo").ToString) & _
                             "," & SafeSQL(dtr1("ItemNo").ToString) & "," & SafeSQL(dtr1("UOM").ToString) & "," & SafeSQL(dtr1("BaseUOM").ToString) & "," & SafeSQL(dtr1("Price").ToString) & _
                             "," & SafeSQL(dtr1("Qty").ToString) & "," & SafeSQL(dtr1("Amt").ToString) & "," & SafeSQL(dtr1("LineNo").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export CreditNoteDet - " & SafeSQL(arrInvNo(i))) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sSql) & ")")


                        ExecuteNavSQL(sSql)
                    End While
                    dtr1.Close()
                    ExecuteSQLAnother("Update CreditNote Set Exported = 1 Where CreditNoteNo = " & SafeSQL(arrInvNo(i)))
                Catch ex As Exception
                    dtr1.Close()
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export CreditNoteDet'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                    System.IO.File.AppendAllText(Application.StartupPath & "\ErrorLog.txt", "Export CreditNoteDet  Error" & "   " & Date.Now.ToString & " " & ex.Message & vbCrLf)
                    Exit Sub
                End Try
            Next
        Catch ex As Exception
            dtr.Close()
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export CreditNote'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try
    End Sub
    Private Sub ExportPayment()
        Dim sCurCode As String = ""
        Dim arrInvNo = New ArrayList
        Dim sQry As String = ""
        Dim dTransport As Double = 0
        sCurCode = "" 'dtr1("LCY Code")
        Dim dExRate As Double = 0
        Dim dtr As SqlDataReader
        Dim dtr1 As SqlDataReader
        Dim sExtDocNo As String = ""
        Dim sSql As String
        Try
            dtr = ReadRecord("Select  RcptNo from Receipt where isnull(Exported,0) = 0 Order by RcptNo")
            While dtr.Read
                If arrInvNo.Contains(dtr("RcptNo").ToString) = False Then arrInvNo.Add(dtr("RcptNo").ToString)
            End While
            dtr.Close()
            dtr.Dispose()


            For i = 0 To arrInvNo.Count - 1
                Try

                    dtr = ReadRecord("Select * from Receipt where RcptNo =  " & SafeSQL(arrInvNo(i)) & " Order by RcptNo")

                    While dtr.Read
                        ExecuteNavSQL("Delete from Receipt where RcptNo=" & SafeSQL(dtr("RcptNo").ToString))
                        ExecuteNavSQL("Delete from RcptItem where RcptNo=" & SafeSQL(dtr("RcptNo").ToString))

                        sQry = "Insert into Receipt (RcptNo,RcptDt,CustId,AgentId,PayMethod,ChqNo,ChqDt,Amount,Void,DTG,BankName) Values (" & SafeSQL(dtr("RcptNo").ToString) & _
                            "," & SafeSQL(Format(dtr("RcptDt"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("CustId").ToString) & "," & SafeSQL(dtr("AgentId").ToString) & "," & SafeSQL(dtr("PayMethod").ToString) & _
                            "," & SafeSQL(dtr("ChqNo").ToString) & "," & SafeSQL(Format(dtr("ChqDt"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("Amount").ToString) & "," & SafeSQL(dtr("Void").ToString) & "," & SafeSQL(Format(dtr("DTG"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("BankName").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export Receipt - " & dtr("RcptNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteNavSQL(sQry)

                    End While
                    dtr.Close()

                Catch ex As Exception
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export Receipt - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try

                Try
                    dtr1 = ReadRecord("Select * From RcptItem  Where RcptNo =  " & SafeSQL(arrInvNo(i)) & " Order by RcptNo")
                    While dtr1.Read
                        sSql = "Insert into RcptItem (RcptNo,InvNo,AmtPaid) Values (" & SafeSQL(dtr1("RcptNo").ToString) & _
                             "," & SafeSQL(dtr1("InvNo").ToString) & "," & SafeSQL(dtr1("AmtPaid").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export RcptItem - " & SafeSQL(arrInvNo(i))) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sSql) & ")")


                        ExecuteNavSQL(sSql)
                    End While
                    dtr1.Close()
                    ExecuteSQLAnother("Update Receipt Set Exported = 1 Where RcptNo = " & SafeSQL(arrInvNo(i)))
                Catch ex As Exception
                    dtr1.Close()
                    System.IO.File.AppendAllText(Application.StartupPath & "\ErrorLog.txt", "Export Inv Item Error" & "   " & Date.Now.ToString & " " & ex.Message & vbCrLf)
                    Exit Sub
                End Try
            Next
        Catch ex As Exception
            dtr.Close()
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export RcptItem'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try
    End Sub
    Private Sub ExportStockOrder()
        Dim sCurCode As String
        Dim arrInvNo = New ArrayList
        Dim sQry As String = ""
        Dim dTransport As Double = 0
        sCurCode = "" 'dtr1("LCY Code")
        Dim dExRate As Double = 0
        Dim dtr As SqlDataReader
        Dim dtr1 As SqlDataReader
        Dim sExtDocNo As String = ""
        Dim sSql As String
        Try
            dtr = ReadRecord("Select  StockNo from StockOrder where isnull(Exported,0) = 0 Order by StockNo")
            While dtr.Read
                If arrInvNo.Contains(dtr("StockNo").ToString) = False Then arrInvNo.Add(dtr("StockNo").ToString)
            End While
            dtr.Close()
            dtr.Dispose()


            For i = 0 To arrInvNo.Count - 1
                Try

                    dtr = ReadRecord("Select * from StockOrder where StockNo =  " & SafeSQL(arrInvNo(i)) & " Order by StockNo")

                    While dtr.Read
                        ExecuteNavSQL("Delete from StockOrder where StockNo=" & SafeSQL(dtr("StockNo").ToString))
                        ExecuteNavSQL("Delete from StockOrderItem where StockNo=" & SafeSQL(dtr("StockNo").ToString))

                        sQry = "Insert into StockOrder (StockNo,OrdDt,TrnDate,Location,AgentId,Remarks) Values (" & SafeSQL(dtr("StockNo").ToString) & _
                            "," & SafeSQL(Format(dtr("OrdDt"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(Format(dtr("TrnDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("Location").ToString) & "," & SafeSQL(dtr("AgentId").ToString) & _
                            "," & SafeSQL(dtr("Remarks").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export StockOrder - " & dtr("StockNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteNavSQL(sQry)

                    End While
                    dtr.Close()

                Catch ex As Exception
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export StockOrder - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try

                Try
                    dtr1 = ReadRecord("Select * From StockOrderItem  Where StockNo =  " & SafeSQL(arrInvNo(i)) & " Order by StockNo")
                    While dtr1.Read
                        sSql = "Insert into StockOrderItem (StockNo,ItemNo,UOM,Qty,Location,[LineNo]) Values (" & SafeSQL(dtr1("StockNo").ToString) & _
                             "," & SafeSQL(dtr1("ItemNo").ToString) & "," & SafeSQL(dtr1("UOM").ToString) & "," & SafeSQL(dtr1("Qty").ToString) & "," & SafeSQL(dtr1("Location").ToString) & _
                             "," & SafeSQL(dtr1("LineNo").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export StockOrderItem - " & SafeSQL(arrInvNo(i))) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sSql) & ")")


                        ExecuteNavSQL(sSql)
                    End While
                    dtr1.Close()
                    ExecuteSQLAnother("Update StockOrder Set Exported = 1 Where StockNo = " & SafeSQL(arrInvNo(i)))
                Catch ex As Exception
                    dtr1.Close()
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export StockOrderItem'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                    System.IO.File.AppendAllText(Application.StartupPath & "\ErrorLog.txt", "Export StockOrderItem Error" & "   " & Date.Now.ToString & " " & ex.Message & vbCrLf)
                    Exit Sub
                End Try
            Next
        Catch ex As Exception
            dtr.Close()
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export StockOrderItem'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try
    End Sub
    Private Sub ExportReturn()
        Dim sCurCode As String
        Dim arrInvNo = New ArrayList
        Dim sQry As String = ""
        Dim dTransport As Double = 0
        sCurCode = "" 'dtr1("LCY Code")
        Dim dExRate As Double = 0
        Dim dtr As SqlDataReader
        Dim dtr1 As SqlDataReader
        Dim sExtDocNo As String = ""
        Dim sSql As String
        Try
            dtr = ReadRecord("Select  RtnNo from [Return] where isnull(Exported,0) = 0 Order by RtnNo")
            While dtr.Read
                If arrInvNo.Contains(dtr("RtnNo").ToString) = False Then arrInvNo.Add(dtr("RtnNo").ToString)
            End While
            dtr.Close()
            dtr.Dispose()


            For i = 0 To arrInvNo.Count - 1
                Try
                    dtr = ReadRecord("Select * from [Return] where RtnNo =  " & SafeSQL(arrInvNo(i)) & " Order by RtnNo")

                    While dtr.Read
                        ExecuteNavSQL("Delete from [Return] where RtnNo =" & SafeSQL(dtr("RtnNo").ToString))
                        ExecuteNavSQL("Delete from ReturnDet where RtnNo=" & SafeSQL(dtr("RtnNo").ToString))

                        sQry = "Insert into [Return] (RtnNo,RtnDate,CustId,AgentId,SubTotal,GST,GSTAmt,TotalAmt,PaidAmt,Void,VoidDate,Exported,ExportDate,Isconfirmed,ConfirmedBy) Values (" & SafeSQL(dtr("RtnNo").ToString) & _
                            "," & SafeSQL(dtr("RtnNo").ToString) & "," & SafeSQL(Format(dtr("RtnDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("CustId").ToString) & "," & SafeSQL(dtr("AgentId").ToString) & _
                             "," & SafeSQL(dtr("SubTotal").ToString) & "," & SafeSQL(dtr("GST").ToString) & "," & SafeSQL(dtr("GSTAmt").ToString) & "," & SafeSQL(dtr("TotalAmt").ToString) & _
                              "," & SafeSQL(dtr("PaidAmt").ToString) & "," & SafeSQL(dtr("Void").ToString) & "," & SafeSQL(Format(dtr("VoidDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("Exported").ToString) & "," & SafeSQL(Format(dtr("ExportedDate"), "yyyyMMdd HH:mm:ss")) & _
                               "," & SafeSQL(dtr("IsConfirmed").ToString) & "," & SafeSQL(dtr("ConfirmedBy").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export Return - " & dtr("RtnNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteNavSQL(sQry)

                    End While
                    dtr.Close()

                Catch ex As Exception
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export Return - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try

                Try
                    dtr1 = ReadRecord("Select * From ReturnDet  Where RtnNo =  " & SafeSQL(arrInvNo(i)) & " Order by RtnNo")
                    While dtr1.Read
                        sSql = "Insert into ReturnDet (RtnNo,ItemNo,UOM,Qty,LineNo,Description,UnitPrice,TotalAmt,Location,Remarks) Values (" & SafeSQL(dtr1("RtnNo").ToString) & _
                             "," & SafeSQL(dtr1("ItemNo").ToString) & "," & SafeSQL(dtr1("UOM").ToString) & "," & SafeSQL(dtr1("Qty").ToString) & "," & SafeSQL(dtr1("LineNo").ToString) & _
                             "," & SafeSQL(dtr1("Description").ToString) & "," & SafeSQL(dtr("UnitPrice").ToString) & "," & SafeSQL(dtr("TotalAmt").ToString) & "," & SafeSQL(dtr("Location").ToString) & _
                              "," & SafeSQL(dtr("Remarks").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export ReturnDet - " & SafeSQL(arrInvNo(i))) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sSql) & ")")

                        ExecuteNavSQL(sSql)
                    End While
                    dtr1.Close()
                    ExecuteSQLAnother("Update [Return] Set Exported = 1 Where RtnNo = " & SafeSQL(arrInvNo(i)))
                Catch ex As Exception
                    dtr1.Close()
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export ReturnDet'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                    System.IO.File.AppendAllText(Application.StartupPath & "\ErrorLog.txt", "Export Inv Item Error" & "   " & Date.Now.ToString & " " & ex.Message & vbCrLf)
                    Exit Sub
                End Try
            Next
        Catch ex As Exception
            dtr.Close()
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export ReturnDet'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try
    End Sub
    Private Sub ExportItemTrans()
        Dim sCurCode As String
        Dim arrInvNo = New ArrayList
        Dim sQry As String = ""
        Dim dTransport As Double = 0
        sCurCode = "" 'dtr1("LCY Code")
        Dim dExRate As Double = 0
        Dim dtr As SqlDataReader
        Dim dtr1 As SqlDataReader
        Dim sExtDocNo As String = ""
        Dim sSql As String
        Try
            dtr = ReadRecord("Select  DocNo from ItemTrans where isnull(Exported,0) = 0 Order by DocNo")
            While dtr.Read
                If arrInvNo.Contains(dtr("DocNo").ToString) = False Then arrInvNo.Add(dtr("DocNo").ToString)
            End While
            dtr.Close()
            dtr.Dispose()


            For i = 0 To arrInvNo.Count - 1
                Try
                    dtr = ReadRecord("Select * from ItemTrans where DocNo =  " & SafeSQL(arrInvNo(i)) & " Order by DocNo")

                    While dtr.Read
                        ExecuteNavSQL("Delete from ItemTrans where DocNo =" & SafeSQL(dtr("DocNo").ToString))
                        ExecuteNavSQL("Delete from ItemTransExp where DocNo=" & SafeSQL(dtr("DocNo").ToString))

                        sQry = "Insert into ItemTrans (DocNo,DocDt,DocType,Location,ItemId,UOM,Qty,Exported,Remarks,IsUpdated) Values (" & SafeSQL(dtr("DocNo").ToString) & _
                            "," & SafeSQL(Format(dtr("DocDt"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("DocType").ToString) & "," & SafeSQL(dtr("Location").ToString) & "," & SafeSQL(dtr("ItemId").ToString) & _
                             "," & SafeSQL(dtr("UOM").ToString) & "," & SafeSQL(dtr("Qty").ToString) & "," & SafeSQL(dtr("Exported").ToString) & "," & SafeSQL(dtr("Remarks").ToString) & _
                              "," & SafeSQL(dtr("IsUpdated").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export ItemTrans - " & dtr("DocNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteNavSQL(sQry)

                    End While
                    dtr.Close()

                Catch ex As Exception
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export ItemTrans - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try

                Try
                    dtr1 = ReadRecord("Select * From ItemTransExp  Where DocNo =  " & SafeSQL(arrInvNo(i)) & " Order by DocNo")
                    While dtr1.Read
                        sSql = "Insert into ItemTransExp (DocNo,ItemNo,UOM,LotNo,Qty,LineNo,ExpiryDate) Values (" & SafeSQL(dtr1("DocNo").ToString) & _
                             "," & SafeSQL(dtr1("ItemNo").ToString) & "," & SafeSQL(dtr1("UOM").ToString) & "," & SafeSQL(dtr1("LotNo").ToString) & "," & SafeSQL(dtr1("Qty").ToString) & _
                             "," & SafeSQL(dtr1("LineNo").ToString) & "," & SafeSQL(dtr("ExpiryDate").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export ItemTransExp - " & SafeSQL(arrInvNo(i))) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sSql) & ")")

                        ExecuteNavSQL(sSql)
                    End While
                    dtr1.Close()
                    ExecuteSQLAnother("Update ItemTrans Set Exported = 1 Where DocNo = " & SafeSQL(arrInvNo(i)))
                Catch ex As Exception
                    dtr1.Close()
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export ItemTransExp'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                    System.IO.File.AppendAllText(Application.StartupPath & "\ErrorLog.txt", "Export ItemTransExp Error" & "   " & Date.Now.ToString & " " & ex.Message & vbCrLf)
                    Exit Sub
                End Try
            Next
        Catch ex As Exception
            dtr.Close()
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export ItemTransExp'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try
    End Sub
    Private Sub ExportCustVisit()
        Dim sCurCode As String
        Dim arrInvNo = New ArrayList
        Dim sQry As String = ""
        Dim dTransport As Double = 0
        sCurCode = "" 'dtr1("LCY Code")
        Dim dExRate As Double = 0
        Dim dtr As SqlDataReader
        Dim dtr1 As SqlDataReader
        Dim sExtDocNo As String = ""
        Dim sSql As String
        Try
            dtr = ReadRecord("Select  CustId from CustVisit where isnull(Exported,0) = 0 Order by CustId")
            While dtr.Read
                If arrInvNo.Contains(dtr("CustId").ToString) = False Then arrInvNo.Add(dtr("CustId").ToString)
            End While
            dtr.Close()
            dtr.Dispose()


            For i = 0 To arrInvNo.Count - 1
                Try
                    dtr = ReadRecord("Select * from CustVisit where CustId =  " & SafeSQL(arrInvNo(i)) & " Order by CustId")

                    While dtr.Read
                        ExecuteNavSQL("Delete from CustVisit where CustId =" & SafeSQL(dtr("CustId").ToString))

                        sQry = "Insert into CustVisit (CustId,TransNo,TransType,TransDate,AgentId,Status,Latitude,Longitude,Remarks,DTG) Values (" & SafeSQL(dtr("CustId").ToString) & _
                            "," & SafeSQL(dtr("TransNo").ToString) & "," & SafeSQL(dtr("TransType").ToString) & "," & SafeSQL(Format(dtr("TransDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("AgentId").ToString) & _
                        "," & SafeSQL(dtr("Status").ToString) & "," & SafeSQL(dtr("Latitude").ToString) & "," & SafeSQL(dtr("Longitude").ToString) & "," & SafeSQL(dtr("Remarks").ToString) & ", GetDate())"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export CustVisit - " & dtr("CustId").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteNavSQL(sQry)

                    End While
                    dtr.Close()
                    ExecuteSQLAnother("Update CustVisit Set Exported = 1 Where CustId = " & SafeSQL(arrInvNo(i)))
                Catch ex As Exception
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export CustVisit - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try
            Next
        Catch ex As Exception
            dtr.Close()
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export CustVisit'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try
    End Sub
    Private Sub ExportException()
        Dim sCurCode As String
        Dim arrInvNo = New ArrayList
        Dim sQry As String = ""
        Dim dTransport As Double = 0
        sCurCode = "" 'dtr1("LCY Code")
        Dim dExRate As Double = 0
        Dim dtr As SqlDataReader
        'Dim dtr1 As SqlDataReader
        Dim sExtDocNo As String = ""
        Dim sSql As String
        Try
            dtr = ReadRecord("Select  CustId from Exception where isnull(Exported,0) = 0 Order by CustId")
            While dtr.Read
                If arrInvNo.Contains(dtr("CustId").ToString) = False Then arrInvNo.Add(dtr("CustId").ToString)
            End While
            dtr.Close()
            dtr.Dispose()


            For i = 0 To arrInvNo.Count - 1
                Try
                    dtr = ReadRecord("Select * from Exception where CustId =  " & SafeSQL(arrInvNo(i)) & " Order by CustId")

                    While dtr.Read
                        ExecuteNavSQL("Delete from Exception where CustId =" & SafeSQL(dtr("CustId").ToString))

                        sQry = "Insert into Exception (CustId,DocNo,DocType,AgentId,DocDate) Values (" & SafeSQL(dtr("CustId").ToString) & _
                            "," & SafeSQL(dtr("DocNo").ToString) & "," & SafeSQL(dtr("DocType").ToString) & "," & SafeSQL(dtr("AgentId").ToString) & "," & SafeSQL(Format(dtr("DocDate"), "yyyyMMdd HH:mm:ss")) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export Exception - " & dtr("CustId").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteNavSQL(sQry)

                    End While
                    dtr.Close()
                    ExecuteSQLAnother("Update Exception Set Exported = 1 Where CustId = " & SafeSQL(arrInvNo(i)))
                Catch ex As Exception
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export Exception - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try
            Next
        Catch ex As Exception
            dtr.Close()
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export Exception'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try
    End Sub
    Private Sub ExportBank()
        Dim sCurCode As String
        Dim arrInvNo = New ArrayList
        Dim sQry As String = ""
        Dim dTransport As Double = 0
        sCurCode = "" 'dtr1("LCY Code")
        Dim dExRate As Double = 0
        Dim dtr As SqlDataReader
        Dim dtr1 As SqlDataReader
        Dim sExtDocNo As String = ""
        Dim sSql As String
        Try
            dtr = ReadRecord("Select  DocNo from BankInHdr where isnull(Exported,0) = 0 Order by DocNo")
            While dtr.Read
                If arrInvNo.Contains(dtr("DocNo").ToString) = False Then arrInvNo.Add(dtr("DocNo").ToString)
            End While
            dtr.Close()
            dtr.Dispose()

            For i = 0 To arrInvNo.Count - 1
                Try
                    dtr = ReadRecord("Select * from BankInHdr where DocNo =  " & SafeSQL(arrInvNo(i)) & " Order by DocNo")

                    While dtr.Read
                        ExecuteNavSQL("Delete from BankInHdr where DocNo =" & SafeSQL(dtr("DocNo").ToString))
                        ExecuteNavSQL("Delete from BankInDet where DocNo=" & SafeSQL(dtr("DocNo").ToString))

                        sQry = "Insert into BankInHdr (DocNo,SlipNo,DocDate,DocType,AgentId,Amount,BankAccount,Remarks,Exported,MDTNo,Void) Values (" & SafeSQL(dtr("DocNo").ToString) & _
                            "," & SafeSQL(dtr("SlipNo").ToString) & "," & SafeSQL(Format(dtr("DocDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("DocType").ToString) & "," & SafeSQL(dtr("AgentId").ToString) & _
                             "," & SafeSQL(dtr("Amount").ToString) & "," & SafeSQL(dtr("BankAccount").ToString) & "," & SafeSQL(dtr("Remarks").ToString) & "," & SafeSQL(dtr("Exported").ToString) & _
                              "," & SafeSQL(dtr("MDTNo").ToString) & "," & SafeSQL(dtr("Void").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export BankInHdr - " & dtr("DocNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteNavSQL(sQry)

                    End While
                    dtr.Close()

                Catch ex As Exception
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export BankInHdr - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try

                Try
                    dtr1 = ReadRecord("Select * From BankInDet  Where DocNo =  " & SafeSQL(arrInvNo(i)) & " Order by DocNo")
                    While dtr1.Read
                        sSql = "Insert into BankInDet (DocNo,ReceiptNo,ChqNo,ChqDate,ChqAmount,BankName,Remarks) Values (" & SafeSQL(dtr1("DocNo").ToString) & _
                             "," & SafeSQL(dtr1("ReceiptNo").ToString) & "," & SafeSQL(dtr1("ChqNo").ToString) & "," & SafeSQL(Format(dtr1("ChqDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr1("BankName").ToString) & _
                             "," & SafeSQL(dtr1("Remarks").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export BankInDet - " & SafeSQL(arrInvNo(i))) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sSql) & ")")

                        ExecuteNavSQL(sSql)
                    End While
                    dtr1.Close()
                    ExecuteSQLAnother("Update BankInHdr Set Exported = 1 Where DocNo = " & SafeSQL(arrInvNo(i)))
                Catch ex As Exception
                    dtr1.Close()
                    System.IO.File.AppendAllText(Application.StartupPath & "\ErrorLog.txt", "BankInDet  Error" & "   " & Date.Now.ToString & " " & ex.Message & vbCrLf)
                    Exit Sub
                End Try
            Next
        Catch ex As Exception
            dtr.Close()
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export BankInDet'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try
    End Sub
    Private Sub ExportExchange()
        Dim sCurCode As String
        Dim arrInvNo = New ArrayList
        Dim sQry As String = ""
        Dim dTransport As Double = 0
        sCurCode = "" 'dtr1("LCY Code")
        Dim dExRate As Double = 0
        Dim dtr As SqlDataReader
        Dim dtr1 As SqlDataReader
        Dim sExtDocNo As String = ""
        Dim sSql As String
        Try
            dtr = ReadRecord("Select  ExchangeNo from GoodsExchange where isnull(Exported,0) = 0 Order by ExchangeNo")
            While dtr.Read
                If arrInvNo.Contains(dtr("ExchangeNo").ToString) = False Then arrInvNo.Add(dtr("ExchangeNo").ToString)
            End While
            dtr.Close()
            dtr.Dispose()


            For i = 0 To arrInvNo.Count - 1
                Try
                    dtr = ReadRecord("Select * from GoodsExchange where ExchangeNo =  " & SafeSQL(arrInvNo(i)) & " Order by ExchangeNo")

                    While dtr.Read
                        ExecuteNavSQL("Delete from GoodsExchange where ExchangeNo =" & SafeSQL(dtr("ExchangeNo").ToString))
                        ExecuteNavSQL("Delete from GoodsReturnItem where ReturnNo=" & SafeSQL(dtr("ExchangeNo").ToString))

                        sQry = "Insert into GoodsExchange (ExchangeNo,ExchangeDate,CustId,SalesPersonCode,SubTotal,Gst,GstAmt,TotalAmt,Approved,ApprovedBy) Values (" & SafeSQL(dtr("ExchangeNo").ToString) & _
                            "," & SafeSQL(Format(dtr("ExchangeDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("CustNo").ToString) & "," & SafeSQL(dtr("SalesPersonCode").ToString) & "," & SafeSQL(dtr("SubTotal").ToString) & _
                             "," & SafeSQL(dtr("Gst").ToString) & "," & SafeSQL(dtr("GstAmt").ToString) & "," & SafeSQL(dtr("TotalAmt").ToString) & "," & SafeSQL(dtr("Approved").ToString) & _
                              "," & SafeSQL(dtr("ApprovedBy").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export GoodsExchange - " & dtr("ExchangeNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteNavSQL(sQry)

                    End While
                    dtr.Close()

                Catch ex As Exception
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export GoodsExchange - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try

                Try
                    dtr1 = ReadRecord("Select * From GoodsReturnItem  Where ReturnNo =  " & SafeSQL(arrInvNo(i)) & " Order by ReturnNo")
                    While dtr1.Read
                        sSql = "Insert into GoodsReturnItem (ReturnNo,ItemNo,UOM,Quantity,Remarks,Price,Amt,CustProdCode,GoodQty,ReasonCode,LineNo) Values (" & SafeSQL(dtr1("ReturnNo").ToString) & _
                             "," & SafeSQL(dtr1("ItemNo").ToString) & "," & SafeSQL(dtr1("UOM").ToString) & "," & SafeSQL(dtr1("Quantity").ToString) & "," & SafeSQL(dtr1("Remarks").ToString) & _
                             "," & SafeSQL(dtr1("Price").ToString) & "," & SafeSQL(dtr1("Amt").ToString) & "," & SafeSQL(dtr1("CustProdCode").ToString) & "," & SafeSQL(dtr1("GoodQty").ToString) & _
                             "," & SafeSQL(dtr1("ReasonCode").ToString) & "," & SafeSQL(dtr1("LineNo").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export GoodsReturnItem - " & SafeSQL(arrInvNo(i))) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sSql) & ")")

                        ExecuteNavSQL(sSql)
                    End While
                    dtr1.Close()
                    ExecuteSQLAnother("Update GoodsExchange Set Exported = 1 Where ExchangeNo = " & SafeSQL(arrInvNo(i)))
                Catch ex As Exception
                    dtr1.Close()
                    System.IO.File.AppendAllText(Application.StartupPath & "\ErrorLog.txt", "GoodsReturnItem  Error" & "   " & Date.Now.ToString & " " & ex.Message & vbCrLf)
                    Exit Sub
                End Try
            Next
        Catch ex As Exception
            dtr.Close()
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export GoodsReturnItem'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try
    End Sub
    Private Sub ExportService()
        Dim sCurCode As String
        Dim arrInvNo = New ArrayList
        Dim sQry As String = ""
        Dim dTransport As Double = 0
        sCurCode = "" 'dtr1("LCY Code")
        Dim dExRate As Double = 0
        Dim dtr As SqlDataReader
        'Dim dtr1 As SqlDataReader
        Dim sExtDocNo As String = ""
        Dim sSql As String
        Try
            dtr = ReadRecord("Select  ServiceId from Service where isnull(Exported,0) = 0 Order by ServiceId")
            While dtr.Read
                If arrInvNo.Contains(dtr("ServiceId").ToString) = False Then arrInvNo.Add(dtr("ServiceId").ToString)
            End While
            dtr.Close()
            dtr.Dispose()

            For i = 0 To arrInvNo.Count - 1
                Try
                    dtr = ReadRecord("Select * from Service where ServiceId =  " & SafeSQL(arrInvNo(i)) & " Order by ServiceId")

                    While dtr.Read
                        ExecuteNavSQL("Delete from Service where ServiceId =" & SafeSQL(dtr("ServiceId").ToString))

                        sQry = "Insert into Service (ServiceId,ServiceDt,Details,CustId,AgentId,ReasonCode) Values (" & SafeSQL(dtr("ServiceId").ToString) & _
                            "," & SafeSQL(dtr("ServiceDt").ToString) & "," & SafeSQL(dtr("Details").ToString) & "," & SafeSQL(dtr("CustId").ToString) & "," & SafeSQL(dtr("AgentId").ToString) & _
                            "," & SafeSQL(dtr("ReasonCode").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export Service - " & dtr("ServiceId").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteNavSQL(sQry)

                    End While
                    dtr.Close()

                Catch ex As Exception
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export Service - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try
            Next
        Catch ex As Exception
            dtr.Close()
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export Exception'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try
    End Sub
    Public Sub ImportSalesAgent()
        Dim dtr As SqlDataReader
        'ExecuteSQL("Update SalesAgent Set Active = 0")
        dtr = ReadNavRecord("Select Distinct AgentID as SalesAgent from CustAgent  order by AgentID")

        While dtr.Read
            If dtr("SalesAgent").ToString <> "" Then

                If IsAgentExists(dtr("SalesAgent").ToString) = False Then
                    ExecuteSQLAnother("Insert into SalesAgent(Code,Name, UserID, Password, RouteNo, VehicleID, SolutionName) Values (" & SafeSQL(dtr("SalesAgent").ToString) & "," & SafeSQL(dtr("SalesAgent").ToString) & "," & SafeSQL("") & "," & SafeSQL(dtr("SalesAgent").ToString) & "," & SafeSQL("") & "," & SafeSQL("") & ", 'SALES')")
                Else
                    ExecuteSQLAnother("Update SalesAgent Set Name = " & SafeSQL(dtr("SalesAgent").ToString) & ", Active = 1 Where Code = " & SafeSQL(dtr("SalesAgent").ToString))
                End If

            End If
        End While
        dtr.Close()
        dtr.Dispose()
        dtr = Nothing

        'ExecuteSQL("Update SalesAgent Set Active = 1 where Code = 'ADMIN'")


    End Sub

    Public Function GetGeoCoords(ByVal inString As String, ByVal inType As Integer) As String
        Try
            ' Explanation of function:
            ' Use inType=0 and feed in a specific Google Maps URL to parse out the GeoCoords from the URL
            ' e.g. http://maps.google.com/maps?f=q&source=s_q&hl=en&geocode=&q=53154&sll=37.0625,-95.677068&sspn=52.505328,80.507812&ie=UTF8&ll=42.858224,-88.000832&spn=0.047943,0.078621&t=h&z=14
            ' Function returns a string of geocoords (e.g. "-87.9010610,42.8864960")
            '
            ' Use inType=1 and feed in a zip code, address, or business name
            ' Function returns a string of geocoords (e.g. "-87.9010610,42.8864960")
            ' If an invalid address, zip code or location was entered, the function will return "0,0"
            Dim skey As String = "AIzaSyC5XH-ALvr81IiuEDWekI2k91ujPeZL864"
            Dim Chunks As String()
            Dim outString As String = ""

            If inType = 0 Then
                Chunks = Regex.Split(inString, "&")
                For Each s As String In Chunks
                    If InStr(s, "ll") > 0 Then outString = s
                Next
                outString = Replace(Replace(outString, "sll=", ""), "ll=", "")
            Else
                'Dim xmlString As String = GetHTML("http://maps.google.com/maps/geo?output=xml&key=abcdefg&q=" & inString, 1)
                Dim xmlString As String = GetHTML("https://maps.googleapis.com/maps/api/geocode/xml?address=" & inString & "&sensor=true_or_false&key=" & skey, 1)
                Dim pos1 As Integer = 0
                Dim pos2 As Integer = 0
                pos1 = xmlString.IndexOf("<location>")
                pos2 = xmlString.IndexOf("</location>")
                Dim stmp As String = ""
                If pos1 > 0 And pos2 > pos1 Then
                    stmp = xmlString.Substring(pos1, pos2 - pos1).Trim
                    pos1 = 0
                    pos2 = 0
                    pos1 = stmp.IndexOf("<lat>")
                    pos2 = stmp.IndexOf("</lat>")
                    outString = stmp.Substring(pos1, pos2 - pos1).Trim.Trim.Replace("<lat>", "")
                    pos1 = stmp.IndexOf("<lng>")
                    pos2 = stmp.IndexOf("</lng>")
                    outString = outString & "," & stmp.Substring(pos1, pos2 - pos1).Trim.Replace("<lng>", "")
                End If

                'Chunks = Regex.Split(xmlString, "coordinates>", RegexOptions.Multiline)
                'If Chunks.Length > 1 Then
                '    outString = Replace(Chunks(1), ",0</", "")
                'Else
                '    outString = "0,0"
                'End If

            End If
            Return outString
        Catch ex As Exception
            Return "0,0"
        End Try
    End Function


    Public Function GetHTML(ByVal sURL As String, ByVal e As Integer) As String

        Dim oHttpWebRequest As System.Net.HttpWebRequest
        Dim oStream As System.IO.Stream
        Dim sChunk As String
        oHttpWebRequest = (System.Net.HttpWebRequest.Create(sURL))

        Dim oHttpWebResponse As System.Net.WebResponse = oHttpWebRequest.GetResponse()
        oStream = oHttpWebResponse.GetResponseStream
        sChunk = New System.IO.StreamReader(oStream).ReadToEnd()
        oStream.Close()
        oHttpWebResponse.Close()
        If e = 0 Then


            'Return Server.HtmlEncode(sChunk)
            Return System.Web.HttpUtility.HtmlEncode(sChunk)
        Else
            'Return Server.HtmlDecode(sChunk)
            Return System.Web.HttpUtility.HtmlDecode(sChunk)
        End If
    End Function


    Public Sub UpdateGPSCoordinates()
        Dim sSQL As String
        Dim arrCustNo As New ArrayList()
        Dim dtr As SqlDataReader
        arrCustNo.Clear()
        'and (Latitude=0 or Longitude=0 or Latitude is Null or Longitude is Null) 
        dtr = ReadRecord("Select Distinct Address  from Customer where Active=1 and isnull(address,'')<>'' order by Address")
        While dtr.Read = True
            arrCustNo.Add(dtr("Address").ToString)
        End While
        dtr.Close()

        For i = 0 To arrCustNo.Count - 1
            Try
                Dim sLoc As String = GetGeoCoords(arrCustNo(i) & ",  philippines ", 1)
                If sLoc <> "" Then
                    Dim S() As String = sLoc.Split(",")
                    sSQL = "UPDATE Customer SET Longitude= " & S(1) & " , Latitude=" & S(0) & " Where Address=" & SafeSQL(arrCustNo(i))
                    ExecuteSQL(sSQL)
                    'sb.Append(sLoc & vbCrLf)
                End If
            Catch ex As Exception

            End Try
        Next
    End Sub


    Public Sub ImportUOM()
        Dim dtr As SqlDataReader
        dtr = ReadNavRecord("Select  ""Item No_"", Code, ""Qty_ per Unit of Measure"" from """ & sNavCompanyName & "Item Unit of Measure"" A, """ & sNavCompanyName & "Item""  B Where A.""Item No_"" = B.No_ ")
        While dtr.Read
            If IsExists("Select ItemNo from UOM where ItemNo=" & SafeSQL(dtr("Item No_").ToString.Trim) & " and Uom=" & SafeSQL(dtr("Code").ToString.Trim)) = False Then
                ExecuteSQLAnother("Insert into UOM(ItemNo, Uom , BaseQty, DTG) Values (" & SafeSQL(dtr("Item No_").ToString) & "," & SafeSQL(dtr("Code").ToString) & "," & SafeSQL(dtr("Qty_ per Unit of Measure").ToString) & "," & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")) & ")")
            Else
                ExecuteSQLAnother("Update UOM set BaseQty=" & SafeSQL(dtr("Qty_ per Unit of Measure").ToString) & " Where ItemNo=" & SafeSQL(dtr("Item No_").ToString.Trim) & " and UOM=" & SafeSQL(dtr("Code").ToString.Trim))
            End If
        End While
        dtr.Close()
    End Sub

    Public Sub ImportPayMethod()

    End Sub


    Public Sub ImportPayterms()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Distinct PaymentTerms from Customer")
        While dtr.Read
            If IsExists("Select Code from PayTerms where code=" & SafeSQL(dtr("PaymentTerms"))) = False Then
                ExecuteSQLAnother("Insert into PayTerms(Code , Description, DueDateCalc, DisDateCalc, DiscountPercent,Active,DTG) Values (" & SafeSQL(dtr("PaymentTerms").ToString) & "," & SafeSQL(dtr("PaymentTerms").ToString) & "," & SafeSQL(dtr("PaymentTerms").ToString + "D") & "," & SafeSQL("0D") & ",0,1," & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")) & ")")
            Else
                ExecuteSQLAnother("Update PayTerms set Description=" & SafeSQL(dtr("PaymentTerms").ToString) & ", DueDateCalc =" & SafeSQL(dtr("PaymentTerms").ToString + "D") & " Where Code=" & SafeSQL(dtr("GroupNum").ToString))
            End If
        End While
        dtr.Close()
        dtr.Dispose()
        dtr = Nothing
    End Sub


    Public Sub ImportPriceGroup()
        Dim dtr As SqlDataReader
        dtr = ReadNavRecord("Select * from OPLN")
        While dtr.Read
            If IsExists("Select Code from PriceGroup where code=" & SafeSQL(dtr("ListNum"))) = False Then
                ExecuteSQLAnother("Insert into PriceGroup(Code , Description, InvoiceDiscount, LineDiscount, DTG) Values (" & SafeSQL(dtr("ListNum").ToString) & "," & SafeSQL(dtr("ListName").ToString) & ",1,1," & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")) & ")")
            Else
                ExecuteSQLAnother("Update PriceGroup set Description=" & SafeSQL(IIf(dtr("ListName").ToString = "", dtr("ListNum").ToString, dtr("ListName").ToString)) & " ,DTG=GetDate() Where Code=" & SafeSQL(dtr("ListNum").ToString))
            End If
        End While
        dtr.Close()
        dtr.Dispose()
        dtr = Nothing
    End Sub

    Public Sub ImportInvoice()
        Dim dtr As SqlDataReader
        Dim rs As SqlDataReader
        Dim dNewRecord As Int16 = 0
        Dim iValue As Date = Date.Now
        Dim ivalueLast30 As Date = iValue.AddDays(-30)
        Dim bSync As Boolean = GetLastTimeStamp("Invoice", iValue, dNewRecord)
        Dim arrList As New ArrayList
        If dNewRecord = 0 Then
            'dtr = ReadNavRecord("SELECT * FROM OInv where (UpdateDate >= " & SafeSQL(Format(ivalueLast30, "yyyyMMdd")))
            dtr = ReadNavRecord("Select * from OInv where (UpdateDate >= " & SafeSQL(Format(ivalueLast30, "yyyyMMdd")) & " or UpdateDate is null) order by UpdateDate")
        Else
            dtr = ReadNavRecord("Select * from OInv where (UpdateDate >= " & SafeSQL(Format(iValue, "yyyyMMdd")) & " or UpdateDate is null) order by UpdateDate")
        End If
        While dtr.Read
            Dim dDisAmt As Double = 0
            rs = ReadRecord("Select PaidAmt from Invoice where InvNo= " & SafeSQL(dtr("DocNum")))
            If rs.Read = True Then
                rs.Close()
                ExecuteSQL("Update Invoice set PaidAmt = " & dtr("PaidToDate") & ", custid = " & SafeSQL(dtr("cardcode")) & ", DTG = " & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")) & " where InvNo =" & SafeSQL(dtr("DocNum")))
            Else
                Dim aInv As New ArrInvoice
                aInv.InvNo = dtr("DocNum").ToString
                aInv.InvDate = dtr("DocDate")
                aInv.OrdNo = dtr("Ref1").ToString
                aInv.CustID = dtr("CardCode").ToString
                aInv.AgentID = dtr("SlpCode").ToString
                aInv.Discount = 0
                aInv.CurCode = dtr("DocCur").ToString
                aInv.CurExRate = dtr("DocRate")
                '  inv()
                aInv.GSTAmt = dtr("VatSum")
                aInv.Subtotal = dtr("DocTotal") - dtr("VatSum")
                aInv.TotalAmt = dtr("DocTotal")
                aInv.PaidAmt = dtr("PaidToDate")
                aInv.payterms = dtr("GroupNum").ToString
                arrList.Add(aInv)
                rs.Close()
                ExecuteSQL("Insert into Invoice(InvNo, InvDt, DONo, DoDt ,OrdNo, CustId, AgentID, Discount, SubTotal, GSTAmt, TotalAmt, PaidAmt, Void, PrintNo, PayTerms, CurCode, CurExRate, PONo, Exported,DTG,CompanyName,AcBillRef) Values (" & SafeSQL(dtr("DocNum").ToString) & "," & SafeSQL(Format(dtr("DocDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("DocNum").ToString) & "," & SafeSQL(Format(dtr("DocDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("Ref1").ToString) & "," & SafeSQL(dtr("CardCode").ToString) & "," & SafeSQL(dtr("SlpCode").ToString) & "," & dDisAmt & ", " & dtr("DocTotal") - dtr("VatSum") & ", " & dtr("VatSum") & " , " & dtr("DocTotal") & ", " & dtr("PaidToDate") & ",0,1," & SafeSQL(dtr("GroupNum").ToString) & "," & SafeSQL(dtr("DocCur").ToString) & "," & IIf(IsDBNull(dtr("DocRate")), 1, dtr("DocRate")) & ",'',1," & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")) & "," & SafeSQL("STD") & "," & SafeSQL(dtr("CardCode").ToString) & ")")
            End If
        End While
        dtr.Close()
        dtr.Dispose()
        dtr = Nothing

        If rs Is Nothing = False Then
            rs.Dispose()
            rs = Nothing
        End If

        ImportInvItem(arrList)
        ExecuteSQLAnother("Update InvItem set UOM=Item.BaseUOM from InvItem, Item where InvItem.ItemNo= Item.ItemNo and (InvItem.UOM is Null or InvItem.UOM='')")
        'For iIndex As Integer = 0 To arrList.Count - 1
        '    Dim aP As New ArrInvoice
        '    aP = arrList(iIndex)
        '    rs = ReadRecord("Select SUM(SubAmt) as Amount, sum(gstamt) as Gst from InvItem where InvNo= " & SafeSQL(aP.InvNo))
        '    If rs.Read = True Then
        '        ExecuteSQLAnother("Update Invoice set SubTotal = " & IIf(IsDBNull(rs("Amount")), 0, rs("Amount")) & ", " & _
        '                          "Gstamt = " & IIf(IsDBNull(rs("Gst")), 0, rs("Gst")) & ", TotalAmt = " & IIf(IsDBNull(rs("Amount")), 0, rs("Amount")) + IIf(IsDBNull(rs("gst")), 0, rs("gst")) & " where InvNo =" & SafeSQL(aP.InvNo))
        '    End If
        '    rs.Close()
        'Next
        If bSync = True Then
            UpdateLastTimeStamp("Invoice", iValue)
        Else
            InsertLastTimeStamp("Invoice", iValue)
        End If
    End Sub

    Public Sub ImportPurchaseOrder()
        Dim dtr As SqlDataReader
        Dim rs As SqlDataReader
        Dim dNewRecord As Int16 = 0
        Dim iValue As Date = Date.Now
        Dim ivalueLast30 As Date = iValue.AddDays(-200)
        '  Dim ivalueLast30 As Date = iValue.AddDays(-30)
        Dim bSync As Boolean = GetLastTimeStamp("PurchaseOrder", iValue, dNewRecord)
        Dim arrList As New ArrayList
        If dNewRecord = 0 Then

            'dtr = ReadNavRecord("SELECT * FROM OPOR where (UpdateDate >= " & SafeSQL(Format(ivalueLast30, "yyyyMMdd")))
            dtr = ReadNavRecord("Select * from OPOR where (UpdateDate >= " & SafeSQL(Format(ivalueLast30, "yyyyMMdd")) & " or UpdateDate is null) order by UpdateDate")
        Else

            dtr = ReadNavRecord("Select * from OPOR where (UpdateDate >= " & SafeSQL(Format(iValue, "yyyyMMdd")) & " or UpdateDate is null) order by UpdateDate")
        End If
        While dtr.Read
            'MsgBox("Entered purchaseorder1")
            Dim dDisAmt As Double = 0
            rs = ReadRecord("Select PONo from POHDR where PONo= " & SafeSQL(dtr("DocNum")))
            If rs.Read = True Then
                rs.Close()
                'MsgBox("Entered purchaseorderif")
                ExecuteSQL("Update POHDR set PONo = " & dtr("DocNum") & ", Vendorid = " & SafeSQL(dtr("cardcode")) & ", DTG = " & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")) & " where PONo =" & SafeSQL(dtr("DocNum")))
            Else
                '    MsgBox("Entered purchaseorderelse")
                Dim aPO As New ArrPurchaseOrder
                aPO.PONo = dtr("DocNum").ToString
                aPO.DocEntry = dtr("DocEntry").ToString
                'aPO.PODt = Format(dtr("DocDate"), "yyyyMMdd HH:mm:ss")
                aPO.PODt = dtr("DocDate")
                aPO.VendorId = dtr("SlpCode").ToString
                aPO.AgentID = dtr("SlpCode").ToString
                aPO.Discount = 0
                aPO.CurCode = dtr("DocCur").ToString
                aPO.CurExRate = dtr("DocRate")
                ' aPO.DTG = Format(Date.Now, "yyyyMMdd HH:mm:ss")
                aPO.DTG = Date.Now
                aPO.PrintNo = ""
                aPO.GSTAmt = dtr("VatSum")
                aPO.SubTotal = dtr("DocTotal") - dtr("VatSum")
                aPO.TotalAmt = dtr("DocTotal")
                aPO.Void = dtr("Canceled")
                aPO.PayTerms = dtr("GroupNum").ToString
                aPO.Exported = ""
                'aPO.DeliveryDate = Format(dtr("DocDueDate"), "yyyyMMdd HH:mm:ss")
                aPO.DeliveryDate = dtr("DocDueDate")
                aPO.LocationCode = ""
                aPO.ExternalDocNo = ""
                aPO.Remarks = ""
                aPO.POType = ""
                aPO.ContainerNo = ""
                aPO.Department = ""
                aPO.ManufacturerCode = ""
                aPO.CompanyName = ""
                arrList.Add(aPO)

                rs.Close()
                '   MsgBox("entering insert")
                ExecuteSQL("Insert into POHdr(PONo, PODt, VendorId, AgentID, Discount, SubTotal, GSTAmt, TotalAmt, Void, PrintNo, PayTerms, CurCode, CurExRate, Exported,DTG, DeliveryDate,LocationCode, ExternalDocNo, Remarks, POType, ContainerNo, Department, ManufacturerCode, CompanyName) Values (" & SafeSQL(dtr("DocNum").ToString) & "," & SafeSQL(Format(dtr("DocDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("cardcode").ToString) & "," & SafeSQL(dtr("SlpCode").ToString) & "," & dDisAmt & ", 0, 0, 0, 0,1," & SafeSQL(dtr("GroupNum").ToString) & "," & SafeSQL(dtr("DocCur").ToString) & "," & dtr("DocRate").ToString & ",1," & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(Format(dtr("DocDueDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL("") & "," & SafeSQL("") & "," & SafeSQL("") & ",'Purchase'," & SafeSQL("") & "," & SafeSQL("") & "," & SafeSQL("") & "," & SafeSQL("") & ")")

            End If
        End While
        dtr.Close()
        dtr.Dispose()
        dtr = Nothing
        ImportPODET(arrList)
        If rs Is Nothing = False Then
            rs.Dispose()
            rs = Nothing
        End If

        '        ImportPODET(arrList)
        If bSync = True Then
            UpdateLastTimeStamp("PurchaseHeader", iValue)
        Else
            InsertLastTimeStamp("PurchaseHeader", iValue)
        End If
    End Sub

    Public Sub ImportPODET(ByVal arrList As ArrayList)
        'MsgBox(arrList.Count)

        'Dim sInvNos As String = "''"
        'For idx As Integer = 0 To arrList.Count - 1
        '    Dim aPr As ArrInvoice
        '    aPr = arrList(idx)
        '    sInvNos &= "," & SafeSQL(aPr.InvNo)
        'Next

        Dim dtr As SqlDataReader

        ' ExecuteSQL("Delete from PODET where PONo in (" & sInvNos & ")")

        For iIndex As Integer = 0 To arrList.Count - 1

            Dim aPo As ArrPurchaseOrder
            aPo = arrList(iIndex)
            '   MsgBox(aPo.PONo)


            ExecuteSQL("Delete from PODET where PONo= " & SafeSQL(aPo.PONo))

            dtr = ReadNavRecord("Select * from POR1 where DocEntry = " & SafeSQL(aPo.DocEntry) & " order by LineNum")
            While dtr.Read
                'MsgBox(dtr("DocEntry"))
                '  MsgBox(aPo.PONo)

                If dtr("DocEntry").ToString = aPo.DocEntry Then
                    ExecuteSQL("Insert into PODet(PONo, [LineNo], [AttachedToLineNo], ItemNo, UOM, Qty, Foc, Price, DisPer, DisPr, Discount, SubAmt, GstAmt, DeliQty, BaseQty, VariantCode, Description, OutLet, BinCode, Location, ParentNo, CompanyName) Values (" & SafeSQL(aPo.PONo.ToString) & ", " & SafeSQL(dtr("Linenum").ToString) & ", " & SafeSQL(dtr("Flags").ToString) & "," & SafeSQL(dtr("ItemCode").ToString) & ", " & SafeSQL(dtr("UomEntry").ToString) & ", " & SafeSQL(dtr("Quantity").ToString) & ", 0, " & SafeSQL(dtr("DiscPrcnt").ToString) & ", " & SafeSQL(dtr("DiscPrcnt").ToString) & ", 0, " & SafeSQL(dtr("DiscPrcnt")) & ", " & SafeSQL(dtr("LineTotal").ToString) & ", " & SafeSQL(dtr("TotalFrgn")) & ", " & SafeSQL(dtr("DelivrdQty").ToString) & ", " & "1" & "," & SafeSQL(dtr("OCRCode").ToString) & "," & SafeSQL(dtr("Dscription").ToString) & "," & SafeSQL("") & "," & SafeSQL(dtr("SlpCode").ToString) & "," & SafeSQL(dtr("WhsCode").ToString) & ",''," & SafeSQL("") & ")")
                    ' ExecuteSQL("Insert into InvItem(InvNo, [LineNo], ItemNo, UOM, Qty, Foc, Price, DisPer, DisPr, Discount, SubAmt, GstAmt, DeliQty, BaseUOM, BaseQty, Description, ColorRemarks) Values (" & SafeSQL(dtr("Ref1").ToString) & ", " & dtr("DocLineNum").ToString & ", " & SafeSQL(dtr("ItemCode").ToString) & ", " & SafeSQL("") & ", " & dtr("OutQty").ToString & ", 0, " & dtr("Price").ToString & ", 0, 0, 0, " & dtr("OutQty").ToString * dtr("Price").ToString & ", 0, " & dtr("OutQty").ToString & ", " & SafeSQL("") & ", 1," & SafeSQL(dtr("Dscription").ToString) & ",'')")
                    'Exit For
                End If
            End While
            dtr.Close()
        Next
        If dtr Is Nothing = False Then
            dtr.Dispose()
            dtr = Nothing
        End If
        ExecuteSQLAnother("Update PODet set UOM=Item.BaseUOM from PODEt, Item where PODEt.ItemNo= Item.ItemNo")

    End Sub
    Public Sub ImportTransferOrder()
        Dim dtr As SqlDataReader
        Dim rs As SqlDataReader
        Dim dNewRecord As Int16 = 0
        Dim iValue As Date = Date.Now
        Dim ivalueLast30 As Date = iValue.AddDays(-30)
        Dim bSync As Boolean = GetLastTimeStamp("TransferOrder", iValue, dNewRecord)
        Dim arrList As New ArrayList
        If dNewRecord = 0 Then
            'dtr = ReadNavRecord("SELECT * FROM OWTR where (UpdateDate >= " & SafeSQL(Format(ivalueLast30, "yyyyMMdd")))
            dtr = ReadNavRecord("Select * from OWTR where (UpdateDate >= " & SafeSQL(Format(ivalueLast30, "yyyyMMdd")) & " or UpdateDate is null) order by UpdateDate")
        Else
            dtr = ReadNavRecord("Select * from OWTR where (UpdateDate >= " & SafeSQL(Format(iValue, "yyyyMMdd")) & " or UpdateDate is null) order by UpdateDate")
        End If
        While dtr.Read
            Dim dDisAmt As Double = 0
            rs = ReadRecord("Select ordno from transferhdr where Transno= " & SafeSQL(dtr("DocNum")))
            If rs.Read = True Then
                rs.Close()
                ExecuteSQL("Update transferhdr set Transno = " & dtr("DocNum") & ", custid = " & SafeSQL(dtr("cardcode")) & ", DTG = " & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")) & " where ordno =" & SafeSQL(dtr("DocNum")))
            Else
                Dim atrans As New ArrTrans
                atrans.transno = dtr("DocNum").ToString
                atrans.DocEntry = dtr("DocEntry").ToString
                arrList.Add(atrans)
                rs.Close()
                ' ExecuteSQL("Insert into Invoice(InvNo, InvDt, DONo, DoDt ,OrdNo, CustId, AgentID, Discount, SubTotal, GSTAmt, TotalAmt, PaidAmt, Void, PrintNo, PayTerms, CurCode, CurExRate, PONo, Exported,DTG,CompanyName,AcBillRef) Values (" & SafeSQL(dtr("DocNum").ToString) & "," & SafeSQL(Format(dtr("DocDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("DocNum").ToString) & "," & SafeSQL(Format(dtr("DocDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("Ref1").ToString) & "," & SafeSQL(dtr("CardCode").ToString) & "," & SafeSQL(dtr("SlpCode").ToString) & "," & dDisAmt & ", " & dtr("DocTotal") - dtr("VatSum") & ", " & dtr("VatSum") & " , " & dtr("DocTotal") & ", " & dtr("PaidToDate") & ",0,1," & SafeSQL(dtr("GroupNum").ToString) & "," & SafeSQL(dtr("DocCur").ToString) & "," & IIf(IsDBNull(dtr("DocRate")), 1, dtr("DocRate")) & ",'',1," & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")) & "," & SafeSQL("STD") & "," & SafeSQL(dtr("CardCode").ToString) & ")")
                'ExecuteSQL("Insert into OrderHdr (OrdNo, OrdDt, CustId, PONo, AgentId, Discount, SubTotal, GSTAmt, TotalAmt, PayTerms, CurCode, CurExRate,Delivered,void,exported,DeliveryDate, DTG, Remarks) Values (" & SafeSQL(dtr("DocNum").ToString) & "," & SafeSQL(Format(dtr("DocDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("cardcode").ToString) & "," & SafeSQL(dtr("DocNum").ToString) & "," & SafeSQL(dtr("Ref1").ToString) & "," & SafeSQL(0) & "," & SafeSQL(dtr("SlpCode").ToString) & "," & SafeSQL(dtr("SlpCode").ToString) & "," & SafeSQL(dtr("SlpCode").ToString) & "," & SafeSQL(dtr("SlpCode").ToString) & ", 0, " & "1,0,0,0" & "," & SafeSQL(Format(Date.Now, "yyyyMMdd 00:00:00")).ToString & "," & SafeSQL(Format(Date.Now, "yyyyMMdd 00:00:00")).ToString & "," & SafeSQL(dtr("CardCode").ToString) & ")")
                ExecuteSQL("Insert into TransferHdr(TransNo, TransDt, FromLoc, ToLoc, TransStatus, InTransitCode, DTG, Exported, RGNo, ReceiptDate, ShipmentDate, IsNavisionTO) Values (" & SafeSQL(dtr("DocNum").ToString) & "," & SafeSQL(Format(dtr("DocDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL("Transfer-from Code").ToString & "," & SafeSQL("Transfer-to Code").ToString & ",1, " & SafeSQL("In-Transit Code").ToString & "," & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")) & ", 1," & SafeSQL("") & "," & SafeSQL(Format(dtr("DocDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(Format(dtr("DocDueDate"), "yyyyMMdd HH:mm:ss")) & ",1)")
            End If
        End While
        dtr.Close()
        dtr.Dispose()
        dtr = Nothing

        If rs Is Nothing = False Then
            rs.Dispose()
            rs = Nothing
        End If

        Importtransferdet(arrList)
        ' ExecuteSQLAnother("Update InvItem set UOM=Item.BaseUOM from InvItem, Item where InvItem.ItemNo= Item.ItemNo and (InvItem.UOM is Null or InvItem.UOM='')")

        If bSync = True Then
            UpdateLastTimeStamp("TransferOrder", iValue)
        Else
            InsertLastTimeStamp("TransferOrder", iValue)
        End If
    End Sub
    Public Sub ImportSalesorder()
        Dim dtr As SqlDataReader
        Dim rs As SqlDataReader
        Dim dNewRecord As Int16 = 0
        Dim iValue As Date = Date.Now
        Dim ivalueLast30 As Date = iValue.AddDays(-30)
        Dim bSync As Boolean = GetLastTimeStamp("Salesorder", iValue, dNewRecord)
        Dim arrList As New ArrayList
        If dNewRecord = 0 Then

            dtr = ReadNavRecord("Select * from ORDR where (UpdateDate >= " & SafeSQL(Format(ivalueLast30, "yyyyMMdd")) & " or UpdateDate is null) order by UpdateDate")
        Else
            dtr = ReadNavRecord("Select * from ORDR where (UpdateDate >= " & SafeSQL(Format(iValue, "yyyyMMdd")) & " or UpdateDate is null) order by UpdateDate")
        End If
        While dtr.Read
            Dim dDisAmt As Double = 0
            rs = ReadRecord("Select ordno from orderhdr where ordno= " & SafeSQL(dtr("DocNum")))
            If rs.Read = True Then
                rs.Close()
                ExecuteSQL("Update orderhdr set ordno = " & dtr("DocNum") & ", custid = " & SafeSQL(dtr("cardcode")) & ", DTG = " & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")) & " where ordno =" & SafeSQL(dtr("DocNum")))
            Else
                Dim aord As New Arrord
                aord.Ordno = dtr("DocNum").ToString
                aord.DocEntry = dtr("DocEntry").ToString
                arrList.Add(aord)
                rs.Close()
                ExecuteSQL("Insert into OrderHdr (OrdNo, OrdDt, CustId, PONo, AgentId, Discount, SubTotal, GSTAmt, TotalAmt, PayTerms, CurCode, CurExRate,Delivered,void,exported,DeliveryDate, DTG, Remarks) Values (" & SafeSQL(dtr("DocNum").ToString) & "," & SafeSQL(Format(dtr("DocDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("cardcode").ToString) & "," & SafeSQL(dtr("DocNum").ToString) & "," & SafeSQL(dtr("Ref1").ToString) & "," & SafeSQL(0) & "," & SafeSQL(dtr("SlpCode").ToString) & "," & SafeSQL(dtr("SlpCode").ToString) & "," & SafeSQL(dtr("SlpCode").ToString) & "," & SafeSQL(dtr("SlpCode").ToString) & ", 0, " & "1,0,0,0" & "," & SafeSQL(Format(Date.Now, "yyyyMMdd 00:00:00")).ToString & "," & SafeSQL(Format(Date.Now, "yyyyMMdd 00:00:00")).ToString & "," & SafeSQL(dtr("CardCode").ToString) & ")")
            End If
        End While
        dtr.Close()
        dtr.Dispose()
        dtr = Nothing

        If rs Is Nothing = False Then
            rs.Dispose()
            rs = Nothing
        End If

        ImportordItem(arrList)


        If bSync = True Then
            UpdateLastTimeStamp("Salesorder", iValue)
        Else
            InsertLastTimeStamp("Salesorder", iValue)
        End If
    End Sub

    Public Sub Importtransferdet(ByVal arrList As ArrayList)

        Dim dtr As SqlDataReader
        For iIndex As Integer = 0 To arrList.Count - 1
            Dim aPo As ArrTrans
            aPo = arrList(iIndex)
            ExecuteSQL("Delete from transferdet where transno= " & SafeSQL(aPo.transno))

            dtr = ReadNavRecord("Select * from WTR1 where DocEntry = " & SafeSQL(aPo.DocEntry) & " order by LineNum")
            While dtr.Read

                If dtr("DocEntry").ToString = aPo.DocEntry Then
                    ExecuteSQL("Insert into TransferDet(TransNo, [LineNo], [AttachedToLineNo], ItemNo, Description, UOM, Qty, UnitPrice, ShippedQty, ReceivedQty, ShipmentDate, ReceivedDate, VariantCode) Values (" & SafeSQL(aPo.transno.ToString) & ", " & dtr("LineNum").ToString & ", " & 1 & "," & SafeSQL(dtr("ItemCode").ToString) & "," & SafeSQL(dtr("Dscription").ToString) & ", " & SafeSQL(dtr("UomEntry").ToString) & ", " & SafeSQL(dtr("Quantity").ToString) & ", " & 0 & ", " & SafeSQL(dtr("Quantity").ToString) & ", " & SafeSQL(dtr("Quantity").ToString) & ", " & SafeSQL("0") & ", " & SafeSQL("0") & ", " & SafeSQL(dtr("slpcode").ToString) & ")")
                End If
            End While
            dtr.Close()
        Next
        If dtr Is Nothing = False Then
            dtr.Dispose()
            dtr = Nothing
        End If
        ExecuteSQLAnother("Update orditem set UOM=Item.BaseUOM from orditem, Item where orditem.ItemNo= Item.ItemNo")
    End Sub



    Public Sub ImportordItem(ByVal arrList As ArrayList)

        Dim dtr As SqlDataReader
        For iIndex As Integer = 0 To arrList.Count - 1
            Dim aPo As Arrord
            aPo = arrList(iIndex)
            ExecuteSQL("Delete from orditem where ordno= " & SafeSQL(aPo.Ordno))

            dtr = ReadNavRecord("Select * from RDR1 where DocEntry = " & SafeSQL(aPo.DocEntry) & " order by LineNum")
            While dtr.Read
                'MsgBox(dtr("DocEntry"))
                '  MsgBox(aPo.PONo)

                If dtr("DocEntry").ToString = aPo.DocEntry Then
                    ExecuteSQL("Insert into OrdItem (OrdNo, [LineNo],VariantCode,  ItemNo, Uom, Qty,Price, DisPr, DisPer, Discount, SubAmt, GSTAmt, DeliQty,FOC, Location, Description,DeliveryDate) Values (" & SafeSQL(aPo.Ordno.ToString) & ", " & SafeSQL(dtr("Linenum").ToString) & ", " & SafeSQL(dtr("Flags").ToString) & "," & SafeSQL(dtr("ItemCode").ToString) & ", " & SafeSQL(dtr("UomEntry").ToString) & ", " & SafeSQL(dtr("Quantity").ToString) & ", 0, " & SafeSQL(dtr("DiscPrcnt").ToString) & ", " & SafeSQL(dtr("DiscPrcnt").ToString) & ", 0, " & SafeSQL("0") & ", " & SafeSQL("0").ToString & ", " & SafeSQL(dtr("DelivrdQty").ToString) & ", " & "1" & "," & SafeSQL(dtr("WhsCode").ToString) & "," & SafeSQL(dtr("Dscription").ToString) & "," & SafeSQL("") & ")")
                    ' ExecuteSQL("Insert into OrdItem (OrdNo, [LineNo],  ItemNo,VariantCode, Description, Uom, Qty,Foc, Price, DisPr, DisPer, Discount, SubAmt, GSTAmt, DeliQty, Location, DeliveryDate, PromoId, PromoOffer, Priority) Values (" & SafeSQL(aPo.PONo.ToString) & ", " & SafeSQL(dtr("Linenum").ToString) & ", " & SafeSQL(dtr("Flags").ToString) & "," & SafeSQL(dtr("ItemCode").ToString) & ", " & SafeSQL(dtr("UomEntry").ToString) & ", " & SafeSQL(dtr("Quantity").ToString) & ", 0, " & SafeSQL(dtr("DiscPrcnt").ToString) & ", " & SafeSQL(dtr("DiscPrcnt").ToString) & ", 0, " & SafeSQL(dtr("DiscPrcnt")) & ", " & SafeSQL(dtr("LineTotal").ToString) & ", " & SafeSQL(dtr("TotalFrgn")) & ", " & SafeSQL(dtr("DelivrdQty").ToString) & ", " & "1" & "," & SafeSQL(dtr("OCRCode").ToString) & "," & SafeSQL(dtr("Dscription").ToString) & "," & SafeSQL("") & "," & SafeSQL(dtr("SlpCode").ToString) & "," & SafeSQL(dtr("LocCode").ToString) & ",''," & SafeSQL("") & ")")
                    ' ExecuteSQL("Insert into InvItem(InvNo, [LineNo], ItemNo, UOM, Qty, Foc, Price, DisPer, DisPr, Discount, SubAmt, GstAmt, DeliQty, BaseUOM, BaseQty, Description, ColorRemarks) Values (" & SafeSQL(dtr("Ref1").ToString) & ", " & dtr("DocLineNum").ToString & ", " & SafeSQL(dtr("ItemCode").ToString) & ", " & SafeSQL("") & ", " & dtr("OutQty").ToString & ", 0, " & dtr("Price").ToString & ", 0, 0, 0, " & dtr("OutQty").ToString * dtr("Price").ToString & ", 0, " & dtr("OutQty").ToString & ", " & SafeSQL("") & ", 1," & SafeSQL(dtr("Dscription").ToString) & ",'')")
                    'Exit For
                End If
            End While
            dtr.Close()
        Next
        If dtr Is Nothing = False Then
            dtr.Dispose()
            dtr = Nothing
        End If
        ExecuteSQLAnother("Update orditem set UOM=Item.BaseUOM from orditem, Item where orditem.ItemNo= Item.ItemNo")
    End Sub

    Public Sub ImportPurchasecreditnote()
        Dim dNewRecord As Int16 = 0
        Dim iValue As Date = Date.Now
        Dim ivalueLast30 As Date = iValue.AddDays(-200)
        Dim bSync As Boolean = GetLastTimeStamp("Purchasecreditnote", iValue, dNewRecord)
        Dim dtr As SqlDataReader
        Dim rs As SqlDataReader
        Dim arrList As New ArrayList
        If dNewRecord = 0 Then
            dtr = ReadNavRecord("Select * from ORPC where DocTotal > PaidToDate and (UpdateDate >= " & SafeSQL(Format(ivalueLast30, "yyyyMMdd")) & " or UpdateDate is null) order by UpdateDate")
        Else
            dtr = ReadNavRecord("Select * from ORPC where DocTotal > PaidToDate and (UpdateDate >= " & SafeSQL(Format(ivalueLast30, "yyyyMMdd")) & " or UpdateDate is null) order by UpdateDate")
        End If
        While dtr.Read
            Dim dDisAmt As Double = 0
            rs = ReadRecord("Select creditnoteNo from POCreditNote where CreditNoteNo= " & SafeSQL(dtr("DocNum")))
            ' ExecuteSQL("Delete POCreditNote where CreditNoteNo= " & SafeSQL(dtr("DocNum")))
            Dim apCr As New ArrPrCrNote
            apCr.CreditNoteNo = dtr("DocNum").ToString
            apCr.DocEntry = dtr("DocEntry").ToString
            'apCr.DeliQty = ""
            'apCr.FromBin = dtr("CardCode").ToString
            'apCr.FromLocation = dtr("SlpCode").ToString
            arrList.Add(apCr)
            rs.Close()
            'ExecuteSQL("Insert into CreditNote(CreditNoteNo, CreditDate, CustNo, GoodsReturnNo, SalesPersonCode, Discount, SubTotal, GST, TotalAmt, PaidAmt, Void, Exported,DTG) Values (" & SafeSQL(dtr("No_").ToString) & "," & SafeSQL(Format(aDocDate, "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("Sell-to Customer No_").ToString) & ",''," & SafeSQL(dtr("Salesperson Code").ToString) & "," & dDisAmt & "," & (dtr("Amount")) & "," & CStr(((-1 * dtr("Original Amt_ (LCY)")) - (dtr("Amount")))) & "," & -1 * dtr("Original Amt_ (LCY)") & "," & CStr(-1 * ((-1 * dtr("Original Amt_ (LCY)")) - (-1 * dtr("Remaining Amount")))) & ",0,1," & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")) & ")")
            'ExecuteSQL("Insert into CreditNote(CreditNoteNo, CreditDate, CustNo, GoodsReturnNo, SalesPersonCode, Discount, SubTotal, GST, TotalAmt, PaidAmt, Void, Exported,DTG, CompanyName) Values (" & SafeSQL(dtr("DocNum").ToString) & "," & SafeSQL(Format(dtr("DocDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("CardCode").ToString) & ",''," & SafeSQL(dtr("SlpCode").ToString) & "," & dDisAmt & "," & dtr("DocTotal") - dtr("VatSum") & "," & dtr("VatSum") & "," & dtr("DocTotal") & "," & dtr("PaidToDate") & ",0,1," & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")) & "," & SafeSQL("STD") & ")")
            '            ExecuteSQL("Insert into POCreditNote(CreditNoteNo, CreditDate, CustNo, SalesPersonCode, Discount, SubTotal, GST, TotalAmt, Void, Exported, LocationCode, DeliveryDate, CompanyName) Values (" & SafeSQL(dtr("DocNum").ToString) & "," & SafeSQL(dtr("DocDate")) & "," & SafeSQL(dtr("CardCode").ToString) & "," & SafeSQL("") & ", 0, 0, 0, 0,1,1," & SafeSQL("") & "," & SafeSQL(dtr("PaidToDate")) & "," & SafeSQL("") & ")")
            ExecuteSQL("Insert into POCreditNote(CreditNoteNo, CreditDate, CustNo, SalesPersonCode, Discount, SubTotal, GST, TotalAmt, Void, Exported, LocationCode,  CompanyName) Values (" & SafeSQL(dtr("DocNum").ToString) & "," & SafeSQL(dtr("DocDate")) & "," & SafeSQL(dtr("CardCode").ToString) & "," & SafeSQL("") & ", 0, 0, 0, 0,1,1," & SafeSQL("") & "," & SafeSQL("") & ")")
            'End If
        End While
        dtr.Close()
        dtr.Dispose()
        dtr = Nothing

        If rs Is Nothing = False Then
            rs.Dispose()
            rs = Nothing
        End If
        ImportPOCrNoteItem(arrList)

        'For iIndex As Integer = 0 To arrList.Count - 1
        '    Dim aP As New ArrCrNote
        '    aP = arrList(iIndex)
        '    dtr = ReadNavRecord("Select sum(Amount) as Amount, sum(""Amount Including VAT"") as TotalAmt,sum(""Amount Including VAT"" - Amount) as Gst  from ""Sales Cr_Memo Line"" where ""Document No_"" = " & SafeSQL(aP.CrNo) & " ")
        '    'rs = ReadRecord("Select SUM(SubAmt) as Amount, sum(gstamt) as Gst from CreditNoteDet where CreditNoteNo= " & SafeSQL(aP.CrNo))
        '    If dtr.Read = True Then
        '        ExecuteSQLAnother("Update CreditNote set SubTotal = " & IIf(IsDBNull(dtr("Amount")), 0, dtr("Amount")) & ", " & _
        '                          "GST = " & IIf(IsDBNull(dtr("Gst")), 0, dtr("Gst")) & ", TotalAmt = " & IIf(IsDBNull(dtr("TotalAmt")), 0, dtr("TotalAmt")) & " where CreditNoteNo =" & SafeSQL(aP.CrNo))
        '    End If
        '    dtr.Close()
        'Next
        If bSync = True Then
            UpdateLastTimeStamp("Purchasecreditnote", iValue)
        Else
            InsertLastTimeStamp("Purchasecreditnote", iValue)
        End If
    End Sub



    Public Sub ImportPOCrNoteItem(ByVal arrList As ArrayList)

        Dim dtr As SqlDataReader
        For iIndex As Integer = 0 To arrList.Count - 1
            Dim aPo As ArrPrCrNote
            aPo = arrList(iIndex)
            ExecuteSQL("Delete from POCreditNoteDet where CreditNoteNo= " & SafeSQL(aPo.CreditNoteNo))

            dtr = ReadNavRecord("Select * from RPC1 where DocEntry = " & SafeSQL(aPo.DocEntry) & " order by LineNum")
            While dtr.Read
                If dtr("DocEntry").ToString = aPo.DocEntry Then
                    '  ExecuteSQL("Insert into OrdItem (OrdNo, [LineNo],VariantCode,  ItemNo, Uom, Qty,Price, DisPr, DisPer, Discount, SubAmt, GSTAmt, DeliQty,FOC, Location, Description,DeliveryDate) Values (" & SafeSQL(aPo.Ordno.ToString) & ", " & SafeSQL(dtr("Linenum").ToString) & ", " & SafeSQL(dtr("Flags").ToString) & "," & SafeSQL(dtr("ItemCode").ToString) & ", " & SafeSQL(dtr("UomEntry").ToString) & ", " & SafeSQL(dtr("Quantity").ToString) & ", 0, " & SafeSQL(dtr("DiscPrcnt").ToString) & ", " & SafeSQL(dtr("DiscPrcnt").ToString) & ", 0, " & SafeSQL("0") & ", " & SafeSQL("0").ToString & ", " & SafeSQL(dtr("DelivrdQty").ToString) & ", " & "1" & "," & SafeSQL(dtr("WhsCode").ToString) & "," & SafeSQL(dtr("Dscription").ToString) & "," & SafeSQL("") & ")")
                    ExecuteSQL("Insert into POCreditNoteDet(CreditNoteNo, [LineNo], [AttachedToLineNo], ItemNo, UOM, Qty, Price, Amt, VariantCode,Description, FromLocation, FromBin, DeliQty, Remarks) Values (" & SafeSQL(aPo.CreditNoteNo.ToString) & ", " & SafeSQL(dtr("Linenum").ToString) & ", " & SafeSQL(dtr("Flags").ToString) & "," & SafeSQL(dtr("ItemCode").ToString) & ", " & SafeSQL(dtr("UomEntry").ToString) & ", " & SafeSQL(dtr("Quantity").ToString) & ", 0, " & SafeSQL(dtr("DiscPrcnt").ToString) & "," & SafeSQL("0") & "," & SafeSQL(dtr("Dscription").ToString) & "," & SafeSQL(dtr("WhsCode").ToString) & "," & SafeSQL("") & ",0," & SafeSQL(dtr("Dscription").ToString) & ")")
                End If
            End While
            dtr.Close()
        Next
        If dtr Is Nothing = False Then
            dtr.Dispose()
            dtr = Nothing
        End If
        ExecuteSQLAnother("Update orditem set UOM=Item.BaseUOM from orditem, Item where orditem.ItemNo= Item.ItemNo")
    End Sub



    Public Function GetShipToCode(ByVal scustno As String, ByVal sShipToName As String) As String
        Dim dtr As SqlDataReader
        Dim sShipToCode As String = ""
        'System.IO.File.AppendAllText(Application.StartupPath & "\ShipToCode.txt", "Select * from CustomerBill where CustNo = " & SafeSQL(scustno) & " and Address = " & SafeSQL(sShipToName))
        dtr = ReadRecord("Select * from CustomerBill where CustNo = " & SafeSQL(scustno) & " and Address = " & SafeSQL(sShipToName))
        If dtr.Read = True Then
            sShipToCode = dtr("AcBillRef").ToString
        End If
        dtr.Close()
        Return sShipToCode
    End Function


    Public Sub ImportCreditNote()
        Dim dNewRecord As Int16 = 0
        Dim iValue As Date = Date.Now
        Dim ivalueLast30 As Date = iValue.AddDays(-30)
        Dim bSync As Boolean = GetLastTimeStamp("CreditNote", iValue, dNewRecord)
        Dim dtr As SqlDataReader
        Dim rs As SqlDataReader
        Dim arrList As New ArrayList
        If dNewRecord = 0 Then
            ' dtr = ReadNavRecord("SELECT * FROM ORIN where DocTotal > PaidToDate and (UpdateDate >= " & SafeSQL(Format(ivalueLast30, "YYYYMMdd")))
            dtr = ReadNavRecord("Select * from ORIN where DocTotal > PaidToDate and (UpdateDate >= " & SafeSQL(Format(ivalueLast30, "yyyyMMdd")) & " or UpdateDate is null) order by UpdateDate")
        Else
            dtr = ReadNavRecord("Select * from ORIN where DocTotal > PaidToDate and (UpdateDate >= " & SafeSQL(Format(iValue, "yyyyMMdd")) & " or UpdateDate is null) order by UpdateDate")
        End If
        While dtr.Read
            Dim dDisAmt As Double = 0
            rs = ReadRecord("Select PaidAmt from CreditNote where CreditNoteNo= " & SafeSQL(dtr("DocNum")))
            If rs.Read = True Then
                rs.Close()
                'ExecuteSQL("Update CreditNote set PaidAmt = " & -1 * (dtr("Original Amt_ (LCY)") - dtr("Rem_ Amt")) & " where CreditNoteNo =" & SafeSQL(dtr("No_")))
                ExecuteSQL("Update CreditNote set PaidAmt = " & dtr("PaidToDate") & " where CreditNoteNo =" & SafeSQL(dtr("DocNum")))
            Else
                Dim aCr As New ArrCrNote
                aCr.CrNo = dtr("DocNum").ToString
                aCr.CrDate = dtr("DocDate")
                aCr.OrdNo = ""
                aCr.CustID = dtr("CardCode").ToString
                aCr.AgentID = dtr("SlpCode").ToString
                aCr.Discount = dDisAmt
                aCr.CurCode = dtr("DocCur").ToString
                aCr.CurExRate = dtr("DocRate")
                aCr.GSTAmt = dtr("VatSum")
                aCr.Subtotal = dtr("DocTotal") - dtr("VatSum")
                aCr.TotalAmt = dtr("DocTotal")
                aCr.PaidAmt = dtr("PaidToDate")
                aCr.payterms = dtr("GroupNum").ToString
                arrList.Add(aCr)
                rs.Close()
                'ExecuteSQL("Insert into CreditNote(CreditNoteNo, CreditDate, CustNo, GoodsReturnNo, SalesPersonCode, Discount, SubTotal, GST, TotalAmt, PaidAmt, Void, Exported,DTG) Values (" & SafeSQL(dtr("No_").ToString) & "," & SafeSQL(Format(aDocDate, "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("Sell-to Customer No_").ToString) & ",''," & SafeSQL(dtr("Salesperson Code").ToString) & "," & dDisAmt & "," & (dtr("Amount")) & "," & CStr(((-1 * dtr("Original Amt_ (LCY)")) - (dtr("Amount")))) & "," & -1 * dtr("Original Amt_ (LCY)") & "," & CStr(-1 * ((-1 * dtr("Original Amt_ (LCY)")) - (-1 * dtr("Remaining Amount")))) & ",0,1," & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")) & ")")
                ExecuteSQL("Insert into CreditNote(CreditNoteNo, CreditDate, CustNo, GoodsReturnNo, SalesPersonCode, Discount, SubTotal, GST, TotalAmt, PaidAmt, Void, Exported,DTG, CompanyName) Values (" & SafeSQL(dtr("DocNum").ToString) & "," & SafeSQL(Format(dtr("DocDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("CardCode").ToString) & ",''," & SafeSQL(dtr("SlpCode").ToString) & "," & dDisAmt & "," & dtr("DocTotal") - dtr("VatSum") & "," & dtr("VatSum") & "," & dtr("DocTotal") & "," & dtr("PaidToDate") & ",0,1," & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")) & "," & SafeSQL("STD") & ")")
            End If
        End While
        dtr.Close()
        dtr.Dispose()
        dtr = Nothing

        If rs Is Nothing = False Then
            rs.Dispose()
            rs = Nothing
        End If
        '  ImportCrNoteItem(arrList)

        'For iIndex As Integer = 0 To arrList.Count - 1
        '    Dim aP As New ArrCrNote
        '    aP = arrList(iIndex)
        '    dtr = ReadNavRecord("Select sum(Amount) as Amount, sum(""Amount Including VAT"") as TotalAmt,sum(""Amount Including VAT"" - Amount) as Gst  from ""Sales Cr_Memo Line"" where ""Document No_"" = " & SafeSQL(aP.CrNo) & " ")
        '    'rs = ReadRecord("Select SUM(SubAmt) as Amount, sum(gstamt) as Gst from CreditNoteDet where CreditNoteNo= " & SafeSQL(aP.CrNo))
        '    If dtr.Read = True Then
        '        ExecuteSQLAnother("Update CreditNote set SubTotal = " & IIf(IsDBNull(dtr("Amount")), 0, dtr("Amount")) & ", " & _
        '                          "GST = " & IIf(IsDBNull(dtr("Gst")), 0, dtr("Gst")) & ", TotalAmt = " & IIf(IsDBNull(dtr("TotalAmt")), 0, dtr("TotalAmt")) & " where CreditNoteNo =" & SafeSQL(aP.CrNo))
        '    End If
        '    dtr.Close()
        'Next
        If bSync = True Then
            UpdateLastTimeStamp("CreditNote", iValue)
        Else
            InsertLastTimeStamp("CreditNote", iValue)
        End If
    End Sub



    Public Sub ImportCrNoteItem(ByVal arrList As ArrayList)


        '        Dim dtr As SqlDataReader
        '        For iIndex As Integer = 0 To arrList.Count - 1
        '            Dim aPr As ArrCrNote
        '            Dim CNT As Integer = 0
        '            aPr = arrList(iIndex)
        '            dtr = ReadNavRecord("Select * from ""Sales Cr_Memo Line"" where ""Document No_"" = " & SafeSQL(aPr.CrNo) & " order by ""Line No_""")
        '            While dtr.Read
        '                Try
        '                    CNT = 0
        'INSERT:
        '                    '              ExecuteSQL("Insert into InvItem(InvNo, [LineNo], ItemNo, UOM, Qty, Foc, Price, DisPer, DisPr, Discount, SubAmt, GstAmt, DeliQty, BaseUOM, BaseQty, Description) Values (" & SafeSQL(dtr("Document No_").ToString) & ", " & dtr("Line No_").ToString & ", " & SafeSQL(dtr("No_").ToString) & ", " & SafeSQL(dtr("Unit of Measure Code").ToString) & ", " & dtr("Quantity").ToString & ", 0, " & dtr("Unit Price").ToString & ", " & dtr("Line Discount %").ToString & ", 0, " & dtr("Line Discount Amount") & ", " & dtr("Amount").ToString & ", " & dtr("Amount Including GST") - dtr("Amount") & ", " & dtr("Quantity").ToString & ", " & SafeSQL(dtr("Unit of Measure Code").ToString) & ", " & dtr("Quantity (Base)").ToString & "," & SafeSQL(dtr("Description").ToString) & ")")
        '                    ExecuteSQL("Insert into CreditNoteDet(CreditNoteNo, ItemNo, UOM, BaseUOM, Price, Qty, Amt) Values (" & SafeSQL(dtr("Document No_").ToString) & ", " & SafeSQL(dtr("No_").ToString) & ", " & SafeSQL(dtr("Unit of Measure Code").ToString) & ",'', " & dtr("Unit Price").ToString & "," & dtr("Quantity").ToString & "," & dtr("Amount").ToString & ")")
        '                Catch
        '                    System.Threading.Thread.Sleep(5000)
        '                    CNT += 1
        '                    If CNT < 3 Then GoTo iNSERT
        '                    '       MsgBox(aPr.InvNo)
        '                End Try

        '            End While
        '            dtr.Close()
        '        Next

    End Sub

    Public Sub ImportInventory()
        Dim dtr As SqlDataReader
        ExecuteSQLAnother("Delete from GoodsInvn")
        dtr = ReadNavRecord("Select [WhsCode], [ItemCode], Sum([OnHand]) as Qty from OITW Where OnHand > 0 group by [WhsCode], [ItemCode]")
        While dtr.Read
            ExecuteSQLAnother("Insert into GoodsInvn(Location, ItemNo, Qty, UOM) Values (" & SafeSQL(dtr("WhsCode").ToString) & "," & SafeSQL(dtr("ItemCode").ToString) & "," & dtr("Qty") & ",'')")
        End While
        dtr.Close()
        dtr.Dispose()
        dtr = Nothing
        ExecuteSQLAnother("Update GoodsInvn set UOM=Item.BaseUOM from GoodsInvn, Item where GoodsInvn.ItemNo= Item.ItemNo")
    End Sub

    Private Function IsItemExists(ByVal sItemNo As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = ReadRecord("Select ItemNo from Item where ItemNo = " & SafeSQL(sItemNo))
        bAns = dtr.Read
        dtr.Close()
        Return bAns
    End Function
    Private Function IsItemPrExists(ByVal sItemNo As String, ByVal sPrGroup As String, ByVal dMinQty As Double, ByVal sSDate As Date, ByVal sEDate As Date) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = ReadRecordAnother("Select PriceGroup, ItemNo, MinQty from ItemPr where ItemNo = " & SafeSQL(sItemNo) & " and PriceGroup = " & SafeSQL(sPrGroup) & " and MinQty = " & SafeSQL(dMinQty) & " and FromDate = " & sSDate & " and ToDate = " & sEDate)
        If dtr.Read = True Then
            bAns = True
        Else
            bAns = False
        End If
        dtr.Close()
        Return bAns
    End Function

    Private Function IsAgentExists(ByVal sCode As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = ReadRecordAnother("Select Code from SalesAgent where Code = " & SafeSQL(sCode))
        bAns = dtr.Read
        dtr.Close()
        Return bAns
    End Function


    Private Function IsMDTExists(ByVal sCode As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = ReadRecordAnother("Select MDTNo from MDT where MDTNo = " & SafeSQL(sCode))
        bAns = dtr.Read
        dtr.Close()
        Return bAns
    End Function


    Private Function IsNoSeriesExists(ByVal sCode As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = ReadRecordAnother("Select MDTNo from NoSeries where MDTNo = " & SafeSQL(sCode))
        bAns = dtr.Read
        dtr.Close()
        Return bAns
    End Function

    Private Function IsCustomerExists(ByVal sCustNo As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = ReadRecord("Select CustNo from Customer where CustNo = " & SafeSQL(sCustNo))
        bAns = dtr.Read
        dtr.Close()
        Return bAns
    End Function
    Private Function IsStockInNoExists(ByVal sStockInNo As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = ReadRecord("Select StockInNo from StockInItem where StockInNo = " & SafeSQL(sStockInNo))
        bAns = dtr.Read
        dtr.Close()
        Return bAns
    End Function
    Private Function IsVendorExists(ByVal sVendNo As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = ReadRecord("Select VendorNo from Vendor where VendorNo = " & SafeSQL(sVendNo))
        bAns = dtr.Read
        dtr.Close()
        Return bAns
    End Function

    Private Function IsNavCustomerExists(ByVal sCustNo As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = ReadNavRecord("Select * from Customer where No_ = " & SafeSQL(sCustNo))
        bAns = dtr.Read
        dtr.Close()
        Return bAns
    End Function

    Private Function UpdateCustName(ByVal sCustNo As String) As Integer
        Dim dtr As SqlDataReader
        Dim iAns As Integer = 0
        dtr = ReadRecord("Select CustName from Customer where AcCustCode = " & SafeSQL(sCustNo))
        If dtr.Read = True Then
            If dtr("CustName") = "" Then
                iAns = 1
            End If
        End If
        dtr.Close()
        Return iAns
    End Function

    Private Function HasBranch(ByVal sCustNo As String) As Integer
        Dim dtr As SqlDataReader
        Dim iAns As Integer = 0
        dtr = ReadRecord("Select CustNo from Customer where AcCustCode = " & SafeSQL(sCustNo))
        If dtr.Read = True Then
            If dtr("CustNo") = sCustNo Then
                iAns = 1
            Else
                iAns = 2
            End If
        End If
        dtr.Close()
        Return iAns
    End Function
    Private Sub UpdateCustomer()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Distinct AcCustCode, PriceGroup, PaymentTerms from Customer")
        While dtr.Read
            If IsDBNull(dtr("AcCustCode")) = False Then
                If dtr("AcCustCode").ToString <> "" Then
                    ExecuteSQLAnother("Update Customer Set PriceGroup =" & SafeSQL(dtr("PriceGroup").ToString) & ", PaymentTerms=" & SafeSQL(dtr("PaymentTerms").ToString) & " where Accustcode=" & SafeSQL(dtr("Accustcode").ToString))
                End If
            End If
        End While
        dtr.Close()
        ExecuteSQLAnother("Update Customer Set PriceGroup ='STD' Where (PriceGroup='' or PriceGroup is Null)")
    End Sub

    Public Sub ImportInvItem(ByVal arrList As ArrayList)
        'Dim sInvNos As String = "''"
        'For idx As Integer = 0 To arrList.Count - 1
        '    Dim aPr As ArrInvoice
        '    aPr = arrList(idx)
        '    sInvNos &= "," & SafeSQL(aPr.InvNo)
        'Next

        Dim dtr As SqlDataReader

        ' ExecuteSQL("Delete from InvItem where InvNo in (" & sInvNos & ")")

        For iIndex As Integer = 0 To arrList.Count - 1
            Dim aPr As ArrInvoice
            aPr = arrList(iIndex)
            ExecuteSQL("Delete from InvItem where InvNo= " & SafeSQL(aPr.InvNo))

            dtr = ReadNavRecord("Select * from OINM where Ref1 = " & SafeSQL(aPr.InvNo) & " order by Ref1")
            While dtr.Read
                'If dtr("Document No_").ToString = aPr.InvNo Then
                ExecuteSQL("Insert into InvItem(InvNo, [LineNo], ItemNo, UOM, Qty, Foc, Price, DisPer, DisPr, Discount, SubAmt, GstAmt, DeliQty, BaseUOM, BaseQty, Description, ColorRemarks) Values (" & SafeSQL(dtr("Ref1").ToString) & ", " & dtr("DocLineNum").ToString & ", " & SafeSQL(dtr("ItemCode").ToString) & ", " & SafeSQL("") & ", " & dtr("OutQty").ToString & ", 0, " & dtr("Price").ToString & ", 0, 0, 0, " & dtr("OutQty").ToString * dtr("Price").ToString & ", 0, " & dtr("OutQty").ToString & ", " & SafeSQL("") & ", 1," & SafeSQL(dtr("Dscription").ToString) & ",'')")
                'Exit For
                'End If
            End While
            dtr.Close()
        Next
        If dtr Is Nothing = False Then
            dtr.Dispose()
            dtr = Nothing
            ExecuteSQLAnother("Update InvItem set UOM=Item.BaseUOM from InvItem, Item where InvItem.ItemNo= Item.ItemNo")
        End If
    End Sub


    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click

    End Sub


    Public Sub loadCombo()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Code, Name from SalesAgent, MDT where MDT.AgentID=SalesAgent.Code")
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


    Private Sub chkSelAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSelAll.CheckedChanged
        Dim i As Integer
        If chkSelAll.Checked = True Then
            For i = 0 To dgvStatus.Rows.Count - 1
                If dgvStatus.Rows(i).IsNewRow = True Then Exit For
                dgvStatus.Item(1, i).Value = True
            Next
        Else
            For i = 0 To dgvStatus.Rows.Count - 1
                If dgvStatus.Rows(i).IsNewRow = True Then Exit For
                dgvStatus.Item(1, i).Value = False
            Next
        End If
    End Sub

    Private Sub btnEx_Click(sender As Object, e As EventArgs) Handles btnEx.Click
        Try
            Me.Cursor = Cursors.WaitCursor
            ExportInvoices()
            ExportSalesOrder()
            ExportCreditMemo()
            ExportPayment()
            ExportStockOrder()
            ExportReturn()
            ExportCustVisit()
            ExportBank()
            ExportStockInItem()
            ExportExchange()
            'ExportException()
            'ExportItemTrans() detail table not found
            'ExportService()  to be check
            Me.Cursor = Cursors.Default
            MsgBox("Export Completed", vbInformation, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, "Warning")
        End Try

    End Sub



End Class