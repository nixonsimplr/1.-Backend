Imports System
Imports System.Resources
Imports System.Globalization
Imports System.Threading
Imports System.Reflection
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Imports System.Text.RegularExpressions
'Imports System.Data.Odbc
Imports System.Data.SqlTypes
Public Class ImportStatus
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
    Dim bDeleteIns As Boolean = False
    Public sCustPostGroup, sGenPostGroup, sGSTPostGroup, sGSTProdGroup, sGenJournalTemplate, sGenJournalBatch, sWorkSheetTemplate, sJournalBatch, sItemJnlTemplate, sItemJnlBatch, sFocBatch, sExBatch, sBadLoc, sItemReclassTemplate, sItemReclassBatch As String
    Private Structure DelCust
        Dim CustID As String
        Dim PrGroup As String
    End Structure
    Dim i, igCnt As Integer
    'Dim cnt As Integer = 0
    Private NavCompanyName As String = GetCompanyName()
    Private NavDBName As String = GetNavDBName()
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
                    Return row("Value").ToString()
                Next row
            End If
        Next table
        Return ""
    End Function
    Private Function GetNavDBName() As String
        Dim ds As New DataSet
        Dim dataDirectory As String
        dataDirectory = Windows.Forms.Application.StartupPath
        ds.ReadXml(dataDirectory & "\Simplr.xml")
        Dim table As DataTable
        For Each table In ds.Tables
            Dim row As DataRow
            If table.TableName = "NavDBName" Then
                For Each row In table.Rows
                    Return row("Value").ToString()
                Next row
            End If
        Next table
        Return ""
    End Function

    Private Sub ImportStatus_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            'btnIm_Click(Me, Nothing)
            btnEx_Click(Me, Nothing)
            Me.Close()
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Error in FormLoad - ") & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
            Me.Close()
        End Try
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

    Private Function GetLastImpDate() As Date
        Dim dDate As Date = Date.Now
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select LastImExDate from System")
        If dtr.Read Then
            dDate = dtr("LastImExDate")
            'bFound = True
        End If
        dtr.Close()
        dtr.Dispose()
        Return dDate
    End Function

    Private Sub btnIm_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnIm.Click
        Try

            Dim dImpStartDate As Date = Date.Now
            btnIm.Enabled = False
            btnEx.Enabled = False

            ConnectNavDB()
            ConnectAnotherDB()

            Dim LastImpDate As Date = GetLastImpDate()

            ExecuteSQL("Update System set LastImExDate = " & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")))

            If Format(LastImpDate, "dd") <> Format(dImpStartDate, "dd") Then
                bDeleteIns = True
                Dim sTarget As String = Application.StartupPath & "\Archieve"
                Try
                    Dim sFolder As String = sTarget & "\" & Format(DateTime.Now, "yyyyMMddHHmmss")
                    If System.IO.Directory.Exists(sTarget) = False Then
                        System.IO.Directory.CreateDirectory(sTarget)
                    End If
                    If System.IO.Directory.Exists(sFolder) = False Then
                        System.IO.Directory.CreateDirectory(sFolder)
                    End If
                    Dim dir As System.IO.DirectoryInfo = New System.IO.DirectoryInfo(Application.StartupPath)
                    Dim arrfInfo() As System.IO.FileInfo = dir.GetFiles("ErrorLog.txt")
                    If Not arrfInfo Is Nothing Then
                        For Each f As System.IO.FileInfo In arrfInfo
                            f.MoveTo(sFolder & "\" & f.Name)
                        Next
                    End If
                    System.IO.File.Delete(Application.StartupPath & "\ErrorLog.txt")
                Catch ex As Exception
                End Try
            End If

            If bDeleteIns = True Then

                Try
                    ExecuteSQLAnother("update customer set SalesIndicator = 0 ")
                    ExecuteSQLAnother("update customer set SalesIndicator = 1 where custno in (select distinct custid from invoice where subtotal>3000 and invdt>getdate()-365 and isnull(mdtno,'') <> '')")

                    ExecuteSQLAnother("Delete from ErrorLog where DTG < GetDate()-30")

                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'ErrorLog Clear - Finished'," & SafeSQL(NavCompanyName) & ",''" & ")")

                Catch ex As Exception
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'ErrorLog Clear - Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try


            End If
            ImportCustomer()
            ImportMapCustomer()
            ImportProduct()
            ImportItemPrice()
            ImportPayterms()
            ImportPriceGroup()
            ImportLocation()
            ImportProductTeam()

            ImportArea()
            ImportBank()
            'ImportCustAgent()
            ImportSalesAgent()
            ImportUserBrand()


            ImportShopType()
            ImportBusinessType()
            ImportLocationType()
            ImportProductList()
            ImportInstitutionType()
            ImportProvince()
            ImportBarangay()
            ImportTerritory()

            If bDeleteIns = True Then
                ImportTeam()
            End If
            ImportErpStockRequest()
            ImportInvoice()
            ImportCreditNote()

            DisConnect()
            'MsgBox("Import completed", vbInformation, "Information")
        Catch ex As Exception
            btnIm.Enabled = True
            btnEx.Enabled = True

            ' MsgBox(ex.Message)
            System.IO.File.AppendAllText(Application.StartupPath & "\ErrorLog.txt", Date.Now & vbCrLf & ex.Message & vbCrLf)
            DisConnect()
        End Try

    End Sub

    Public Sub DisConnect()
        DisconnectNavDB()
        DisconnectAnotherDB()
        btnIm.Enabled = True
        btnEx.Enabled = True
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
        dtr = ReadRecordAnother(strSql)
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
        Dim custtype As String = String.Empty

        Try
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Customer - Insert Start'," & SafeSQL(NavCompanyName) & ",''" & ")")
            dtr = ReadNavRecord("Select * from Customer where IsNull(IsRead,0)= 0 ")
            While dtr.Read

                ''--------
                'Dim SimplrNewCustNo As String = ""
                'If IsDBNull(dtr("SimplrNewCustNo")) = False Or dtr("SimplrNewCustNo").ToString() <> "" Then

                '    If IsCustomerExistsInSalesOrder(dtr("CustNo").ToString) Then
                '        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Customer - Update Customer In Sales Order'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("CustNo").ToString + " - " + dtr("SimplrNewCustNo").ToString) & ")")
                '        sQry = " Update Orderhdr set CustId= " & SafeSQL(dtr("SimplrNewCustNo").ToString) &
                '               " WHERE CustId = " & SafeSQL(dtr("CustNo").ToString)
                '        ExecuteSQL(sQry)
                '    End If

                '    If IsCustomerExistsInInvoice(dtr("CustNo").ToString) Then
                '        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Customer - Update Customer In Invoice'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("CustNo").ToString + " - " + dtr("SimplrNewCustNo").ToString) & ")")
                '        sQry = " Update Invoice set CustId= " & SafeSQL(dtr("SimplrNewCustNo").ToString) &
                '               " WHERE CustId = " & SafeSQL(dtr("CustNo").ToString)
                '        ExecuteSQL(sQry)
                '    End If

                '    If IsCustomerExistsInPayment(dtr("CustNo").ToString) Then
                '        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Customer - Update Customer In Receipt'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("CustNo").ToString + " - " + dtr("SimplrNewCustNo").ToString) & ")")
                '        sQry = " Update Receipt set CustId= " & SafeSQL(dtr("SimplrNewCustNo").ToString) &
                '               " WHERE CustId = " & SafeSQL(dtr("CustNo").ToString)
                '        ExecuteSQL(sQry)
                '    End If

                '    SimplrNewCustNo = SafeSQL(dtr("SimplrNewCustNo").ToString)
                '    ExecuteSQLAnother("UPDATE NewCust SET active=0 WHERE CustID = " & SimplrNewCustNo)
                'Else
                '    ExecuteSQLAnother("UPDATE NewCust SET active=0 WHERE CustID = " & SafeSQL(dtr("CustNo").ToString))
                'End If

                ''ExecuteSQLAnother("UPDATE NewCust SET active=0 WHERE CustID = CASE WHEN " & SimplrNewCustNo & " IS NOT NULL THEN " & SimplrNewCustNo & " ELSE " & SafeSQL(dtr("CustNo").ToString) & " END ")
                ''--------

                If dtr("CustNo").ToString <> "" Then
                    If IsCustomerExists(dtr("CustNo").ToString) Then
                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Customer - Update Customer'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("CustNo").ToString) & ")")
                        sQry = "UPDATE [dbo].[Customer]" &
                               " SET [CustName] = " & SafeSQL(dtr("CustName").ToString) &
                                 " ,[ChineseName] = " & SafeSQL(dtr("CustName").ToString) &
                                  ",[SearchName] = " & SafeSQL(dtr("CustName").ToString) &
                                  ",[Address] = " & SafeSQL(dtr("Address").ToString) &
                                  ",[Address2] = " & SafeSQL(dtr("House_Number").ToString) &
                                  ",[Address3] = " & SafeSQL(dtr("Alley").ToString) &
                                  ",[Address4] = " & SafeSQL(dtr("Amphoe").ToString) &
                                  ",[Tambon] = " & SafeSQL(dtr("Tambon").ToString) &
                                  ",[Province] = " & SafeSQL(dtr("Province").ToString) &
                                  ",[PostCode] = " & SafeSQL(dtr("PostCode").ToString) &
                                  ",[CountryCode] = " & SafeSQL(dtr("CountryCode").ToString) &
                                  ",[Phone] = " & SafeSQL(dtr("Phone").ToString) &
                                  ",[ContactPerson] = " & SafeSQL(dtr("ContactPerson").ToString) &
                                  ",[Balance] = " & SafeSQL(dtr("Balance").ToString) &
                                  ",[CreditLimit] = " & SafeSQL(dtr("CreditLimit").ToString) &
                                  ",[ZoneCode] = " & SafeSQL(dtr("Area").ToString) &
                                  ",[FaxNo] = " & SafeSQL(dtr("FaxNo").ToString) &
                                  ",[PaymentTerms] = " & SafeSQL(dtr("PaymentTerms").ToString) &
                                  ",[ShipAgent] = " & SafeSQL(dtr("ShipAgent").ToString) &
                                  ",[Bill-toNo] = " & SafeSQL(dtr("Bill-toNo").ToString) &
                                  ",[Active] = " & SafeSQL(dtr("Active").ToString) &
                                  ",[ShipName] = " & SafeSQL(dtr("ShipName").ToString) &
                                  ",[ShipAddr] = " & SafeSQL(dtr("DeliveryAddress").ToString) &
                                  ",[ShipAddr2] = " & SafeSQL(dtr("Ship_House_Number").ToString) &
                                  ",[ShipAddr3] = " & SafeSQL(dtr("Ship_Alley").ToString) &
                                  ",[ShipAddr4] = " & SafeSQL(dtr("Ship_Amphoe").ToString) &
                                  ",[ShipCountryCode] = " & SafeSQL(dtr("Ship_Tambon").ToString) &
                                  ",[ShipCity] = " & SafeSQL(dtr("Ship_Province").ToString) &
                                  ",[ShipPost] = " & SafeSQL(dtr("ShipPostCode").ToString) &
                                  ",[GSTType] = " & SafeSQL(dtr("GSTType").ToString) &
                                  ",[Remarks] = " & SafeSQL(dtr("Remarks").ToString) &
                                  ",[Dimension1] = " & SafeSQL(dtr("Reference").ToString) &
                                  ",[DiscountGroup] = " & SafeSQL(dtr("DiscountGroup").ToString) &
                                  ",[GSTNO] = " & SafeSQL(dtr("GSTNO").ToString) &
                                  ",[Channel] = " & SafeSQL(dtr("Branch").ToString) &
                                  ",[Bussiness_Type] = " & SafeSQL(dtr("Buss_Type").ToString) &
                                  ",[Dimension2] = " & SafeSQL(dtr("Buss_TypeOther").ToString) &
                                  ",[Remarks2] = " & SafeSQL(dtr("Credit_Time").ToString) &
                                  ",[CustGrade] = " & SafeSQL(dtr("Institution_Type").ToString) &
                                  ",[CompanyName] = " & SafeSQL(dtr("Institution_Other").ToString) &
                                  ",[Location] = " & SafeSQL(dtr("Location").ToString) &
                                  ",[Location_Other] = " & SafeSQL(dtr("Location_Other").ToString) &
                                  ",[GSTProdGroup] = " & SafeSQL(dtr("Product_List").ToString) &
                                  ",[Shop_Type] = " & SafeSQL(dtr("Shop_Type").ToString) &
                                  ",[Tax_Id] = " & SafeSQL(dtr("Tax_Id").ToString) &
                                  ",[SalesAgent] = " & SafeSQL(dtr("Area").ToString) &
                                  ",[Area] = " & SafeSQL(dtr("Area").ToString) &
                                  ",[GSTCustGroup] = 'VAT7'" &
                                  ",[PaymentMethod] = 'Cash' " &
                                  ",[Email] = " & SafeSQL(dtr("Email").ToString) &
                                  ",[SOI] = " & SafeSQL(dtr("SOI").ToString) &
                                  ",[Road] = " & SafeSQL(dtr("Road").ToString) &
                                  ",[District] = " & SafeSQL(dtr("District").ToString) &
                                  ",[SubDistrict] = " & SafeSQL(dtr("SubDistrict").ToString) &
                                  ",[LocationType] = " & SafeSQL(dtr("LocationType").ToString) &
                                  ",[StoreType] = " & SafeSQL(dtr("StoreType").ToString) &
                                  " WHERE CustNo = " & SafeSQL(dtr("CustNo").ToString)

                        'ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Import Customer - " & dtr("CustNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")
                        Try
                            ExecuteSQL(sQry)
                            'UpdateGPSCoordinates(dtr("CustNo").ToString)
                        Catch ex As Exception
                            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Customer - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                        End Try

                    Else
                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Customer - Insert Customer'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("CustNo").ToString) & ")")

                        sQry = "Insert into Customer ([CustNo],	[CustName],	[ChineseName],	[SearchName],	[Address],	Address2,	Address3,	Address4,	[Tambon],	[Province],	[PostCode],                    [CountryCode],	[Phone],	[ContactPerson],	[Balance],	[CreditLimit],	[ZoneCode],	[FaxNo],	[PaymentTerms],	[ShipAgent],	[Bill-toNo],[Active],	[ShipName],	ShipAddr,	ShipAddr2,	ShipAddr3,	ShipAddr4,	ShipCountryCode,	ShipCity,	[ShipPost],	[GSTType],[Remarks],	Dimension1,	[DiscountGroup],	[GSTNO],	Channel,	Bussiness_Type,	Dimension2,	Remarks2,	CustGrade,	CompanyName, [Location],	[Location_Other],	GSTProdGroup,	[Shop_Type],	[Tax_Id],	[SalesAgent], [Area], PaymentMethod, GSTCustGroup, [Email], [SOI], [Road], [District], [SubDistrict], [LocationType], [StoreType] )" &
                                "Values (" & SafeSQL(dtr("CustNo").ToString) & "," & SafeSQL(dtr("CustName").ToString) & "," & SafeSQL(dtr("CustName").ToString) & "," & SafeSQL(dtr("CustName").ToString) & "," & SafeSQL(dtr("Address").ToString) & "," & SafeSQL(dtr("House_Number").ToString) & "," & SafeSQL(dtr("Alley").ToString) & "," & SafeSQL(dtr("Amphoe").ToString) & "," & SafeSQL(dtr("Tambon").ToString) & "," & SafeSQL(dtr("Province").ToString) & "," & SafeSQL(dtr("PostCode").ToString) & "," & SafeSQL(dtr("CountryCode").ToString) & "," & SafeSQL(dtr("Phone").ToString) & "," & SafeSQL(dtr("ContactPerson").ToString) & "," & SafeSQL(dtr("Balance").ToString) & "," & SafeSQL(dtr("CreditLimit").ToString) & "," & SafeSQL(dtr("Area").ToString) & "," & SafeSQL(dtr("FaxNo").ToString) & "," & SafeSQL(dtr("PaymentTerms").ToString) & "," & SafeSQL(dtr("ShipAgent").ToString) & "," & SafeSQL(dtr("Bill-toNo").ToString) & "," & SafeSQL(dtr("Active").ToString) & "," & SafeSQL(dtr("ShipName").ToString) & "," & SafeSQL(dtr("DeliveryAddress").ToString) & "," & SafeSQL(dtr("Ship_House_Number").ToString) & "," & SafeSQL(dtr("Ship_Alley").ToString) & "," & SafeSQL(dtr("Ship_Amphoe").ToString) & "," & SafeSQL(dtr("Ship_Tambon").ToString) & "," & SafeSQL(dtr("Ship_Province").ToString) & "," & SafeSQL(dtr("ShipPostCode").ToString) & "," & SafeSQL(dtr("GSTType").ToString) & "," & SafeSQL(dtr("Remarks").ToString) & "," & SafeSQL(dtr("Reference").ToString) & "," & SafeSQL(dtr("DiscountGroup").ToString) & "," & SafeSQL(dtr("GSTNO").ToString) & "," & SafeSQL(dtr("Branch").ToString) & "," & SafeSQL(dtr("Buss_Type").ToString) & "," & SafeSQL(dtr("Buss_TypeOther").ToString) & "," & SafeSQL(dtr("Credit_Time").ToString) & "," & SafeSQL(dtr("Institution_Type").ToString) & "," & SafeSQL(dtr("Institution_Other").ToString) & "," & SafeSQL(dtr("Location").ToString) & "," & SafeSQL(dtr("Location_Other").ToString) & "," & SafeSQL(dtr("Product_List").ToString) & "," & SafeSQL(dtr("Shop_Type").ToString) & "," & SafeSQL(dtr("Tax_Id").ToString) & "," & SafeSQL(dtr("Area").ToString) & "," & SafeSQL(dtr("Area").ToString) & ",'Cash','VAT7'," & SafeSQL(dtr("Email").ToString) & "," & SafeSQL(dtr("SOI").ToString) & "," & SafeSQL(dtr("Road").ToString) & "," & SafeSQL(dtr("District").ToString) & "," & SafeSQL(dtr("SubDistrict").ToString) & "," & SafeSQL(dtr("LocationType").ToString) & "," & SafeSQL(dtr("StoreType").ToString) & ")"

                        'ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Import Customer - " & dtr("CustNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")
                        Try
                            ExecuteSQL(sQry)
                            'UpdateGPSCoordinates(dtr("CustNo").ToString)
                        Catch ex As Exception
                            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Customer - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                        End Try
                    End If
                    ExecuteNavAnotherSQL("Update Customer set IsRead = 1  where CustNo = " & SafeSQL(dtr("CustNo").ToString))
                    'ExecuteSQLAnother("Update NewCust set Active = 0 where CustID = " & SafeSQL(dtr("SimplrNewCustNo").ToString))
                End If
            End While
            dtr.Close()
            ExecuteSQLAnother("Update Customer Set PriceGroup ='STANDARD' Where ISNULL(PriceGroup,'') = ''")
            ExecuteSQLAnother("Update Customer set CustType = 'CREDIT'  where PaymentTerms in (Select code from Payterms where duedatecalc <> '0D')")
            ExecuteSQLAnother("Update Customer set CustType = 'CASH'  where PaymentTerms in (Select code from Payterms where duedatecalc = '0D')")


            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Customer - Insert Finish'," & SafeSQL(NavCompanyName) & ",''" & ")")
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Customer - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
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
        Dim myarray As New ArrayList
        Try
            'ExecuteSQL("Update Customer set Active = 0")
            dtr = ReadNavRecord("Select Distinct StockInNo,TransDate,AgentId from StockRequestErp where IsNull(IsRead,0) = 0 Order by StockInNo")

            While dtr.Read
                If dtr("StockInNo").ToString <> "" Then
                    If IsDOExists(dtr("StockInNo").ToString) = True Then

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Import ErpStockReq Insert Error : " & dtr("StockInNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL("Record Already Exits") & ")")
                    Else
                        sQry = "Insert into DeliveryOrderHdr (OrdNo, OrdDt,AgentID,CustId,PickingNo,InvNo,SoNo,Discount,SubTotal,GstAmt,TotalAmt,Payterms,CurCode,CurExRate,Delivered,Exported,ExportDate,DTG,Gst,MDTNo,[AcBillRef],[ShipName],[ShipAdd],[ShipAdd2],[ShipAdd3],[ShipAdd4],[ShipCity],[ShipPin],[Remarks],[CompanyNo],[DiscountPer],[DisPer],[IsCompleted],[IsApproved])" &
                                "Values (" & SafeSQL(dtr("StockInNo").ToString) & "," & SafeSQL(Format(dtr("TransDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("AgentId").ToString) & ",''," & SafeSQL(dtr("StockInNo").ToString) & ",'', " & SafeSQL(dtr("StockInNo").ToString) & ", 0 , 0 , 0 , 0 ,'', '', 1, 0, 0, getdate(), getdate(), 0, 'ADMIN', '', '', '', '', '', '', '', '', '', '', 0, 0, 1, 1)"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Import ErpStockReq - " & dtr("StockInNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteSQL(sQry)
                    End If
                End If
            End While
            dtr.Close()
            dtr = ReadNavRecord("Select * from StockRequestErp where IsNull(IsRead,0) = 0 Order by StockInNo")

            Dim i As Integer = 1
            While dtr.Read
                If dtr("StockInNo").ToString <> "" Then
                    If IsDOItemExists(dtr("StockInNo").ToString, dtr("ItemNo").ToString) = True Then

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Import ErpStockReq Insert Error : " & dtr("StockInNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL("Record Already Exits") & ")")

                    Else
                        sQry = "Insert into DeliveryOrdItem (OrdNo, ItemNo,UOM,Qty,Location,Remarks,ReasonCode,[LineNo],Price, Disper, DisPr, Discount, SubAmt, GstAmt, Description, SalesType )" &
                              "Values (" & SafeSQL(dtr("StockInNo").ToString) & "," & SafeSQL(dtr("ItemNo").ToString) & "," & SafeSQL(dtr("UOM").ToString) & "," & SafeSQL(dtr("Qty").ToString) & "," & SafeSQL(dtr("Location").ToString) & "," & SafeSQL(dtr("Remarks").ToString) & "," & SafeSQL(dtr("Reason").ToString) & "," & i & ",0,0,0,0,0,0,'','')"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Import ErpStockReq - " & dtr("StockInNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteSQL(sQry)
                        i = i + 1

                        myarray.Add(SafeSQL(dtr("StockInNo").ToString))

                    End If
                End If
            End While
            dtr.Close()
            For Each d As String In myarray
                sQry = "Update StockRequestErp set IsRead = 1 Where StockInNo = " & d
                ExecuteNavSQL(sQry)
            Next
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import ErpStockRequest - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
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
        dtr = ReadNavRecord("Select * from Category")
        While dtr.Read
            If dtr("ItmsGrpCod").ToString <> "" Then
                If IsExists("Select Code from Category where code=" & SafeSQL(dtr("Code"))) = False Then
                    ExecuteSQLAnother("Insert into Category(Code , Description) Values (" & SafeSQL(dtr("Code").ToString) & "," & SafeSQL(IIf(dtr("Description").ToString = "", dtr("Description").ToString, dtr("Description").ToString)) & ")")
                Else
                    ExecuteSQLAnother("Update Category set Description=" & SafeSQL(IIf(dtr("Description").ToString = "", dtr("Description").ToString, dtr("Description").ToString)) & " Where Code=" & SafeSQL(dtr("Code").ToString))
                End If
            End If
        End While
        dtr.Close()
        dtr.Dispose()
        dtr = Nothing
    End Sub

    Public Sub ImportBrand()
        Dim dtr As SqlDataReader
        'ExecuteSQL("Delete From Brand")
        dtr = ReadNavRecord("Select * from Brand")
        While dtr.Read
            If dtr("FirmCode").ToString <> "" Then
                If IsExists("Select Code from Brand where code=" & SafeSQL(dtr("Code"))) = False Then
                    ExecuteSQLAnother("Insert into Brand(Code , Description) Values (" & SafeSQL(dtr("Code").ToString) & "," & SafeSQL(IIf(dtr("Description").ToString = "", dtr("Description").ToString, dtr("Description").ToString)) & ")")
                Else
                    ExecuteSQLAnother("Update Brand set Description=" & SafeSQL(IIf(dtr("Description").ToString = "", dtr("Description").ToString, dtr("Description").ToString)) & " Where Code=" & SafeSQL(dtr("Code").ToString))
                End If
            End If
        End While
        dtr.Close()
        dtr.Dispose()
        dtr = Nothing
    End Sub

    Public Sub ImportProduct()
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
        Try
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Product - Insert Start'," & SafeSQL(NavCompanyName) & ",''" & ")")
            dtr = ReadNavRecord("Select * from Item where IsNull(IsRead,0) = 0")

            While dtr.Read
                If dtr("ItemNo").ToString <> "" Then
                    If IsItemExists(dtr("ItemNo")) Then
                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Product - Update Item'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("ItemNo").ToString) & ")")
                        sQry = "Update Item Set ItemNo = " & SafeSQL(dtr("ItemNo").ToString) &
                                ", Description = " & SafeSQL(dtr("Item_Name").ToString) &
                                ", ItemName = " & SafeSQL(dtr("Item_Name").ToString) &
                                ", ShortDesc= " & SafeSQL(dtr("Item_Name").ToString) &
                                ", ChineseDesc = " & SafeSQL(dtr("Item_Name").ToString) &
                                ", BaseUOM = " & SafeSQL(dtr("LooseUOm").ToString) &
                                ", UnitPrice = " & SafeSQL(dtr("Price1").ToString) &
                                ", Active = " & SafeSQL(dtr("Active").ToString) &
                                ", Category = " & SafeSQL(dtr("Category").ToString) &
                                ", Brand = " & SafeSQL(dtr("Brand").ToString) &
                                ", SubCategory = " & SafeSQL(dtr("Sub_Category").ToString) &
                                ", ToPDA = 1" &
                                ", SubBrand = " & SafeSQL(dtr("Sub_Brand").ToString) &
                                ", Favourite = " & SafeSQL(dtr("Favourite").ToString) &
                                ", BulkUOM = " & SafeSQL(dtr("BulkUOM").ToString) &
                                ", BulkQty = " & SafeSQL(dtr("BulkQty").ToString) &
                                ", LooseUOM = " & SafeSQL(dtr("LooseUOM").ToString) &
                                ", LooseQty = " & SafeSQL(dtr("LooseQty").ToString) &
                                ", MaxQty1 = " & If(IsDBNull(dtr("MaxQty1")), 0, dtr("MaxQty1")) &
                                ", MaxQty2 = " & If(IsDBNull(dtr("MaxQty2")), 0, dtr("MaxQty2")) &
                                ", MaxQty3 = " & If(IsDBNull(dtr("MaxQty3")), 0, dtr("MaxQty3")) &
                                ", MaxQty4 = " & If(IsDBNull(dtr("MaxQty4")), 0, dtr("MaxQty4")) &
                                ", MaxQty5 = " & If(IsDBNull(dtr("MaxQty5")), 0, dtr("MaxQty5")) &
                                ", Price1 = " & SafeSQL(dtr("Price1").ToString) &
                                ", Price2 = " & SafeSQL(dtr("Price2").ToString) &
                                ", Price3 = " & SafeSQL(dtr("Price3").ToString) &
                                ", Price4 = " & SafeSQL(dtr("Price4").ToString) &
                                ", Price5 = " & SafeSQL(dtr("Price5").ToString) &
                                ", PackType = " & SafeSQL(dtr("PackType").ToString) &
                                ", PriceGroup = " & SafeSQL(dtr("PriceGroup").ToString) &
                                ", [Return] = " & SafeSQL(dtr("Return").ToString) &
                                ", Sale = " & SafeSQL(dtr("Sale").ToString) &
                                ", GSTProdGroup = " & SafeSQL(dtr("VAT").ToString) &
                                ", [Trading_Discount] = " & SafeSQL(dtr("Trading_Discount").ToString) &
                                ", [Size_Unit] = " & SafeSQL(dtr("Size_Unit").ToString) &
                                ", Size  = " & SafeSQL(dtr("Size").ToString) &
                                ", IsPriced  = 0 " &
                                ", GrossWeight  = " & If(IsDBNull(dtr("M3")), 0, dtr("M3")) &
                                ", DTG = Getdate() " &
                                " Where ItemNo = " & SafeSQL(dtr("ItemNo").ToString)

                        'ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Import Item - " & dtr("ItemNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteSQL(sQry)

                    Else
                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Product - Insert Item'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("ItemNo").ToString) & ")")

                        sQry = "Insert into Item (ItemNo, Description, ItemName, ShortDesc, ChineseDesc, BaseUOM, UnitPrice, Active, Category, Brand, SubCategory, ToPDA, SubBrand, Favourite,  BulkUOM, BulkQty, LooseUOM, LooseQty, MaxQty1, MaxQty2, MaxQty3, MaxQty4, MaxQty5, Price1, Price2, Price3, Price4, Price5, PackType, PriceGroup, [Return], Sale, GSTProdGroup, [Trading_Discount], [Size_Unit], Size, GrossWeight,IsPriced, DTG)" &
                            "Values (" & SafeSQL(dtr("ItemNo").ToString) & "," & SafeSQL(dtr("Item_Name").ToString) & "," & SafeSQL(dtr("Item_Name").ToString) & "," & SafeSQL(dtr("Item_Name").ToString) & "," & SafeSQL(dtr("Item_Name").ToString) & "," & SafeSQL(dtr("LooseUOM").ToString) & "," & SafeSQL(dtr("Price1").ToString) & "," & SafeSQL(dtr("Active").ToString) & "," & SafeSQL(dtr("Category").ToString) &
                            "," & SafeSQL(dtr("Brand").ToString) & "," & SafeSQL(dtr("Sub_Category").ToString) & ",1," & SafeSQL(dtr("Sub_Brand").ToString) & "," & SafeSQL(dtr("Favourite").ToString) & "," & SafeSQL(dtr("BulkUOM").ToString) & "," & SafeSQL(dtr("BulkQty").ToString) & "," & SafeSQL(dtr("LooseUOM").ToString) & "," & SafeSQL(dtr("LooseQty").ToString) & "," & If(IsDBNull(dtr("MaxQty1")), 0, dtr("MaxQty1")) &
                            "," & If(IsDBNull(dtr("MaxQty2")), 0, dtr("MaxQty2")) & "," & If(IsDBNull(dtr("MaxQty3")), 0, dtr("MaxQty3")) & "," & If(IsDBNull(dtr("MaxQty4")), 0, dtr("MaxQty4")) & "," & If(IsDBNull(dtr("MaxQty5")), 0, dtr("MaxQty5")) & "," & SafeSQL(dtr("Price1").ToString) & "," & SafeSQL(dtr("Price2").ToString) & "," & SafeSQL(dtr("Price3").ToString) & "," & SafeSQL(dtr("Price4").ToString) & "," & SafeSQL(dtr("Price5").ToString) &
                            "," & SafeSQL(dtr("PackType").ToString) & "," & SafeSQL(dtr("PriceGroup").ToString) & "," & SafeSQL(dtr("Return").ToString) & "," & SafeSQL(dtr("Sale").ToString) & "," & SafeSQL(dtr("VAT").ToString) & "," & SafeSQL(dtr("Trading_Discount").ToString) & "," & SafeSQL(dtr("Size_Unit").ToString) & "," & SafeSQL(dtr("Size").ToString) & "," & If(IsDBNull(dtr("M3")), 0, dtr("M3")) & ", 0, Getdate())"


                        'ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Import Item - " & dtr("ItemNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteSQL(sQry)
                    End If
                    ExecuteNavAnotherSQL("Update item set IsRead = 1 where ItemNo = " & SafeSQL(dtr("ItemNo").ToString))
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

            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Product - Insert Finish'," & SafeSQL(NavCompanyName) & ",''" & ")")

        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Item - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try
    End Sub


    Public Sub ImportItemPrice()
        Try
            Dim dtr As SqlDataReader
            Dim sType As String = ""
            Dim sPriceGroup As String = ""
            Dim sItemNo As String = ""
            Dim sUOm As String = ""
            Dim dQty As Double = 0
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import ItemPr - Insert Start'," & SafeSQL(NavCompanyName) & ",''" & ")")
            ExecuteSQL("Delete from ItemPr where ItemNo in (Select ItemNO From Item Where IsNull(IsPriced,0) = 0)")
            dtr = ReadRecord("Select ItemNo, isnull(MaxQty1,0) as MaxQty1, isnull(MaxQty2,0) as MaxQty2, isnull(MaxQty3,0) as MaxQty3, isnull(MaxQty4,0) as MaxQty4, isnull(MaxQty5,0) as MaxQty5, isnull(Price1,0) as Price1, isnull(Price2,0) as Price2, isnull(Price3,0) as Price3, isnull(Price4,0) as Price4, isnull(Price5,0) as Price5, BulkUOM   from Item where ItemNo in (Select ItemNO From Item Where IsNull(IsPriced,0) = 0) order by ItemNo")
            While dtr.Read
                sType = "Customer Price Group"
                If dtr("Price1") <> 0 Then
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import ItemPr - Insert Itempr'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("ItemNo").ToString) & ")")

                    ExecuteSQLAnother("Insert into ItemPr(PriceGroup, ItemNo, UnitPrice, SalesType, Minqty, VariantCode, UOM, FromDate, ToDate) Values (" & SafeSQL("STANDARD") & "," & SafeSQL(dtr("ItemNo").ToString.Trim) & "," & IIf(IsDBNull(dtr("Price1")), 0, dtr("Price1")) & "," & SafeSQL(sType) & ",0," & SafeSQL("") & "," & SafeSQL(dtr("BulkUOM").ToString) & "," & SafeSQL(Format(Date.Now.AddDays(-1), "yyyyMMdd 00:00:00")) & "," & SafeSQL(Format(Date.Now.AddYears(1), "yyyyMMdd 23:59:59")) & ")")
                End If
                If dtr("Price2") <> 0 Then
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import ItemPr - Insert Itempr'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("ItemNo").ToString) & ")")

                    ExecuteSQLAnother("Insert into ItemPr(PriceGroup, ItemNo, UnitPrice, SalesType, Minqty, VariantCode, UOM, FromDate, ToDate) Values (" & SafeSQL("STANDARD") & "," & SafeSQL(dtr("ItemNo").ToString.Trim) & "," & IIf(IsDBNull(dtr("Price2")), 0, dtr("Price2")) & "," & SafeSQL(sType) & "," & dtr("MaxQty1") + 0.01 & "," & SafeSQL("") & "," & SafeSQL(dtr("BulkUOM").ToString) & "," & SafeSQL(Format(Date.Now.AddDays(-1), "yyyyMMdd 00:00:00")) & "," & SafeSQL(Format(Date.Now.AddYears(1), "yyyyMMdd 23:59:59")) & ")")
                End If

                If dtr("Price3") <> 0 Then
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import ItemPr - Insert Itempr'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("ItemNo").ToString) & ")")

                    ExecuteSQLAnother("Insert into ItemPr(PriceGroup, ItemNo, UnitPrice, SalesType, Minqty, VariantCode, UOM, FromDate, ToDate) Values (" & SafeSQL("STANDARD") & "," & SafeSQL(dtr("ItemNo").ToString.Trim) & "," & IIf(IsDBNull(dtr("Price3")), 0, dtr("Price3")) & "," & SafeSQL(sType) & "," & dtr("MaxQty2") + 0.01 & "," & SafeSQL("") & "," & SafeSQL(dtr("BulkUOM").ToString) & "," & SafeSQL(Format(Date.Now.AddDays(-1), "yyyyMMdd 00:00:00")) & "," & SafeSQL(Format(Date.Now.AddYears(1), "yyyyMMdd 23:59:59")) & ")")
                End If

                If dtr("Price4") <> 0 Then
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import ItemPr - Insert Itempr'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("ItemNo").ToString) & ")")

                    ExecuteSQLAnother("Insert into ItemPr(PriceGroup, ItemNo, UnitPrice, SalesType, Minqty, VariantCode, UOM, FromDate, ToDate) Values (" & SafeSQL("STANDARD") & "," & SafeSQL(dtr("ItemNo").ToString.Trim) & "," & IIf(IsDBNull(dtr("Price4")), 0, dtr("Price4")) & "," & SafeSQL(sType) & "," & dtr("MaxQty3") + 0.01 & "," & SafeSQL("") & "," & SafeSQL(dtr("BulkUOM").ToString) & "," & SafeSQL(Format(Date.Now.AddDays(-1), "yyyyMMdd 00:00:00")) & "," & SafeSQL(Format(Date.Now.AddYears(1), "yyyyMMdd 23:59:59")) & ")")
                End If

                If dtr("Price5") <> 0 Then
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import ItemPr - Insert Itempr'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("ItemNo").ToString) & ")")

                    ExecuteSQLAnother("Insert into ItemPr(PriceGroup, ItemNo, UnitPrice, SalesType, Minqty, VariantCode, UOM, FromDate, ToDate) Values (" & SafeSQL("STANDARD") & "," & SafeSQL(dtr("ItemNo").ToString.Trim) & "," & IIf(IsDBNull(dtr("Price5")), 0, dtr("Price5")) & "," & SafeSQL(sType) & "," & dtr("MaxQty4") + 0.01 & "," & SafeSQL("") & "," & SafeSQL(dtr("BulkUOM").ToString) & "," & SafeSQL(Format(Date.Now.AddDays(-1), "yyyyMMdd 00:00:00")) & "," & SafeSQL(Format(Date.Now.AddYears(1), "yyyyMMdd 23:59:59")) & ")")
                End If

                ExecuteSQLAnother("Update Item Set IsPriced =1 where ItemNo = " & SafeSQL(dtr("ItemNo").ToString.Trim))

            End While
            dtr.Close()
            dtr.Dispose()
            dtr = Nothing

            ExecuteSQL("Delete from ItemPr where UnitPrice = 0")
            ExecuteSQL("Update ItemPr Set MinPrice = 0 where MinPrice is Null")
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import ItemPr - Insert Finish'," & SafeSQL(NavCompanyName) & ",''" & ")")
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import ItemPr - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try
        '  ExecuteSQL("Update ItemPr set UOM = Item.BaseUOM from Item where Item.ItemNo = ItemPr.ItemNo and ItemPr.UOM = ''")
    End Sub
    Public Sub ImportUserBrand()
        Try

            Dim dtr As SqlDataReader
            Dim IsFirst As String = "Y"

            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import UserBrand - Insert Start'," & SafeSQL(NavCompanyName) & ",''" & ")")

            dtr = ReadNavRecord("Select * from userbrand Where IsNull(IsRead,0) = 0")

            'If dtr.Read = True Then
            '    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import UserBrand - BackUp UserBrand'," & SafeSQL(NavCompanyName) & ",'BackUp All')")

            '    ExecuteSQLAnother("Insert into UserBrandBackup select UserID, Brand, Getdate() from userbrand")

            '    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import UserBrand - Delete UserBrand'," & SafeSQL(NavCompanyName) & ",'Delete All')")

            '    ExecuteSQLAnother("Delete from userbrand")

            'End If

            While dtr.Read

                'If IsFirst = "Y" Then
                '    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import UserBrand - BackUp UserBrand'," & SafeSQL(NavCompanyName) & ",'BackUp All')")

                '    ExecuteSQLAnother("Insert into UserBrandBackup select UserID, Brand, Getdate() from userbrand")

                '    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import UserBrand - Delete UserBrand'," & SafeSQL(NavCompanyName) & ",'Delete All')")

                '    ExecuteSQLAnother("Delete from userbrand")
                'End If

                If IsExists("Select * from UserBrand where UserId = " & SafeSQL(dtr("UserId").ToString.Trim) & " and Brand = " & SafeSQL(dtr("Brand").ToString.Trim)) = True Then
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import UserBrand - Update UserBrand'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("Brand").ToString) & ")")

                    ExecuteSQLAnother("Update userbrand set Brand = " & SafeSQL(dtr("Brand").ToString.Trim) & " Where UserId = " & SafeSQL(dtr("UserId").ToString.Trim) & " and Brand = " & SafeSQL(dtr("Brand").ToString.Trim))
                Else
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import UserBrand - Insert UserBrand'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("Brand").ToString) & ")")

                    ExecuteSQLAnother("Insert into userbrand(UserId, Brand) Values (" & SafeSQL(dtr("UserId").ToString.Trim) & "," & SafeSQL(dtr("Brand").ToString.Trim) & ")")
                End If
                ExecuteNavAnotherSQL("Update UserBrand set IsRead = 1 where UserId =  " & SafeSQL(dtr("UserId").ToString.Trim) & " and Brand = " & SafeSQL(dtr("Brand").ToString.Trim))

                'IsFirst = "N"
            End While

            dtr.Close()
            dtr.Dispose()
            dtr = Nothing


            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import UserBrand - Start Data Comparision'," & SafeSQL(NavCompanyName) & ",''" & ")")


            Dim sSql As String = " Insert into UserBrandBackup " &
                                 " select A.UserID, A.Brand, Getdate() " &
                                 " FROM UserBrand A LEFT JOIN  " & NavDBName & ".dbo.UserBrand B ON (A.UserID = B.UserID AND A.Brand = B.Brand)" &
                                 " WHERE B.UserID IS NULL and B.Brand is null "
            ExecuteSQLAnother(sSql)

            sSql = " Delete from userbrand " &
                   " FROM UserBrand A " &
                   " LEFT JOIN  " & NavDBName & ".dbo.UserBrand B ON (A.UserID = B.UserID AND A.Brand = B.Brand) " &
                   " WHERE B.UserID IS NULL and B.Brand is null "
            ExecuteSQLAnother(sSql)

            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import UserBrand - End Data Comparision'," & SafeSQL(NavCompanyName) & ",''" & ")")


            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import UserBrand - Insert Finish'," & SafeSQL(NavCompanyName) & ",''" & ")")
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import UserBrand - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try


    End Sub
    Public Sub ImportBarangay()
        Try
            Dim dtr As SqlDataReader
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Barangay - Insert Start'," & SafeSQL(NavCompanyName) & ",''" & ")")

            dtr = ReadNavRecord("Select * from Barangay where IsNull(IsRead,0)=0")

            While dtr.Read
                ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Barangay - Delete & Insert Barangay'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("Code").ToString) & ")")

                ExecuteSQLAnother("Delete from Barangay where code =  " & SafeSQL(dtr("Code").ToString.Trim))
                ExecuteSQLAnother("Insert into Barangay(Code, BarangayName,CityCode,CityName,ProvinceCode) Values (" & SafeSQL(dtr("Code").ToString.Trim) & "," & SafeSQL(dtr("BarangayName").ToString) & "," & SafeSQL(dtr("CityCode").ToString) & "," & SafeSQL(dtr("CityName").ToString) & "," & SafeSQL(dtr("ProvinceCode").ToString.Trim) & ")")
                ExecuteNavAnotherSQL("Update Barangay set IsRead = 1 where code =  " & SafeSQL(dtr("Code").ToString.Trim))
            End While

            dtr.Close()
            dtr.Dispose()
            dtr = Nothing
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Barangay - Insert Finish'," & SafeSQL(NavCompanyName) & ",''" & ")")
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Barangay - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try


    End Sub
    Public Sub ImportTerritory()
        Try
            Dim dtr As SqlDataReader
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Territory - Insert Start'," & SafeSQL(NavCompanyName) & ",''" & ")")

            dtr = ReadNavRecord("Select * from Territory where IsNull(IsRead,0) = 0")

            While dtr.Read
                ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Territory - Delete & Insert Territory'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("Code").ToString) & ")")

                ExecuteSQLAnother("Delete from Territory where code =  " & SafeSQL(dtr("Code").ToString.Trim))
                ExecuteSQLAnother("Insert into Territory(Code, Province, District, SubDistrict) Values (" & SafeSQL(dtr("Code").ToString.Trim) & "," & SafeSQL(dtr("Province").ToString) & "," & SafeSQL(dtr("District").ToString) & "," & SafeSQL(dtr("SubDistrict").ToString) & ")")
                ExecuteNavAnotherSQL("Update Territory set IsRead=1 where  code =  " & SafeSQL(dtr("Code").ToString.Trim))
            End While

            dtr.Close()
            dtr.Dispose()
            dtr = Nothing
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Territory - Insert Finish'," & SafeSQL(NavCompanyName) & ",''" & ")")
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Territory - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try


    End Sub
    Public Sub ImportProvince()
        Try
            Dim dtr As SqlDataReader
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Province - Insert Start'," & SafeSQL(NavCompanyName) & ",''" & ")")

            dtr = ReadNavRecord("Select * from Province where IsNull(IsRead,0) = 0")

            While dtr.Read
                ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Province - Delete & Insert Province'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("Code").ToString) & ")")

                ExecuteSQLAnother("Delete from Province where Code =  " & SafeSQL(dtr("Code").ToString.Trim))
                ExecuteSQLAnother("Insert into Province(Code, Description) Values (" & SafeSQL(dtr("Code").ToString.Trim) & "," & SafeSQL(dtr("Description").ToString) & ")")
                ExecuteNavAnotherSQL("Update Province set IsRead =1 where Code = " & SafeSQL(dtr("Code").ToString.Trim))
            End While

            dtr.Close()
            dtr.Dispose()
            dtr = Nothing
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Province - Insert Finish'," & SafeSQL(NavCompanyName) & ",''" & ")")
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Province - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try


    End Sub
    Public Sub ImportInstitutionType()
        Try
            Dim dtr As SqlDataReader
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Institution - Insert Start'," & SafeSQL(NavCompanyName) & ",''" & ")")

            dtr = ReadNavRecord("Select * from InstitutionType where IsNull(IsRead,0) = 0")

            While dtr.Read
                ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Institution - Delete & Insert InstitutionType'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("Code").ToString) & ")")

                ExecuteSQLAnother("Delete from InstitutionType where code = " & SafeSQL(dtr("Code").ToString.Trim))
                ExecuteSQLAnother("Insert into InstitutionType(Code, Description,DisplayNo) Values (" & SafeSQL(dtr("Code").ToString.Trim) & "," & SafeSQL(dtr("Description").ToString) & "," & SafeSQL(dtr("DisplayNo").ToString) & ")")
                ExecuteNavAnotherSQL("Update InstitutionType set IsRead =1 where Code = " & SafeSQL(dtr("Code").ToString.Trim))
            End While

            dtr.Close()
            dtr.Dispose()
            dtr = Nothing
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Institution - Insert Finish'," & SafeSQL(NavCompanyName) & ",''" & ")")
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Institution - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try


    End Sub
    Public Sub ImportProductList()
        Try
            Dim dtr As SqlDataReader
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import ProductList - Insert Start'," & SafeSQL(NavCompanyName) & ",''" & ")")

            dtr = ReadNavRecord("Select * from ProductList where IsNull(IsRead,0) = 0")

            While dtr.Read
                ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import ProductList - Delete & Insert ProductList'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("Code").ToString) & ")")

                ExecuteSQLAnother("Delete from ProductList where code = " & SafeSQL(dtr("Code").ToString.Trim))
                ExecuteSQLAnother("Insert into ProductList(Code, Description,DisplayNo) Values (" & SafeSQL(dtr("Code").ToString.Trim) & "," & SafeSQL(dtr("Description").ToString) & "," & SafeSQL(dtr("DisplayNo").ToString) & ")")
                ExecuteNavAnotherSQL("Update ProductList set IsRead =1 where Code = " & SafeSQL(dtr("Code").ToString.Trim))
            End While

            dtr.Close()
            dtr.Dispose()
            dtr = Nothing
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import ProductList - Insert Finish'," & SafeSQL(NavCompanyName) & ",''" & ")")
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import ProductList - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try


    End Sub
    Public Sub ImportLocationType()
        Try
            Dim dtr As SqlDataReader
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import LocationType - Insert Start'," & SafeSQL(NavCompanyName) & ",''" & ")")

            dtr = ReadNavRecord("Select * from LocationType where IsNull(IsRead,0) = 0")

            While dtr.Read
                ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import LocationType - Delete & Insert LocationType'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("Code").ToString) & ")")

                ExecuteSQLAnother("Delete from LocationType where code = " & SafeSQL(dtr("Code").ToString.Trim))
                ExecuteSQLAnother("Insert into LocationType(Code, Description,DisplayNo) Values (" & SafeSQL(dtr("Code").ToString.Trim) & "," & SafeSQL(dtr("Description").ToString) & "," & SafeSQL(dtr("DisplayNo").ToString) & ")")
                ExecuteNavAnotherSQL("Update LocationType set IsRead =1 where Code = " & SafeSQL(dtr("Code").ToString.Trim))
            End While
            dtr.Close()
            dtr.Dispose()
            dtr = Nothing
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import LocationType - Insert Finish'," & SafeSQL(NavCompanyName) & ",''" & ")")
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import LocationType - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try


    End Sub
    Public Sub ImportBusinessType()
        Try
            Dim dtr As SqlDataReader
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import BusinessType - Insert Start'," & SafeSQL(NavCompanyName) & ",''" & ")")

            dtr = ReadNavRecord("Select * from BusinessType where IsNull(IsRead,0) = 0")

            While dtr.Read
                ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import BusinessType - Delete & Insert BusinessType'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("Code").ToString) & ")")

                ExecuteSQLAnother("Delete from BusinessType where code = " & SafeSQL(dtr("Code").ToString.Trim))
                ExecuteSQLAnother("Insert into BusinessType(Code, Description,DisplayNo) Values (" & SafeSQL(dtr("Code").ToString.Trim) & "," & SafeSQL(dtr("Description").ToString) & "," & SafeSQL(dtr("DisplayNo").ToString) & ")")
                ExecuteNavAnotherSQL("Update BusinessType set IsRead =1 where Code = " & SafeSQL(dtr("Code").ToString.Trim))
            End While

            dtr.Close()
            dtr.Dispose()
            dtr = Nothing
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import BusinessType - Insert Finish'," & SafeSQL(NavCompanyName) & ",''" & ")")
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import BusinessType - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try


    End Sub
    Public Sub ImportMapCustomer()
        Try
            Dim dtr As SqlDataReader
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import MapCustomer - Insert Start'," & SafeSQL(NavCompanyName) & ",''" & ")")

            dtr = ReadNavRecord("Select * from MapCustomer where IsNull(IsRead,0) = 0")

            While dtr.Read
                ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import MapCustomer - Update Invoice & OrderHdr '," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("ERPCustNo").ToString + "-" + dtr("NewCustNo").ToString) & ")")

                ExecuteSQLAnother("Update Invoice Set CustId =  " & SafeSQL(dtr("ERPCustNo").ToString.Trim) & " Where CustId = " & SafeSQL(dtr("NewCustNo").ToString.Trim))

                ExecuteSQLAnother("Update OrderHdr Set CustId =  " & SafeSQL(dtr("ERPCustNo").ToString.Trim) & " Where CustId = " & SafeSQL(dtr("NewCustNo").ToString.Trim))

                ExecuteNavAnotherSQL("Update MapCustomer set IsRead =1 where NewCustNo = " & SafeSQL(dtr("NewCustNo").ToString.Trim))

            End While

            dtr.Close()
            dtr.Dispose()
            dtr = Nothing
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import MapCustomer - Insert Finish'," & SafeSQL(NavCompanyName) & ",''" & ")")
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import MapCustomer - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try


    End Sub
    Public Sub ImportShopType()
        Try
            Dim dtr As SqlDataReader
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import ShopType - Insert Start'," & SafeSQL(NavCompanyName) & ",''" & ")")

            dtr = ReadNavRecord("Select * from ShopType where IsNull(IsRead,0) = 0")

            While dtr.Read
                ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import ShopType - Delete & Insert ShopType'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("Code").ToString) & ")")

                ExecuteSQLAnother("Delete from ShopType where code = " & SafeSQL(dtr("Code").ToString.Trim))
                ExecuteSQLAnother("Insert into ShopType(Code, Description,DisplayNo) Values (" & SafeSQL(dtr("Code").ToString.Trim) & "," & SafeSQL(dtr("Description").ToString) & "," & SafeSQL(dtr("DisplayNo").ToString) & ")")
                ExecuteNavAnotherSQL("Update ShopType set IsRead =1 where Code = " & SafeSQL(dtr("Code").ToString.Trim))
            End While

            dtr.Close()
            dtr.Dispose()
            dtr = Nothing
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import ShopType - Insert Finish'," & SafeSQL(NavCompanyName) & ",''" & ")")
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import ShopType - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try

    End Sub
    Public Sub ImportLocation()
        Try
            Dim dtr As SqlDataReader
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Location - Insert Start'," & SafeSQL(NavCompanyName) & ",''" & ")")
            dtr = ReadNavRecord("Select * from Location where IsNull(IsRead,0) = 0")
            While dtr.Read
                If IsExists("Select Code from Location where code=" & SafeSQL(dtr("Code"))) = False Then
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Location - Insert Location'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("Code").ToString) & ")")

                    ExecuteSQLAnother("Insert into Location(Code, Name) Values (" & SafeSQL(dtr("Code").ToString) & "," & SafeSQL(dtr("Name").ToString) & ")")
                Else
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Location - Update Location'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("Code").ToString) & ")")

                    ExecuteSQLAnother("Update Location set Name=" & SafeSQL(IIf(dtr("Name").ToString = "", dtr("Name").ToString, dtr("Name").ToString)) & " Where Code=" & SafeSQL(dtr("Code").ToString))
                End If
                ExecuteNavAnotherSQL("Update Location set IsRead = 1 where Code = " & SafeSQL(dtr("Code")))
            End While

            dtr.Close()
            dtr.Dispose()
            dtr = Nothing
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Location - Insert Finish'," & SafeSQL(NavCompanyName) & ",''" & ")")
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Location - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try


    End Sub
    Public Sub ImportArea()
        Try
            Dim dtr As SqlDataReader
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Area - Insert Start'," & SafeSQL(NavCompanyName) & ",''" & ")")

            dtr = ReadNavRecord("Select Distinct SalesUnit, Area from Area where IsNull(IsRead,0) = 0")

            While dtr.Read
                If IsExists("Select * from CustAgent where AgentId = " & SafeSQL(dtr("SalesUnit").ToString)) = True Then
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Area - Update CustAgent'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("SalesUnit").ToString) & ")")

                    ExecuteSQLAnother("Update CustAgent Set CustAgentId = " & SafeSQL(dtr("Area").ToString) & " Where AgentId = " & SafeSQL(dtr("SalesUnit").ToString))
                Else
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Area - Insert CustAgent'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("SalesUnit").ToString) & ")")

                    ExecuteSQLAnother("Insert into CustAgent (AgentID, CustAgentID, Position) Values (" & SafeSQL(dtr("SalesUnit").ToString) & "," & SafeSQL(dtr("Area").ToString) & ",1)")
                End If
                ExecuteNavAnotherSQL("Update Area Set IsRead = 1 where SalesUnit = " & SafeSQL(dtr("SalesUnit").ToString) & " and Area = " & SafeSQL(dtr("Area").ToString))
            End While

            dtr.Close()
            dtr.Dispose()
            dtr = Nothing
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Area - Insert Finish'," & SafeSQL(NavCompanyName) & ",''" & ")")
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Area - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try

    End Sub
    Public Sub ImportProductTeam()
        Try
            Dim dtr As SqlDataReader
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import ProductTeam - Insert Start'," & SafeSQL(NavCompanyName) & ",''" & ")")

            If bDeleteIns = True Then
                ExecuteSQLAnother("Delete from ProdTeam")
                ExecuteSQLAnother("Insert into ProdTeam Select Team,ItemNo from " & NavDBName & ".dbo.ProductTeam")
                Exit Sub
            Else
                dtr = ReadNavRecord("Select Distinct Team, ItemNo  from ProductTeam where IsNull(IsRead,0) =0")
            End If

            While dtr.Read
                ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import ProductTeam - Delete & Insert ProdTeam'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("Team").ToString) & ")")

                ExecuteSQLAnother("Delete from ProdTeam where team =  " & SafeSQL(dtr("Team").ToString) & " and ItemNo = " & SafeSQL(dtr("ItemNo").ToString))

                ExecuteSQLAnother("Insert into ProdTeam(Team, ItemNo) Values (" & SafeSQL(dtr("Team").ToString) & "," & SafeSQL(dtr("ItemNo").ToString) & ")")

                ExecuteNavAnotherSQL("Update ProductTeam set IsRead =1 where Team = " & SafeSQL(dtr("Team").ToString) & " and ItemNo = " & SafeSQL(dtr("ItemNo").ToString))

            End While

            dtr.Close()
            dtr.Dispose()
            dtr = Nothing
            'ExecuteSQLAnother("Delete from ProdTeam Where Team Not in  (Select Team from SimplrMMIntegration.dbo.ProductTeam P Where P.ItemNo = ProdTeam.ItemNo)")
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import ProductTeam - Insert Finish'," & SafeSQL(NavCompanyName) & ",''" & ")")
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import ProductTeam - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try

    End Sub
    Public Sub ImportTeam()
        Try
            Dim dtr As SqlDataReader
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Team - Insert Start'," & SafeSQL(NavCompanyName) & ",''" & ")")

            dtr = ReadNavRecord("Select Distinct Team, SalesUnit  from Team ")
            ExecuteSQLAnother("Delete from Team")
            While dtr.Read

                ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Team - Insert Team'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("Team").ToString) & ")")

                ExecuteSQLAnother("Insert into Team(Team, SalesUnit) Values (" & SafeSQL(dtr("Team").ToString) & "," & SafeSQL(dtr("SalesUnit").ToString) & ")")

            End While
            ExecuteNavAnotherSQL("Update Team set IsRead =1")
            dtr.Close()
            dtr.Dispose()
            dtr = Nothing
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Team - Insert Finish'," & SafeSQL(NavCompanyName) & ",''" & ")")
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Team - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try

    End Sub
    Public Sub ImportBank()
        Try
            Dim dtr As SqlDataReader
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Bank - Insert Start'," & SafeSQL(NavCompanyName) & ",''" & ")")
            dtr = ReadNavRecord("Select * from Bank where IsNull(IsRead,0) = 0")
            While dtr.Read
                If IsExists("Select Code from Bank where Code =" & SafeSQL(dtr("Code"))) = False Then
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Bank - Insert Bank'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("Code").ToString) & ")")

                    ExecuteSQLAnother("Insert into Bank(Code, BankName,Active) Values (" & SafeSQL(dtr("Code").ToString) & "," & SafeSQL(dtr("BankName").ToString) & "," & SafeSQL(dtr("Active").ToString) & ")")
                Else
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Bank - Update Bank'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("Code").ToString) & ")")

                    ExecuteSQLAnother("Update Bank set BankName=" & SafeSQL(IIf(dtr("BankName").ToString = "", dtr("BankName").ToString, dtr("BankName").ToString)) & ",Active = " & SafeSQL(dtr("Active").ToString) & " Where Code=" & SafeSQL(dtr("Code").ToString))
                End If
                ExecuteNavAnotherSQL("Update Bank Set IsRead =1 where Code = " & SafeSQL(dtr("Code").ToString))
            End While

            dtr.Close()
            dtr.Dispose()
            dtr = Nothing
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Bank - Insert Finish'," & SafeSQL(NavCompanyName) & ",''" & ")")
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import Bank - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try

    End Sub
    Public Sub ImportStockTakeProductTeam()
        Dim dtr As SqlDataReader
        dtr = ReadNavRecord("Select * from StockTakeProductTeam")
        While dtr.Read
            If IsExists("Select Team from StockTakeProductTeam where Team =" & SafeSQL(dtr("Team"))) = False Then
                ExecuteSQLAnother("Insert into StockTakeProductTeam(Team, ItemNo) Values (" & SafeSQL(dtr("Team").ToString) & "," & SafeSQL(dtr("ItemNo").ToString) & ")")
            Else
                ExecuteSQLAnother("Update StockTakeProductTeam set ItemNo =" & SafeSQL(IIf(dtr("ItemNo").ToString = "", dtr("ItemNo").ToString, dtr("ItemNo").ToString)) & " Where Team =" & SafeSQL(dtr("Team").ToString))
            End If
        End While

        dtr.Close()
        dtr.Dispose()
        dtr = Nothing
    End Sub
    Public Sub ImportCustAgent()
        Try
            Dim dtr As SqlDataReader
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import CustAgent - Insert Start'," & SafeSQL(NavCompanyName) & ",''" & ")")
            'ExecuteSQL("Update SalesAgent Set Active = 0")
            dtr = ReadRecord("Select Distinct AgentID as SalesAgent from CustAgent  order by AgentID")

            While dtr.Read
                If dtr("SalesAgent").ToString <> "" Then

                    If IsMDTExists(dtr("SalesAgent").ToString) = False Then
                        ExecuteSQLAnother("Insert into MDT(MDTNo, Description, AgentId, Location, RouteNo, VehicleID, SolutionName) Values (" & SafeSQL(dtr("SalesAgent").ToString.Trim) & "," & SafeSQL(dtr("SalesAgent").ToString.Trim) & "," & SafeSQL("") & "," & SafeSQL(dtr("SalesAgent").ToString.Trim) & "," & SafeSQL("") & "," & SafeSQL("") & ", 'SALES')")
                    Else
                        '     ExecuteSQLAnother("Update SalesAgent Set Name = " & SafeSQL(dtr("SalesAgent").ToString) & ", Active = 1 Where Code = " & SafeSQL(dtr("SalesAgent").ToString))
                    End If

                    If IsNoSeriesExists(dtr("SalesAgent").ToString) = False Then
                        ExecuteSQLAnother("Insert into NoSeries (MDTNo, DocType, ConditionMaster, ConditionType, ConditionValue, Prefix, LastNumber, NoLength, StartDate, EndDate) Select " & SafeSQL(dtr("SalesAgent").ToString.Trim) & ", DocType, ConditionMaster, ConditionType, ConditionValue, CASE WHEN DocType='ITEMTRANS' THEN " & SafeSQL(dtr("SalesAgent").ToString.Trim) & " +SubString(DocType,1,2) ELSE " & SafeSQL(dtr("SalesAgent").ToString.Trim) & " +SubString(DocType,1,1) END, 0 as LastNumber, NoLength, StartDate, EndDate from NoSeries where MDTNo='M1'")
                    Else
                        '     ExecuteSQLAnother("Update SalesAgent Set Name = " & SafeSQL(dtr("SalesAgent").ToString) & ", Active = 1 Where Code = " & SafeSQL(dtr("SalesAgent").ToString))
                    End If

                End If
            End While
            dtr.Close()
            dtr.Dispose()
            dtr = Nothing
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import CustAgent - Insert Finish'," & SafeSQL(NavCompanyName) & ",''" & ")")
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import CustAgent - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try


    End Sub

    Private Sub ExportInvoices()
        Dim sCurCode As String = ""
        Dim arrInvNo = New ArrayList
        Dim sQry As String = ""
        Dim dTransport As Double = 0
        sCurCode = "" 'dtr1("LCY Code")
        Dim dExRate As Double = 0
        Dim dtr As SqlDataReader = Nothing
        Dim dtr1 As SqlDataReader = Nothing
        Dim sExtDocNo As String = ""
        Dim sSql As String
        Try
            dtr = ReadRecord("Select  InvNo from Invoice where isnull(Exported,0) = 0 and IsNull(Void,0) = 0 Order by InvNo")
            While dtr.Read
                If arrInvNo.Contains(dtr("InvNo").ToString) = False Then arrInvNo.Add(dtr("InvNo").ToString)
            End While
            dtr.Close()
            dtr.Dispose()


            For i = 0 To arrInvNo.Count - 1
                Dim j = 0
                Try

                    dtr = ReadRecord("Select IV.InvNo,IV.InvDt,IV.OrdNo,IV.CustId,IV.AgentId,IV.SubTotal,IV.GstAmt,IV.TotalAmt,IV.PaidAmt,IV.PayTerms,IV.Void,IV.GST,IV.Remarks,IsNull(IV.MDTNo,'') as SalesUnit, IV.Discount,IsNull(IV.LineCount,0) as AllItemLine, Format(IV.DoDt,'yyyy-MM-dd HH:mm:ss') as DeliveryDate, IsNull(IV.PONo,'') as PONo, IsNull(IV.Disper,0) as Disper from Invoice IV Inner Join Customer On IV.CustId = Customer.CustNo where IV.InvNo =  " & SafeSQL(arrInvNo(i)) & " Order by IV.InvNo")

                    While dtr.Read
                        j = j + 1
                        ExecuteNavSQL("Delete from Invoice where InvNo=" & SafeSQL(dtr("InvNo").ToString))
                        ExecuteNavSQL("Delete from InvItem where InvNo=" & SafeSQL(dtr("InvNo").ToString))

                        If IsDBNull(dtr("DeliveryDate")) = True Then
                            sQry = "Insert into Invoice (InvNo,InvDt,OrdNo,CustId,AgentId,SubTotal,GstAmt,TotalAmt,PaidAmt,PayTerms,Void,GST,Remarks,SalesUnit,Discount,AllItemLine,PONo,DiscountPer, CreatedDate ) Values (" & SafeSQL(dtr("InvNo").ToString) &
                            "," & SafeSQL(Format(dtr("InvDt"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("OrdNo").ToString) & "," & SafeSQL(dtr("CustId").ToString) & "," & SafeSQL(dtr("AgentId").ToString) &
                            "," & SafeSQL(dtr("SubTotal").ToString) & "," & SafeSQL(dtr("GstAmt").ToString) & "," & SafeSQL(dtr("TotalAmt").ToString) & "," & SafeSQL(dtr("PaidAmt").ToString) &
                            "," & SafeSQL(dtr("PayTerms").ToString) & "," & SafeSQL(dtr("Void").ToString) & "," & SafeSQL(dtr("GST").ToString) & "," & SafeSQL(dtr("Remarks").ToString) & "," & SafeSQL(dtr("SalesUnit").ToString) & "," & SafeSQL(dtr("Discount").ToString) & "," & SafeSQL(dtr("AllItemLine").ToString) & "," & SafeSQL(dtr("PONo")) & "," & SafeSQL(dtr("DisPer")) & ", Format(SYSDATETIMEOFFSET() AT TIME ZONE 'SE Asia Standard Time','yyyy-MM-dd HH:mm:ss') )"
                        Else
                            sQry = "Insert into Invoice (InvNo,InvDt,OrdNo,CustId,AgentId,SubTotal,GstAmt,TotalAmt,PaidAmt,PayTerms,Void,GST,Remarks,SalesUnit,Discount,AllItemLine,DeliveryDate,PONo,DiscountPer, CreatedDate ) Values (" & SafeSQL(dtr("InvNo").ToString) &
                               "," & SafeSQL(Format(dtr("InvDt"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("OrdNo").ToString) & "," & SafeSQL(dtr("CustId").ToString) & "," & SafeSQL(dtr("AgentId").ToString) &
                               "," & SafeSQL(dtr("SubTotal").ToString) & "," & SafeSQL(dtr("GstAmt").ToString) & "," & SafeSQL(dtr("TotalAmt").ToString) & "," & SafeSQL(dtr("PaidAmt").ToString) &
                               "," & SafeSQL(dtr("PayTerms").ToString) & "," & SafeSQL(dtr("Void").ToString) & "," & SafeSQL(dtr("GST").ToString) & "," & SafeSQL(dtr("Remarks").ToString) & "," & SafeSQL(dtr("SalesUnit").ToString) & "," & SafeSQL(dtr("Discount").ToString) & "," & SafeSQL(dtr("AllItemLine").ToString) & "," & SafeSQL(Convert.ToDateTime(dtr("DeliveryDate"))) & "," & SafeSQL(dtr("PONo")) & "," & SafeSQL(dtr("DisPer")) & ", Format(SYSDATETIMEOFFSET() AT TIME ZONE 'SE Asia Standard Time','yyyy-MM-dd HH:mm:ss') )"
                        End If

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export Invoices - " & dtr("InvNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteNavSQL(sQry)

                    End While
                    dtr.Close()

                Catch ex As Exception
                    dtr.Close()
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export Invoices - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try

                Try
                    If j > 0 Then
                        dtr1 = ReadRecord("Select Distinct InvNo,[LineNo],InvItem.ItemNo,InvItem.Uom,InvItem.Qty,Round((NULLIf(UBK.BaseQty, 0) * InvItem.Price),2) as Price,CASE WHEN InvItem.DisPer<>0 then Round((Qty * Price) * (InvItem.DisPer / 100),2) else 0 End as Discount,InvItem.DisPer,InvItem.SubAmt,InvItem.SalesType,ItemPr.PriceGroup, " &
                            " InvItem.ReasonCode,InvItem.AttachedLineNo,Item.BulkUOM,qty/NULLIf(UBK.BaseQty,0)  as BulkQty,LooseUOM,(qty % NULLIf(UBK.BaseQty,0))/NULLIf(ULS.BaseQty,0) as LooseQty," &
                            " Round((Qty * Price),2) as SubAmtBefDis,InvItem.PromoOffer  From InvItem Inner Join Item On InvItem.ItemNo = Item.ItemNo Inner Join ItemPr on InvItem.ItemNo=ItemPr.ItemNo Left Join UOM UBK On Item.ItemNo = UBK.ItemNo and UBK.UOM=Item.BulkUOM " &
                            " Left Join UOM ULS On Item.ItemNo = ULS.ItemNo and ULS.UOM=Item.LooseUOM  Where InvItem.InvNo =  " & SafeSQL(arrInvNo(i)) & " Order by InvItem.InvNo")
                        Dim k = 0
                        While dtr1.Read
                            k = k + 1

                            Dim sPromooffer As String = ""
                            Dim attachedline As String = "0"
                            If dtr1("SalesType") = "F" And dtr1("PromoOffer") <> "0" Then
                                Dim rs As SqlDataReader
                                rs = ReadRecordAnother(" Select PromoOffer from InvItem OI  Where OI.[LineNo]=" & dtr1("LineNo") & " and OI.InvNo =  " & SafeSQL(arrInvNo(i)) & " Order by OI.InvNo")
                                If rs.Read = True Then
                                    sPromooffer = rs("PromoOffer").ToString
                                End If
                                rs.Close()
                                rs.Dispose()

                                rs = ReadRecordAnother(" Select [LineNo] as LineNum from InvItem OI  Where OI.PromoID=" & SafeSQL(sPromooffer) & " and OI.InvNo =  " & SafeSQL(arrInvNo(i)) & " Order by OI.InvNo")
                                If rs.Read = True Then
                                    attachedline = rs("LineNum")
                                End If
                                rs.Close()
                                rs.Dispose()


                            End If

                            sSql = "Insert into InvItem (InvNo,[LineNo],ItemNo,Uom,Qty,Price,Discount,DisPer,SubAmt,SalesType,PriceGroup,ReasonCode,AttachedLineNo,BulkUom,BulkQty,LooseUOM,LooseQty,SubAmtBefDis) Values (" & SafeSQL(dtr1("InvNo").ToString) &
                                 "," & SafeSQL(dtr1("LineNo")) & "," & SafeSQL(dtr1("ItemNo").ToString) & "," & SafeSQL(dtr1("UOM").ToString) & "," & SafeSQL(dtr1("Qty").ToString) &
                                 "," & SafeSQL(dtr1("Price").ToString) & "," & SafeSQL(dtr1("Discount").ToString) & "," & SafeSQL(dtr1("Disper").ToString) & "," & SafeSQL(dtr1("SubAmt").ToString) &
                                 "," & SafeSQL(dtr1("SalesType").ToString) & "," & SafeSQL(dtr1("PriceGroup").ToString) & "," & SafeSQL(dtr1("ReasonCode").ToString) & "," & SafeSQL(attachedline) & "," & SafeSQL(dtr1("BulkUOM").ToString) & "," & dtr1("BulkQty") & "," & SafeSQL(dtr1("LooseUOM").ToString) & "," & dtr1("LooseQTY") & "," & dtr1("SubAmtBefDis") & ")"

                            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export InvItem - " & dtr1("InvNo").ToString) & SafeSQL(arrInvNo(i)) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sSql) & ")")

                            ExecuteNavSQL(sSql)
                        End While
                        dtr1.Close()
                        If k > 0 Then
                            ExecuteSQLAnother("Update Invoice Set Exported = 1 Where InvNo = " & SafeSQL(arrInvNo(i)))
                        Else
                            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export Invitem'," & SafeSQL(NavCompanyName) & "," & SafeSQL("Invitem not found") & ")")
                        End If
                    End If
                Catch ex As Exception
                    dtr1.Close()
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export Invoice'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                    Exit Sub
                End Try
            Next
        Catch ex As Exception
            dtr.Close()
            dtr1.Close()
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
        Dim dtr As SqlDataReader = Nothing
        Dim sExtDocNo As String = ""

        Try
            dtr = ReadRecord("Select  StockInNo from StockInItem where IsNull(approved,0) =1 and (exported is NULL or Exported = 0) Order by StockInNo")
            While dtr.Read
                If arrInvNo.Contains(dtr("StockInNo").ToString) = False Then arrInvNo.Add(dtr("StockInNo").ToString)
            End While
            dtr.Close()
            dtr.Dispose()

            For i = 0 To arrInvNo.Count - 1
                Try
                    dtr = ReadRecord("Select H.StockInNo, H.TransDate, H.AgentID, H.Location, ItemNo, UOM, Qty-TransitQty as Qty , H.Remarks, Reason,[LineNo] " &
                        " from StockinHdr H INner Join  Stockinitem D on H.StockINNo=D.StockINNo where H.StockInNo = " & SafeSQL(arrInvNo(i)) & " and approved =1 and (exported is NULL or Exported = 0) Order by H.StockInNo")
                    While dtr.Read
                        ExecuteNavSQL("Delete from VanStockRequest where StockInNo=" & SafeSQL(dtr("StockInNo")) & "and ItemNo = " & SafeSQL(dtr("ItemNo")))

                        sQry = "Insert into VanStockRequest (StockInNo, TransDate, AgentID, Location, ItemNo, UOM, Qty , Remarks, Reason) Values (" & SafeSQL(dtr("StockInNo").ToString) &
                            "," & SafeSQL(Format(dtr("TransDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("AgentID").ToString) & "," & SafeSQL(dtr("Location").ToString) & "," & SafeSQL(dtr("ItemNo").ToString) &
                            "," & SafeSQL(dtr("UOM").ToString) & "," & dtr("Qty").ToString & "," & SafeSQL(dtr("Remarks").ToString) & "," & SafeSQL(dtr("Reason").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export StockInItem - " & dtr("StockInNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteNavSQL(sQry)
                        ExecuteSQLAnother("Update StockInItem Set Exported = 1 Where StockInNo = " & SafeSQL(arrInvNo(i)) & " and [LineNo] = " & SafeSQL(dtr("LineNo").ToString))
                    End While
                    dtr.Close()

                Catch ex As Exception
                    dtr.Close()
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
        Dim dtr As SqlDataReader = Nothing
        Dim dtr1 As SqlDataReader = Nothing
        Dim sExtDocNo As String = ""
        Dim sSql As String
        Dim qry = ""
        Dim qry1 = ""
        Dim Ord As String = ""
        Dim Ord1 As String = ""
        Dim spricegroup As String = String.Empty
        Try
            dtr = ReadRecord("Select  OrdNo from OrderHdr where isnull(Exported,0) = 0 and IsNull(Void,0) = 0  Order by OrdNo ")
            While dtr.Read
                If arrInvNo.Contains(dtr("OrdNo").ToString) = False Then arrInvNo.Add(dtr("OrdNo").ToString)
            End While
            dtr.Close()
            dtr.Dispose()

            For i = 0 To arrInvNo.Count - 1
                Dim j = 0
                Try
                    Ord = arrInvNo(i)
                    qry = "SELECT OrderHdr.OrdNo, OrderHdr.OrdDt, OrderHdr.CustId, OrderHdr.AgentId, OrderHdr.SubTotal, OrderHdr.GstAmt, OrderHdr.TotalAmt, " &
                           " OrderHdr.PayTerms, OrderHdr.DeliveryDate, OrderHdr.MDTNo as SalesUnit, OrderHdr.Discount, OrderHdr.Gst, " &
                           " OrderHdr.Remarks, IsNull(OrderHdr.LineCount,0) as AllItemLine,OrderHdr.PONo, OrderHdr.DisPer,Customer.PriceGroup FROM  OrderHdr INNER JOIN Customer ON OrderHdr.CustId = Customer.CustNo " &
                            " WHERE (OrderHdr.OrdNo = " & SafeSQL(Ord) & ")  ORDER BY OrderHdr.OrdNo"
                    dtr = ReadRecord(qry)

                    'dtr = ReadRecord("Select OrdNo,OrdDt,CustId,AgentId,SubTotal,GstAmt,TotalAmt,PayTerms,DeliveryDate,SalesUnit,Discount,Gst,Remarks,AllItemLine from OrderHdr where OrdNo =  " & SafeSQL(arrInvNo(i)) & " Order by OrdNo")

                    While dtr.Read
                        j = j + 1
                        ExecuteNavSQL("Delete from OrderHdr where OrdNo=" & SafeSQL(dtr("OrdNo").ToString))
                        ExecuteNavSQL("Delete from OrdItem where OrdNo=" & SafeSQL(dtr("OrdNo").ToString))
                        spricegroup = dtr("Pricegroup").ToString()

                        sQry = "Insert into OrderHdr (OrdNo,OrdDt,CustId,AgentId,SubTotal,GstAmt,TotalAmt,PayTerms,DeliveryDate,SalesUnit,Discount,Gst,Remarks,AllItemLine,PONo,DiscountPer, CreatedDate) Values (" & SafeSQL(dtr("OrdNo").ToString) &
                            "," & SafeSQL(Format(dtr("OrdDt"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("CustId").ToString) & "," & SafeSQL(dtr("AgentId").ToString) & "," & SafeSQL(dtr("SubTotal").ToString) &
                            "," & SafeSQL(dtr("GstAmt").ToString) & "," & SafeSQL(dtr("TotalAmt").ToString) & "," & SafeSQL(dtr("PayTerms").ToString) & "," & SafeSQL(Format(dtr("DeliveryDate"), "yyyyMMdd")) &
                         "," & SafeSQL(dtr("SalesUnit").ToString) & "," & SafeSQL(dtr("Discount").ToString) & "," & SafeSQL(dtr("Gst").ToString) & "," & SafeSQL(dtr("Remarks").ToString) & "," & SafeSQL(dtr("AllItemLine").ToString) & "," & SafeSQL(dtr("PONo").ToString) & "," & SafeSQL(dtr("DisPer").ToString) & ", Format(SYSDATETIMEOFFSET() AT TIME ZONE 'SE Asia Standard Time','yyyy-MM-dd HH:mm:ss') )"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export OrderHdr - " & dtr("OrdNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteNavSQL(sQry)

                    End While
                    dtr.Close()

                Catch ex As Exception
                    dtr.Close()
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export OrderHdr - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try

                Try
                    If j > 0 Then
                        Ord1 = arrInvNo(i)
                        'Itempr.pricegroup removed by jagadish on 29-10-2020' 

                        ' dtr1 = ReadRecord(" Select distinct OI.OrdNo,OI.ItemNo,OI.UOM,OI.Qty,Round((NULLIf(UBK.BaseQty, 0) * OI.Price),2) as Price,CASE WHEN OI.DisPer<>0 then Round((Qty * Price) * (OI.DisPer / 100),2) else 0 End as Discount,OI.SubAmt,OI.SalesType,OI.[LineNo],ItemPr.PriceGroup,OI.AttachedLineNo,OI.DisPer, " & _
                        '" Item.BulkUOM,Cast(qty as Int)/NULLIf(Cast(UBK.BaseQty as Int),0)  as BulkQty,LooseUOM,(Cast(qty as Int) % NULLIf(Cast(UBK.BaseQty as Int),0))/NULLIf(ULS.BaseQty,0) as LooseQty," & _
                        '" Round((Qty * Price),2) as SubAmtBefDis,'' as ProductPriceLevel,ReasonCode,OI.PromoOffer From OrdItem OI Inner Join Item On OI.ItemNo = Item.ItemNo Inner Join ItemPr On OI.ItemNo = ItemPr.ItemNo Left Join UOM UBK " & _
                        '" On Item.ItemNo = UBK.ItemNo and UBK.UOM=Item.BulkUOM Left Join UOM ULS On Item.ItemNo = ULS.ItemNo and ULS.UOM=Item.LooseUOM  Where OI.OrdNo =  " & SafeSQL(arrInvNo(i)) & " Order by OI.OrdNo")

                        'qry1 = " Select distinct OI.OrdNo,OI.ItemNo,OI.UOM,OI.Qty,Round((NULLIf(UBK.BaseQty, 0) * OI.Price),2) as Price,CASE WHEN OI.DisPer<>0 then Round((Qty * Price) * (OI.DisPer / 100),2) else 0 End as Discount,OI.SubAmt,OI.SalesType,OI.[LineNo], IsnUll(OI.AttachedLineNo,0) as AttachedLineNo,OI.DisPer, " &
                        '    " Item.BulkUOM,Cast(qty as Int)/NULLIf(Cast(UBK.BaseQty as Int),0)  as BulkQty,LooseUOM,(Cast(qty as Int) % NULLIf(Cast(UBK.BaseQty as Int),0))/NULLIf(ULS.BaseQty,0) as LooseQty," &
                        '    " Round((Qty * Price),2) as SubAmtBefDis,'' as ProductPriceLevel,ReasonCode,OI.PromoOffer From OrdItem OI Inner Join Item On OI.ItemNo = Item.ItemNo Left Join UOM UBK " &
                        '    " On Item.ItemNo = UBK.ItemNo and UBK.UOM=Item.BulkUOM Left Join UOM ULS On Item.ItemNo = ULS.ItemNo and ULS.UOM=Item.LooseUOM  Where OI.OrdNo =  " & SafeSQL(Ord1) & " Order by OI.OrdNo"

                        qry1 = " Select distinct OI.OrdNo,OI.ItemNo,OI.UOM,OI.Qty,ISNUll(Round((NULLIf(UBK.BaseQty, 0) * OI.Price),2),0) as Price,CASE WHEN OI.DisPer<>0 then Round((Qty * Price) * (OI.DisPer / 100),2) else 0 End as Discount,OI.SubAmt,ISNUll(OI.SalesType,'') as SalesType,OI.[LineNo], IsnUll(OI.AttachedLineNo,0) as AttachedLineNo,OI.DisPer, " &
                            " Item.BulkUOM,Cast(qty as Int)/NULLIf(Cast(UBK.BaseQty as Int),0)  as BulkQty,LooseUOM,(Cast(qty as Int) % NULLIf(Cast(UBK.BaseQty as Int),0))/NULLIf(ULS.BaseQty,0) as LooseQty," &
                            " ISNUll(Round((Qty * Price),2),0) as SubAmtBefDis,'' as ProductPriceLevel,ReasonCode,ISNULL(OI.PromoOffer,'')as PromoOffer From OrdItem OI Inner Join Item On OI.ItemNo = Item.ItemNo Left Join UOM UBK " &
                            " On Item.ItemNo = UBK.ItemNo and UBK.UOM=Item.BulkUOM Left Join UOM ULS On Item.ItemNo = ULS.ItemNo and ULS.UOM=Item.LooseUOM  Where OI.OrdNo =  " & SafeSQL(Ord1) & " Order by OI.OrdNo"

                        dtr1 = ReadRecord(qry1)
                        Dim k = 0
                        While dtr1.Read
                            k = k + 1
                            Dim sPromooffer As String = ""
                            Dim attachedline As String = "0"
                            'If dtr1("SalesType") = "F" And dtr1("PromoOffer") <> "0" Then
                            If (dtr1("SalesType") = "F") And dtr1("PromoOffer") <> "0" Then

                                Dim rs As SqlDataReader
                                rs = ReadRecordAnother(" Select PromoOffer from OrdItem OI  Where OI.[LineNo]=" & dtr1("LineNo") & " and OI.OrdNo =  " & SafeSQL(arrInvNo(i)) & " Order by OI.OrdNo")
                                If rs.Read = True Then
                                    sPromooffer = rs("PromoOffer").ToString
                                End If
                                rs.Close()
                                rs.Dispose()

                                rs = ReadRecordAnother(" Select [LineNo] as LineNum from OrdItem OI  Where OI.PromoID=" & SafeSQL(sPromooffer) & " and OI.OrdNo =  " & SafeSQL(arrInvNo(i)) & " Order by OI.OrdNo")
                                If rs.Read = True Then
                                    attachedline = rs("LineNum")
                                End If
                                rs.Close()
                                rs.Dispose()
                            End If
                            sSql = "Insert into OrdItem (OrdNo,ItemNo,UOM,Qty,Price,Discount,SubAmt,SalesType,[LineNo],PriceGroup,AttachedLineNo,DisPer,BUlkUOM,BulkQty,LooseUOM,LooseQty,SubAmtBefDis,ProductPriceLevel,ReasonCode) Values (" & SafeSQL(dtr1("OrdNo").ToString) &
                                 "," & SafeSQL(dtr1("ItemNo").ToString) & "," & SafeSQL(dtr1("UOM").ToString) & "," & SafeSQL(dtr1("Qty").ToString) & "," & SafeSQL(dtr1("Price").ToString) &
                                 "," & SafeSQL(dtr1("Discount").ToString) & "," & SafeSQL(dtr1("SubAmt").ToString) & "," & SafeSQL(dtr1("SalesType").ToString) & "," & SafeSQL(dtr1("LineNo")) & "," & SafeSQL(spricegroup) & "," & SafeSQL(attachedline) & "," & SafeSQL(dtr1("DisPer").ToString) &
                                 "," & SafeSQL(dtr1("BulkUOM").ToString) & "," & dtr1("BulkQty") & "," & SafeSQL(dtr1("LooseUOM").ToString) & "," & dtr1("LooseQty") & "," & dtr1("SubAmtBefDis") & "," & SafeSQL(dtr1("ProductPriceLevel").ToString) & "," & SafeSQL(dtr1("ReasonCOde").ToString) & ")"

                            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export OrdItem - " & SafeSQL(arrInvNo(i))) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sSql) & ")")

                            ExecuteNavSQL(sSql)

                        End While
                        dtr1.Close()
                        If k > 0 Then
                            ExecuteSQLAnother("Update OrderHdr Set Exported = 1 Where OrdNo = " & SafeSQL(Ord1))
                        Else
                            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export OrdItem'," & SafeSQL(NavCompanyName) & "," & SafeSQL("Order item not found") & ")")
                        End If
                    End If
                Catch ex As Exception
                    dtr1.Close()
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export OrdItem'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")

                    ExecuteNavSQL("Delete from OrderHdr where OrdNo=" & SafeSQL(dtr("OrdNo").ToString))
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Deleting exported Sales Order Header'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")

                    ExecuteNavSQL("Delete from OrdItem where OrdNo=" & SafeSQL(dtr("OrdNo").ToString))
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Deleting exported Sales Order Item'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try
            Next
        Catch ex As Exception
            dtr.Close()
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export OrderHdr'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")

            ExecuteNavSQL("Delete from OrderHdr where OrdNo=" & SafeSQL(dtr("OrdNo").ToString))
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Deleting exported Sales Order Header'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")

            ExecuteNavSQL("Delete from OrdItem where OrdNo=" & SafeSQL(dtr("OrdNo").ToString))
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Deleting exported Sales Order Item'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try
    End Sub
    Private Sub ExportCreditMemo()
        Dim sCurCode As String = ""
        Dim arrInvNo = New ArrayList
        Dim sQry As String = ""
        Dim dTransport As Double = 0
        sCurCode = "" 'dtr1("LCY Code")
        Dim dExRate As Double = 0
        Dim dtr As SqlDataReader = Nothing
        Dim dtr1 As SqlDataReader = Nothing
        Dim sExtDocNo As String = ""
        Dim sSql As String
        Try
            dtr = ReadRecord("Select  CreditNoteNo from CreditNote where isnull(Exported,0) = 0 and IsNull(Approved,0) = 1 Order by CreditNoteNo")
            While dtr.Read
                If arrInvNo.Contains(dtr("CreditNoteNo").ToString) = False Then arrInvNo.Add(dtr("CreditNoteNo").ToString)
            End While
            dtr.Close()
            dtr.Dispose()

            For i = 0 To arrInvNo.Count - 1
                Dim j = 0
                Try
                    dtr = ReadRecord("Select CN.CreditNoteNo,CN.CreditDate,CN.CustNo,CN.GoodsReturnNo,CN.SalesPersonCode,CN.SubTotal,CN.Gst,CN.TotalAmt,CN.PaidAmt,Customer.PaymentTerms as PayTerms, CN.Void,CN.MDTNo as SalesUnit from CreditNote CN " &
                                    " Inner Join Customer On CN.CustNo = Customer.CustNo where CN.CreditNoteNo = " & SafeSQL(arrInvNo(i)) & " Order by CN.CreditNoteNo")

                    While dtr.Read
                        j = j + 1
                        ExecuteNavSQL("Delete from CreditNote where CreditNoteNo=" & SafeSQL(dtr("CreditNoteNo").ToString))
                        ExecuteNavSQL("Delete from CreditNoteDet where CreditNoteNo=" & SafeSQL(dtr("CreditNoteNo").ToString))

                        sQry = "Insert into CreditNote (CreditNoteNo,CreditDate,CustNo,GoodsReturnNo,SalesPersonCode,SubTotal,Gst,TotalAmt,PaidAmt,Payterms,Void,SalesUnit) Values (" & SafeSQL(dtr("CreditNoteNo").ToString) &
                            "," & SafeSQL(Format(dtr("CreditDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("CustNo").ToString) & "," & SafeSQL(dtr("GoodsReturnNo").ToString) & "," & SafeSQL(dtr("SalesPersonCode").ToString) &
                            "," & SafeSQL(dtr("SubTotal").ToString) & "," & SafeSQL(dtr("Gst").ToString) & "," & SafeSQL(dtr("TotalAmt").ToString) & "," & SafeSQL(dtr("PaidAmt").ToString) & "," & SafeSQL(dtr("Payterms").ToString) & "," & SafeSQL(dtr("Void").ToString) & "," & SafeSQL(dtr("SalesUnit").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export CreditNote - " & dtr("CreditNoteNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteNavSQL(sQry)

                    End While
                    dtr.Close()

                Catch ex As Exception
                    dtr.Close()
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export CreditNote - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try

                Try
                    If j > 0 Then

                        dtr1 = ReadRecord("Select * From CreditNoteDet  Where CreditNoteNo =  " & SafeSQL(arrInvNo(i)) & " Order by CreditNoteNo")
                        Dim k = 0
                        While dtr1.Read
                            k = k + 1
                            sSql = "Insert into CreditNoteDet (CreditNoteNo,ItemNo,UOM,BaseUOM,Price,Qty,Amt,[LineNo],DisPer,AttachedLineNo,SalesType) Values (" & SafeSQL(dtr1("CreditNoteNo").ToString) &
                                 "," & SafeSQL(dtr1("ItemNo").ToString) & "," & SafeSQL(dtr1("UOM").ToString) & "," & SafeSQL(dtr1("BaseUOM").ToString) & "," & SafeSQL(dtr1("Price").ToString) &
                                 "," & SafeSQL(dtr1("Qty").ToString) & "," & SafeSQL(dtr1("Amt").ToString) & "," & SafeSQL(dtr1("LineNo").ToString) & "," & SafeSQL(dtr1("DisPer").ToString) & "," & SafeSQL(dtr1("AttachedLineNo").ToString) & "," & SafeSQL(dtr1("SalesType").ToString) & ")"

                            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export CreditNoteDet - " & SafeSQL(arrInvNo(i))) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sSql) & ")")

                            ExecuteNavSQL(sSql)

                        End While
                        dtr1.Close()
                        If k > 0 Then
                            ExecuteSQLAnother("Update CreditNote Set Exported = 1 , Approved =1 Where CreditNoteNo = " & SafeSQL(arrInvNo(i)))
                        Else
                            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export CreditNoteDet'," & SafeSQL(NavCompanyName) & "," & SafeSQL("Creditnote item not found") & ")")
                        End If
                    End If
                Catch ex As Exception
                    dtr1.Close()
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export CreditNoteDet'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                    System.IO.File.AppendAllText(Application.StartupPath & "\ErrorLog.txt", "Export CreditNoteDet  Error" & "   " & Date.Now.ToString & " " & ex.Message & vbCrLf)
                End Try
            Next
        Catch ex As Exception
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
        Dim dtr As SqlDataReader = Nothing
        Dim dtr1 As SqlDataReader = Nothing
        Dim sExtDocNo As String = ""
        Dim sSql As String
        Try
            dtr = ReadRecord("Select  RcptNo from Receipt where isnull(Exported,0) = 0 and IsNull(Void,0) = 0 Order by RcptNo")
            While dtr.Read
                If arrInvNo.Contains(dtr("RcptNo").ToString) = False Then arrInvNo.Add(dtr("RcptNo").ToString)
            End While
            dtr.Close()
            dtr.Dispose()


            For i = 0 To arrInvNo.Count - 1
                Dim j = 0
                Try

                    dtr = ReadRecord("Select RcptNo,RcptDt,CustId,AgentId,PayMethod,ChqNo,Case when PayMethod = 'CASH' then NULL Else ChqDt End as ChqDt,Amount,Void,DTG,BankName,MDTNo From Receipt where RcptNo =  " & SafeSQL(arrInvNo(i)) & " Order by RcptNo")

                    While dtr.Read
                        j = j + 1
                        ExecuteNavSQL("Delete from Receipt where RcptNo=" & SafeSQL(dtr("RcptNo").ToString))
                        ExecuteNavSQL("Delete from RcptItem where RcptNo=" & SafeSQL(dtr("RcptNo").ToString))

                        If IsDBNull(dtr("ChqDt")) = True Then
                            sQry = "Insert into Receipt (RcptNo,RcptDt,CustId,AgentId,PayMethod,ChqNo,Amount,Void,DTG,BankName,SalesUnit) Values (" & SafeSQL(dtr("RcptNo").ToString) &
                            "," & SafeSQL(Format(dtr("RcptDt"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("CustId").ToString) & "," & SafeSQL(dtr("AgentId").ToString) & "," & SafeSQL(dtr("PayMethod").ToString) &
                            "," & SafeSQL(dtr("ChqNo").ToString) & "," & SafeSQL(dtr("Amount").ToString) & "," & SafeSQL(dtr("Void").ToString) & "," & SafeSQL(Format(dtr("DTG"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("BankName").ToString) & "," & SafeSQL(dtr("MDTNo").ToString) & ")"
                        Else
                            sQry = "Insert into Receipt (RcptNo,RcptDt,CustId,AgentId,PayMethod,ChqNo,ChqDt,Amount,Void,DTG,BankName,SalesUnit) Values (" & SafeSQL(dtr("RcptNo").ToString) &
                            "," & SafeSQL(Format(dtr("RcptDt"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("CustId").ToString) & "," & SafeSQL(dtr("AgentId").ToString) & "," & SafeSQL(dtr("PayMethod").ToString) &
                            "," & SafeSQL(dtr("ChqNo").ToString) & "," & SafeSQL(Format(dtr("ChqDt"), "yyyyMMdd")) & "," & SafeSQL(dtr("Amount").ToString) & "," & SafeSQL(dtr("Void").ToString) & "," & SafeSQL(Format(dtr("DTG"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("BankName").ToString) & "," & SafeSQL(dtr("MDTNo").ToString) & ")"
                        End If


                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export Receipt - " & dtr("RcptNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteNavSQL(sQry)

                    End While
                    dtr.Close()

                Catch ex As Exception
                    dtr.Close()
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export Receipt - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try

                Try
                    If j > 0 Then

                        dtr1 = ReadRecord("Select * From RcptItem  Where RcptNo =  " & SafeSQL(arrInvNo(i)) & " Order by RcptNo")
                        Dim k = 0
                        While dtr1.Read
                            k = k + 1
                            sSql = "Insert into RcptItem (RcptNo,InvNo,AmtPaid) Values (" & SafeSQL(dtr1("RcptNo").ToString) &
                                 "," & SafeSQL(dtr1("InvNo").ToString) & "," & SafeSQL(dtr1("AmtPaid").ToString) & ")"

                            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export RcptItem - " & SafeSQL(arrInvNo(i))) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sSql) & ")")

                            ExecuteNavSQL(sSql)
                        End While
                        dtr1.Close()
                        If k > 0 Then
                            ExecuteSQLAnother("Update Receipt Set Exported = 1 Where RcptNo = " & SafeSQL(arrInvNo(i)))
                        Else
                            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export RcptItem'," & SafeSQL(NavCompanyName) & "," & SafeSQL("Receipt item not found..") & ")")
                        End If
                    End If
                Catch ex As Exception
                    dtr1.Close()
                    System.IO.File.AppendAllText(Application.StartupPath & "\ErrorLog.txt", "Export Inv Item Error" & "   " & Date.Now.ToString & " " & ex.Message & vbCrLf)
                End Try
            Next
        Catch ex As Exception
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
        Dim dtr As SqlDataReader = Nothing
        Dim dtr1 As SqlDataReader = Nothing
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
                Dim j = 0
                Try
                    dtr = ReadRecord("Select * from StockOrder where StockNo =  " & SafeSQL(arrInvNo(i)) & " Order by StockNo")
                    While dtr.Read
                        j = j + 1
                        ExecuteNavSQL("Delete from StockOrder where StockNo=" & SafeSQL(dtr("StockNo").ToString))
                        ExecuteNavSQL("Delete from StockOrderItem where StockNo=" & SafeSQL(dtr("StockNo").ToString))

                        sQry = "Insert into StockOrder (StockNo,OrdDt,TrnDate,Location,AgentId,Remarks,SalesUnit) Values (" & SafeSQL(dtr("StockNo").ToString) &
                            "," & SafeSQL(Format(dtr("OrdDt"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(Format(dtr("TrnDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("Location").ToString) & "," & SafeSQL(dtr("AgentId").ToString) &
                            "," & SafeSQL(dtr("Remarks").ToString) & "," & SafeSQL(dtr("MDTNo").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export StockOrder - " & dtr("StockNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteNavSQL(sQry)

                    End While
                    dtr.Close()

                Catch ex As Exception
                    dtr.Close()
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export StockOrder - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try

                Try
                    If j > 0 Then
                        dtr1 = ReadRecord("Select * From StockOrderItem  Where StockNo =  " & SafeSQL(arrInvNo(i)) & " Order by StockNo")
                        Dim k = 0
                        While dtr1.Read
                            k = k + 1
                            sSql = "Insert into StockOrderItem (StockNo,ItemNo,UOM,Qty,Location,[LineNo]) Values (" & SafeSQL(dtr1("StockNo").ToString) &
                                 "," & SafeSQL(dtr1("ItemNo").ToString) & "," & SafeSQL(dtr1("UOM").ToString) & "," & SafeSQL(dtr1("Qty").ToString) & "," & SafeSQL(dtr1("Location").ToString) &
                                 "," & SafeSQL(dtr1("LineNo").ToString) & ")"

                            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export StockOrderItem - " & SafeSQL(arrInvNo(i))) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sSql) & ")")


                            ExecuteNavSQL(sSql)
                        End While
                        dtr1.Close()
                        If k > 0 Then
                            ExecuteSQLAnother("Update StockOrder Set Exported = 1 Where StockNo = " & SafeSQL(arrInvNo(i)))
                        Else
                            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export StockOrderItem'," & SafeSQL(NavCompanyName) & "," & SafeSQL("Stock order item not found") & ")")
                        End If
                    End If
                Catch ex As Exception
                    dtr1.Close()
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export StockOrderItem'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try
            Next
        Catch ex As Exception
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
        Dim dtr As SqlDataReader = Nothing
        Dim dtr1 As SqlDataReader = Nothing
        Dim sExtDocNo As String = ""
        Dim sSql As String
        Try
            dtr = ReadRecord("Select Distinct GOodsReturn.ReturnNo from GOodsReturn Inner Join GOodsReturnItem on GOodsReturn.ReturnNo = GOodsReturnItem.ReturnNo" &
                    " where isnull(GOodsReturn.Exported,0) = 0 and GOodsReturnItem.ItemNo in (Select ItemNO from GOodsReturnItem Where IsNull(Approved,0) =1 and returnno = GOodsReturn.returnno  )  Order by GOodsReturn.ReturnNo")
            While dtr.Read
                If arrInvNo.Contains(dtr("ReturnNo").ToString) = False Then arrInvNo.Add(dtr("ReturnNo").ToString)
            End While
            dtr.Close()
            dtr.Dispose()

            For i = 0 To arrInvNo.Count - 1
                Try
                    dtr = ReadRecord("Select GR.ReturnNo,GR.ReturnDate,GR.CustNo,GR.SalesPersonCOde,Sum(GI.Amt) as SubTotal,7 as GST,(Sum(GI.Amt)* 7 / 100) as GSTAmt,Sum(GI.Amt)  + (Sum(GI.Amt)* 7 / 100) " &
                            " as TotalAmt,GR.Void,GR.VoidDate,IsNull(GR.Exported,0) as Exported , GR.IsConfirmed,GR.ConfirmedBy,GR.MDTNo from GoodsReturn GR Inner Join GoodsReturnItem GI on GR.ReturnNo = GI.ReturnNo " &
                            " where isnull(GR.Exported,0) = 0 and GI.ItemNo in (Select ItemNO from GOodsReturnItem Where IsNull(Approved,0) =1 and GR.ReturnNo = " & SafeSQL(arrInvNo(i)) & " ) " &
                            " Group By GR.ReturnNo,GR.ReturnDate,GR.CustNo,GR.SalesPersonCOde,GR.Void,GR.VoidDate,GR.Exported ,GR.IsConfirmed,GR.ConfirmedBy Order by GR.ReturnNo")
                    While dtr.Read
                        ExecuteNavSQL("Delete from [GoodsReturn] where ReturnNo =" & SafeSQL(arrInvNo(i)))
                        ExecuteNavSQL("Delete from GoodsReturnItem where ReturnNo=" & SafeSQL(arrInvNo(i)))
                        ExecuteNavSQL("Delete from GoodsReturnExpDet where ReturnNo=" & SafeSQL(arrInvNo(i)))

                        sQry = "Insert into [GoodsReturn] (ReturnNo,ReturnDate,CustNo,SalesPersonCode,SubTotal,GST,TotalAmt,Void,VoidDate,Exported,Isconfirmed,ConfirmedBy,SalesUnit) Values (" & SafeSQL(dtr("ReturnNo").ToString) &
                             "," & SafeSQL(Format(dtr("ReturnDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("CustNo").ToString) & "," & SafeSQL(dtr("SalesPersonCode").ToString) &
                             "," & SafeSQL(dtr("SubTotal").ToString) & "," & SafeSQL(dtr("GSTAmt").ToString) & "," & SafeSQL(dtr("TotalAmt").ToString) &
                              "," & SafeSQL(dtr("Void").ToString) & "," & SafeSQL(Format(dtr("VoidDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("Exported").ToString) &
                               "," & SafeSQL(dtr("IsConfirmed").ToString) & "," & SafeSQL(dtr("ConfirmedBy").ToString) & "," & SafeSQL(dtr("MDTNo").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export GoodsReturn - " & dtr("ReturnNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteNavSQL(sQry)

                    End While
                    dtr.Close()

                Catch ex As Exception
                    dtr.Close()
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export GoodsReturn - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try

                Try
                    dtr1 = ReadRecord("Select Distinct GoodsReturnItem.* from GOodsReturn Inner Join GOodsReturnItem on GOodsReturn.ReturnNo = GOodsReturnItem.ReturnNo " &
                            " where isnull(GOodsReturn.Exported,0) = 0 and GOodsReturnItem.ItemNo in (Select ItemNO from GOodsReturnItem Where IsNull(GoodsReturnItem.Approved,0) =1  and GoodsReturnItem.ReturnNo = " & SafeSQL(arrInvNo(i)) & ") Order by GoodsReturnItem.ReturnNo")

                    While dtr1.Read
                        sSql = "Insert into GoodsReturnItem (ReturnNo,ItemNo,UOM,Quantity,[LineNo],Price,Amt,ReasonCode,Remarks) Values (" & SafeSQL(dtr1("ReturnNo").ToString) &
                             "," & SafeSQL(dtr1("ItemNo").ToString) & "," & SafeSQL(dtr1("UOM").ToString) & "," & SafeSQL(dtr1("Quantity").ToString) & "," & SafeSQL(dtr1("LineNo").ToString) &
                             "," & SafeSQL(dtr1("Price").ToString) & "," & SafeSQL(dtr1("Amt").ToString) & "," & SafeSQL(dtr1("ReasonCode").ToString) & "," & SafeSQL(dtr1("Remarks").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export GoodsReturnDet - " & SafeSQL(arrInvNo(i))) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sSql) & ")")

                        ExecuteNavSQL(sSql)
                    End While
                    dtr1.Close()
                    dtr1 = ReadRecord("Select Distinct GoodsReturnExpDet.* from GOodsReturn Inner Join GOodsReturnItem on GOodsReturn.ReturnNo = GOodsReturnItem.ReturnNo " &
                        " Inner Join GoodsReturnExpDet on GoodsReturnItem.ReturnNo = GoodsReturnExpDet.ReturnNo where isnull(GOodsReturn.Exported,0) = 0 " &
                        " and GoodsReturnExpDet.ItemNo in (Select ItemNO from GOodsReturnItem Where IsNull(GoodsReturnItem.Approved,0) =1  " &
                        " and GoodsReturnItem.ReturnNo = " & SafeSQL(arrInvNo(i)) & ") Order by GoodsReturnExpDet.ReturnNo")
                    While dtr1.Read
                        sSql = "Insert into GoodsReturnExpDet (ReturnNo,ItemNo,UOM,Qty,[LineNo],LotNo,ExpiryDate) Values (" & SafeSQL(dtr1("ReturnNo").ToString) &
                             "," & SafeSQL(dtr1("ItemNo").ToString) & "," & SafeSQL(dtr1("UOM").ToString) & "," & SafeSQL(dtr1("Qty").ToString) & "," & SafeSQL(dtr1("LineNo").ToString) &
                             "," & SafeSQL(dtr1("LotNo").ToString) & "," & SafeSQL(Format(dtr1("ExpiryDate"), "yyyyMMdd HH:mm:ss")) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export GoodsReturnExpDet - " & SafeSQL(arrInvNo(i))) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sSql) & ")")

                        ExecuteNavSQL(sSql)
                    End While
                    dtr1.Close()
                    ExecuteSQLAnother("Update [GoodsReturn] Set Exported = 1 Where ReturnNo = " & SafeSQL(arrInvNo(i)))
                Catch ex As Exception
                    dtr1.Close()
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export GoodsReturnExpDet'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try
            Next
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export GoodsReturn'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try
    End Sub
    Private Sub ExportNewCustomer()
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
            dtr = ReadRecord("Select CustID from NewCust where isnull(Exported,0) = 0 ")
            While dtr.Read
                If arrInvNo.Contains(dtr("CustId").ToString) = False Then arrInvNo.Add(dtr("CustId").ToString)
            End While
            dtr.Close()
            dtr.Dispose()

            For i = 0 To arrInvNo.Count - 1
                ExecuteNavSQL("Delete from NewCustomer where CustId =" & SafeSQL(arrInvNo(i)))
                Try
                    dtr = ReadRecord("Select [CustID],[CustName],[Address],[Pin],[Email],[Longitude],[Latitude],[MobileNo],[Province],[Soi],[Road],[District],[SubDistrict],[BillAdd1],[BillAdd2],[AgentID],[Active],[PaymentMethod],[TinNo],[Discount],[BusinessType],[CustType],[LocationType],[Branch],[Outlettype],[InstitutionType],[ProdList],[Remarks],[MDTNo],[EditDate],[CreditLimit],[SearchName],[Address2],[PaymentTerms],[Place] From NewCust where CustId =  " & SafeSQL(arrInvNo(i)))

                    While dtr.Read

                        sQry = "Insert Into NewCustomer ([CustID],[CustName],[ChineseName],[Address],[PostCode],[Email],[Longitude],[Latitude],[Phone],[Province],[Soi],[Road],[District],[SubDistrict],[DeliveryAddress],[BillAdd1],[AgentID],[Active],[PaymentMethod],[GSTNO],[DiscountGroup],[Buss_Type],[CustType],[LocationType],[Branch],[StoreType],[Institution_Type],[Product_List],[Remarks],[Area], CreatedDate, EditDate, CreditLimit, SearchName, Address2, PaymentTerms,Exported,Place ) Values (" &
                            SafeSQL(dtr("CustID").ToString) & "," & SafeSQL(dtr("CustName").ToString) & "," & SafeSQL(dtr("CustName").ToString) & "," & SafeSQL(dtr("Address").ToString) & "," & SafeSQL(dtr("Pin").ToString) & "," & SafeSQL(dtr("Email").ToString) & "," & SafeSQL(dtr("Longitude").ToString) & "," & SafeSQL(dtr("Latitude").ToString) & "," & SafeSQL(dtr("MobileNo").ToString) & "," & SafeSQL(dtr("Province").ToString) & "," & SafeSQL(dtr("Soi").ToString) &
                            "," & SafeSQL(dtr("Road").ToString) & "," & SafeSQL(dtr("District").ToString) & "," & SafeSQL(dtr("SubDistrict").ToString) & "," & SafeSQL(dtr("BillAdd1").ToString) & "," & SafeSQL(dtr("BillAdd2").ToString) & "," & SafeSQL(dtr("AgentID").ToString) & "," & SafeSQL(dtr("Active").ToString) & "," & SafeSQL(dtr("PaymentMethod").ToString) & "," & SafeSQL(dtr("TinNo").ToString) & "," & SafeSQL(dtr("Discount").ToString) &
                            "," & SafeSQL(dtr("BusinessType").ToString) & "," & SafeSQL(dtr("CustType").ToString) & "," & SafeSQL(dtr("LocationType").ToString) & "," & SafeSQL(dtr("Branch").ToString) & "," & SafeSQL(dtr("Outlettype").ToString) & "," & SafeSQL(dtr("InstitutionType").ToString) & "," & SafeSQL(dtr("ProdList").ToString) & "," & SafeSQL(dtr("Remarks").ToString) & "," & SafeSQL(dtr("MDTNo").ToString) & ", Format(SYSDATETIMEOFFSET() AT TIME ZONE 'SE Asia Standard Time','yyyy-MM-dd HH:mm:ss'), " &
                            SafeSQL(Format(dtr("EditDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("CreditLimit").ToString) & "," & SafeSQL(dtr("SearchName").ToString) & "," & SafeSQL(dtr("Address2").ToString) & "," & SafeSQL(dtr("PaymentTerms").ToString) & ",1," & SafeSQL(dtr("Place").ToString) & ")"


                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export NewCustomer - " & dtr("CustId").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteNavSQL(sQry)
                    End While
                    dtr.Close()
                    ExecuteSQLAnother("Update NewCust set exported =1 where CustId =" & SafeSQL(arrInvNo(i)))
                Catch ex As Exception
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export NewCustomer - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try

                ExecuteNavSQL("Delete from Contacts where CustNo =" & SafeSQL(arrInvNo(i)))
                Try
                    dtr = ReadRecord("Select CustNo,Code,IsNull(Designation,'') as Designation, Name,PhoneNo,DTG from Contacts where CustNo =  " & SafeSQL(arrInvNo(i)))
                    While dtr.Read
                        If dtr("Designation") <> "" Then
                            sQry = "Insert Into Contacts (CustNo,Code,Designation,Name,PhoneNo,DTG  ) Values (" &
                            SafeSQL(dtr("CustNo").ToString) & "," & SafeSQL(dtr("Code").ToString) & "," & SafeSQL(dtr("Designation").ToString) & "," & SafeSQL(dtr("Name").ToString) & "," & SafeSQL(dtr("PhoneNo").ToString) & ",GetDate())"
                            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export Contacts - " & dtr("CustNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")
                            ExecuteNavSQL(sQry)
                        End If
                    End While
                    dtr.Close()
                    ExecuteSQLAnother("Update Contacts set exported =1 where CustNo =" & SafeSQL(arrInvNo(i)))
                Catch ex As Exception
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export Contacts - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try
            Next
        Catch ex As Exception
            dtr.Close()
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export NewCustomer'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try
    End Sub
    Private Sub ExportItemTrans()
        Dim sCurCode As String
        Dim arrInvNo = New ArrayList
        Dim sQry As String = ""
        Dim dTransport As Double = 0
        sCurCode = "" 'dtr1("LCY Code")
        Dim dExRate As Double = 0
        Dim dtr As SqlDataReader = Nothing
        Dim dtr1 As SqlDataReader = Nothing
        Dim sExtDocNo As String = ""
        Dim sSql As String
        Try
            dtr = ReadRecord("Select distinct  DocNo from ItemTrans where isnull(Exported,0) = 0 and doctype in ('GIN','GOUT','GVAR') Order by DocNo")
            While dtr.Read
                If arrInvNo.Contains(dtr("DocNo").ToString) = False Then arrInvNo.Add(dtr("DocNo").ToString)
            End While
            dtr.Close()
            dtr.Dispose()

            For i = 0 To arrInvNo.Count - 1
                ExecuteNavSQL("Delete from ItemTrans where DocNo =" & SafeSQL(arrInvNo(i)))
                Try
                    dtr = ReadRecord("Select * from ItemTrans where DocNo =  " & SafeSQL(arrInvNo(i)) & " and doctype in ('GIN','GOUT','GVAR') Order by DocNo")

                    While dtr.Read
                        sQry = "Insert into ItemTrans (DocNo,DocDt,DocType,Location,ItemId,UOM,Qty,Exported,Remarks,IsUpdated) Values (" & SafeSQL(dtr("DocNo").ToString) &
                            "," & SafeSQL(Format(dtr("DocDt"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("DocType").ToString) & "," & SafeSQL(dtr("Location").ToString) & "," & SafeSQL(dtr("ItemId").ToString) &
                             "," & SafeSQL(dtr("UOM").ToString) & "," & SafeSQL(dtr("Qty").ToString) & "," & SafeSQL(dtr("Exported").ToString) & "," & SafeSQL(dtr("Remarks").ToString) &
                              "," & SafeSQL(dtr("IsUpdated").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export ItemTrans - " & dtr("DocNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteNavSQL(sQry)
                    End While
                    dtr.Close()
                    ExecuteSQLAnother("Update ItemTrans set exported =1 where DocNo =" & SafeSQL(arrInvNo(i)))
                Catch ex As Exception
                    dtr.Close()
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export ItemTrans - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try

            Next
        Catch ex As Exception
            dtr.Close()
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export ItemTrans'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try
    End Sub
    Private Sub ExportCustVisit()
        Dim sCurCode As String
        Dim arrInvNo = New ArrayList
        Dim sQry As String = ""
        Dim dTransport As Double = 0
        sCurCode = "" 'dtr1("LCY Code")
        Dim dExRate As Double = 0
        Dim dtr As SqlDataReader = Nothing
        Dim dtr1 As SqlDataReader = Nothing
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

                    ExecuteNavSQL("Delete from CustVisit where CustId =" & SafeSQL(arrInvNo(i)))

                    While dtr.Read

                        sQry = "Insert into CustVisit (CustId,TransNo,TransType,TransDate,AgentId,Status,Latitude,Longitude,Remarks,DTG,SalesUnit) Values (" & SafeSQL(dtr("CustId").ToString) &
                            "," & SafeSQL(dtr("TransNo").ToString) & "," & SafeSQL(dtr("TransType").ToString) & "," & SafeSQL(Format(dtr("TransDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("AgentId").ToString) &
                        "," & SafeSQL(dtr("Status").ToString) & "," & SafeSQL(dtr("Latitude").ToString) & "," & SafeSQL(dtr("Longitude").ToString) & "," & SafeSQL(dtr("Remarks").ToString) & ", GetDate(), " & SafeSQL(dtr("MDTNo").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export CustVisit - " & dtr("CustId").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteNavSQL(sQry)

                    End While
                    dtr.Close()
                    ExecuteSQLAnother("Update CustVisit Set Exported = 1 Where CustId = " & SafeSQL(arrInvNo(i)))
                Catch ex As Exception
                    dtr.Close()
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
        Dim dtr As SqlDataReader = Nothing
        'Dim dtr1 As SqlDataReader
        Dim sExtDocNo As String = ""
        Dim sSql As String
        Try
            dtr = ReadRecord("Select  DocNo from Exception where isnull(Exported,0) = 0 Order by DocNo")
            While dtr.Read
                If arrInvNo.Contains(dtr("DocNo").ToString) = False Then arrInvNo.Add(dtr("DocNo").ToString)
            End While
            dtr.Close()
            dtr.Dispose()


            For i = 0 To arrInvNo.Count - 1
                Try
                    dtr = ReadRecord("Select * from Exception where DocNo =  " & SafeSQL(arrInvNo(i)) & " Order by DocNo")

                    ExecuteNavSQL("Delete from Exception where DocNo =" & SafeSQL(arrInvNo(i)))

                    While dtr.Read

                        sQry = "Insert into Exception (CustId, DocNo, DocType, AgentId, DocDate, ItemNo) Values (" & SafeSQL(dtr("CustNo").ToString) &
                            "," & SafeSQL(dtr("DocNo").ToString) & "," & SafeSQL(dtr("DocType").ToString) & "," & SafeSQL(dtr("AgentId").ToString) & "," & SafeSQL(Format(dtr("DocDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("ItemID").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export Exception - " & dtr("DocNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteNavSQL(sQry)

                    End While
                    dtr.Close()
                    ExecuteSQLAnother("Update Exception Set Exported = 1 Where DocNo = " & SafeSQL(arrInvNo(i)))
                Catch ex As Exception
                    dtr.Close()
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
        Dim dtr As SqlDataReader = Nothing
        Dim dtr1 As SqlDataReader = Nothing
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

                        sQry = "Insert into BankInHdr (DocNo,SlipNo,DocDate,DocType,AgentId,Amount,BankAccount,Remarks,Exported,MDTNo,Void) Values (" & SafeSQL(dtr("DocNo").ToString) &
                            "," & SafeSQL(dtr("SlipNo").ToString) & "," & SafeSQL(Format(dtr("DocDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("DocType").ToString) & "," & SafeSQL(dtr("AgentId").ToString) &
                             "," & SafeSQL(dtr("Amount").ToString) & "," & SafeSQL(dtr("BankAccount").ToString) & "," & SafeSQL(dtr("Remarks").ToString) & "," & SafeSQL(dtr("Exported").ToString) &
                              "," & SafeSQL(dtr("MDTNo").ToString) & "," & SafeSQL(dtr("Void").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export BankInHdr - " & dtr("DocNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteNavSQL(sQry)

                    End While
                    dtr.Close()

                Catch ex As Exception
                    dtr.Close()
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export BankInHdr - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try

                'Chqamount was added previously it was missed -- Nixon 15/09/2021

                Try
                    dtr1 = ReadRecord("Select * From BankInDet  Where DocNo =  " & SafeSQL(arrInvNo(i)) & " Order by DocNo")
                    While dtr1.Read
                        sSql = "Insert into BankInDet (DocNo,ReceiptNo,ChqNo,ChqDate,ChqAmount,BankName,Remarks) Values (" & SafeSQL(dtr1("DocNo").ToString) &
                             "," & SafeSQL(dtr1("ReceiptNo").ToString) & "," & SafeSQL(dtr1("ChqNo").ToString) & "," & SafeSQL(Format(dtr1("ChqDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr1("ChqAmount").ToString) & ",
                             " & SafeSQL(dtr1("BankName").ToString) & "," & SafeSQL(dtr1("Remarks").ToString) & ")"

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
        Dim dtr As SqlDataReader = Nothing
        Dim dtr1 As SqlDataReader = Nothing
        Dim sExtDocNo As String = ""
        Dim sSql As String
        Try
            dtr = ReadRecord("Select Distinct GoodsExchange.ExchangeNo from GoodsExchange Inner Join GoodsExchangeItem on GoodsExchange.ExchangeNo = GoodsExchangeItem.ExchangeNO" &
                    " where isnull(GoodsExchange.Exported,0) = 0 and GoodsExchangeItem.ItemNo in (Select ItemNO from GoodsExchangeItem Where IsNull(Approved,0) =1 ) " &
                    " Order by GoodsExchange.ExchangeNo")
            While dtr.Read
                If arrInvNo.Contains(dtr("ExchangeNo").ToString) = False Then arrInvNo.Add(dtr("ExchangeNo").ToString)
            End While
            dtr.Close()
            dtr.Dispose()


            For i = 0 To arrInvNo.Count - 1
                Try
                    dtr = ReadRecord("Select Distinct GoodsExchange.* from GoodsExchange Inner Join GoodsExchangeItem on GoodsExchange.ExchangeNo = GoodsExchangeItem.ExchangeNO" &
                        " where isnull(GoodsExchange.Exported,0) = 0 and GoodsExchangeItem.ItemNo in (Select ItemNO from GoodsExchangeItem Where IsNull(Approved,0) =1  " &
                        " and GoodsExchange.ExchangeNo = " & SafeSQL(arrInvNo(i)) & ") Order by GoodsExchange.ExchangeNo")

                    While dtr.Read
                        ExecuteNavSQL("Delete from GoodsExchange where ExchangeNo = " & SafeSQL(arrInvNo(i)))
                        ExecuteNavSQL("Delete from GoodsExchangeItem where ExchangeNo= " & SafeSQL(arrInvNo(i)))

                        sQry = "Insert into GoodsExchange (ExchangeNo,ExchangeDate,CustId,SalesPersonCode,SubTotal,Gst,GstAmt,TotalAmt,Approved,ApprovedBy,SalesUnit) Values (" & SafeSQL(dtr("ExchangeNo").ToString) &
                            "," & SafeSQL(Format(dtr("ExchangeDate"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("CustNo").ToString) & "," & SafeSQL(dtr("SalesPersonCode").ToString) & "," & SafeSQL(dtr("SubTotal").ToString) &
                             "," & SafeSQL(dtr("Gst").ToString) & "," & SafeSQL(dtr("GstAmt").ToString) & "," & SafeSQL(dtr("TotalAmt").ToString) & "," & SafeSQL(dtr("Approved").ToString) &
                              "," & SafeSQL(dtr("ApprovedBy").ToString) & "," & SafeSQL(dtr("MDTNo").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export GoodsExchange - " & dtr("ExchangeNo").ToString) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                        ExecuteNavSQL(sQry)

                    End While
                    dtr.Close()

                Catch ex As Exception
                    dtr.Close()
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export GoodsExchange - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try

                Try
                    dtr1 = ReadRecord("Select Distinct GoodsExchangeItem.* from GoodsExchange Inner Join GoodsExchangeItem on GoodsExchange.ExchangeNo = GoodsExchangeItem.ExchangeNO" &
                            " where isnull(Exported,0) = 0 and GoodsExchangeItem.ItemNo in (Select ItemNO from GoodsExchangeItem Where IsNull(GoodsExchangeItem.Approved,0) =1  and GoodsExchangeItem.ExchangeNo = " & SafeSQL(arrInvNo(i)) & ") Order by GoodsExchangeItem.ExchangeNo")
                    While dtr1.Read
                        sSql = "Insert into GoodsExchangeItem (ExchangeNo,ItemNo,UOM,Quantity,Remarks,Price,SubAmt,CustProdCode,ReasonCode,[LineNo]) Values (" & SafeSQL(dtr1("ExchangeNo").ToString) &
                             "," & SafeSQL(dtr1("ItemNo").ToString) & "," & SafeSQL(dtr1("UOM").ToString) & "," & SafeSQL(dtr1("Quantity").ToString) & "," & SafeSQL(dtr1("Remarks").ToString) &
                             "," & SafeSQL(dtr1("Price").ToString) & "," & SafeSQL(dtr1("SubAmt").ToString) & "," & SafeSQL(dtr1("CustProdCode").ToString) &
                             "," & SafeSQL(dtr1("ReasonCode").ToString) & "," & SafeSQL(dtr1("LineNo").ToString) & ")"

                        ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Export GoodsExchangeItem - " & SafeSQL(arrInvNo(i))) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sSql) & ")")

                        ExecuteNavSQL(sSql)
                    End While
                    dtr1.Close()
                    ExecuteSQLAnother("Update GoodsExchange Set Exported = 1 Where ExchangeNo = " & SafeSQL(arrInvNo(i)))
                Catch ex As Exception
                    dtr1.Close()
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export GoodsExchangeItem'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                    System.IO.File.AppendAllText(Application.StartupPath & "\ErrorLog.txt", "GoodsExchangeItem  Error" & "   " & Date.Now.ToString & " " & ex.Message & vbCrLf)
                End Try
            Next
        Catch ex As Exception
            dtr.Close()
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Error in Export GoodsExchangeItem'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
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

                        sQry = "Insert into Service (ServiceId,ServiceDt,Details,CustId,AgentId,ReasonCode) Values (" & SafeSQL(dtr("ServiceId").ToString) &
                            "," & SafeSQL(dtr("ServiceDt").ToString) & "," & SafeSQL(dtr("Details").ToString) & "," & SafeSQL(dtr("CustId").ToString) & "," & SafeSQL(dtr("AgentId").ToString) &
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
    Private Sub ExportExpiryItem()
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

                        sQry = "Insert into Service (ServiceId,ServiceDt,Details,CustId,AgentId,ReasonCode) Values (" & SafeSQL(dtr("ServiceId").ToString) &
                            "," & SafeSQL(dtr("ServiceDt").ToString) & "," & SafeSQL(dtr("Details").ToString) & "," & SafeSQL(dtr("CustId").ToString) & "," & SafeSQL(dtr("AgentId").ToString) &
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
        Try
            Dim dtr As SqlDataReader
            Dim dtr1 As SqlDataReader
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import SalesAgent - Insert Start'," & SafeSQL(NavCompanyName) & ",''" & ")")

            dtr = ReadNavRecord("Select Code, Name, isnull(SalesTarget,0) as SalesTarget, UserID, Password, isnull(SalesSupervisor,'') as SalesSupervisor, SalesUnit, IsNull(UserGroup,'') as UserGroup, UserCategory from SalesAgent where IsNull(IsRead,0) =0 ")


            While dtr.Read
                If IsAgentExists(dtr("Code").ToString) = False Then
                    If dtr("Code").ToString <> "" Then
                        dtr1 = ReadRecord("Select groupid from usergroup where description = " & SafeSQL(dtr("UserGroup")))
                        If dtr1.Read Then
                            ExecuteSQLAnother("Insert into SalesAgent(Code,Name, UserID, Password, Access, Active,  SalesTarget, CurMonTarget, SalesSupervisor, Department,  SolutionName, UserCategory) Values (" & SafeSQL(dtr("Code").ToString.Trim) & "," & SafeSQL(dtr("Name").ToString) & "," & SafeSQL(dtr("UserID").ToString.Trim) & "," & SafeSQL(dtr("Password").ToString) & "," & SafeSQL(dtr1("GroupId")) & ",1," & dtr("SalesTarget") & ",0," & SafeSQL(dtr("SalesSupervisor").ToString) & "," & SafeSQL(dtr("SalesUnit").ToString.Trim) & ", 'SALES' ," & SafeSQL(dtr("UserCategory").ToString) & ")")
                        Else
                            ExecuteSQLAnother("Insert into SalesAgent(Code,Name, UserID, Password, Access, Active,  SalesTarget, CurMonTarget, SalesSupervisor, Department,  SolutionName, UserCategory) Values (" & SafeSQL(dtr("Code").ToString.Trim) & "," & SafeSQL(dtr("Name").ToString) & "," & SafeSQL(dtr("UserID").ToString.Trim) & "," & SafeSQL(dtr("Password").ToString) & ",2,1," & dtr("SalesTarget") & ",0," & SafeSQL(dtr("SalesSupervisor").ToString) & "," & SafeSQL(dtr("SalesUnit").ToString.Trim) & ", 'SALES' ," & SafeSQL(dtr("UserCategory").ToString) & ")")
                        End If
                        dtr1.Close()
                    End If
                Else
                    If dtr("Code").ToString <> "" Then
                        dtr1 = ReadRecord("Select groupid from usergroup where description = " & SafeSQL(dtr("UserGroup")))
                        If dtr1.Read Then
                            ExecuteSQLAnother("Update SalesAgent Set Name = " & SafeSQL(dtr("Name").ToString) & ", SalesSupervisor=" & SafeSQL(dtr("SalesSupervisor").ToString) & ", SalesTarget=" & dtr("SalesTarget") & ", UserId =" & SafeSQL(dtr("UserID").ToString) & ", Password =" & SafeSQL(dtr("Password").ToString) & ", Access = " & SafeSQL(dtr1("GroupId")) & ", Active = 1, UserCategory = " & SafeSQL(dtr("UserCategory").ToString) & " Where Code = " & SafeSQL(dtr("Code").ToString))
                        Else
                            ExecuteSQLAnother("Update SalesAgent Set Name = " & SafeSQL(dtr("Name").ToString) & ", SalesSupervisor=" & SafeSQL(dtr("SalesSupervisor").ToString) & ", SalesTarget=" & dtr("SalesTarget") & ", UserId =" & SafeSQL(dtr("UserID").ToString) & ", Password =" & SafeSQL(dtr("Password").ToString) & ", Active = 1 , UserCategory = " & SafeSQL(dtr("UserCategory").ToString) & " Where Code = " & SafeSQL(dtr("Code").ToString))
                        End If
                        dtr1.Close()
                    End If
                End If
                ExecuteNavAnotherSQL("Update SalesAgent Set IsRead = 1 where Code = " & SafeSQL(dtr("Code").ToString))
            End While
            dtr.Close()
            dtr.Dispose()
            dtr = Nothing

            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import SalesAgent - Insert Finish'," & SafeSQL(NavCompanyName) & ",''" & ")")
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import SalesAgent - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try


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


    Public Sub UpdateGPSCoordinates(_custNo As String)
        Try
            Dim sSQL As String
            Dim arrCustNo As New ArrayList()
            Dim dtr As SqlDataReader
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import GPSCoOrdinates - Insert Start'," & SafeSQL(NavCompanyName) & ",''" & ")")
            arrCustNo.Clear()
            'and (Latitude=0 or Longitude=0 or Latitude is Null or Longitude is Null) 
            dtr = ReadNavRecord("Select Distinct Address  from Customer where Active=1 and isnull(address,'') <>'' and ( IsNull(Longitude,0) = 0 Or IsNull(Latitude,0)  = 0 ) and CustNo = " & SafeSQL(_custNo) & " order by Address")
            While dtr.Read = True
                arrCustNo.Add(dtr("Address").ToString)
            End While
            dtr.Close()

            For i = 0 To arrCustNo.Count - 1
                Try
                    Dim sLoc As String = GetGeoCoords(arrCustNo(i) & ",  thailand ", 1)
                    If sLoc <> "" Then
                        Dim S() As String = sLoc.Split(",")
                        sSQL = "UPDATE Customer SET Longitude= " & S(1) & " , Latitude=" & S(0) & " Where Address=" & SafeSQL(arrCustNo(i))
                        ExecuteSQL(sSQL)
                        'sb.Append(sLoc & vbCrLf)
                    End If
                Catch ex As Exception
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import GPSCoOrdinates - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try
            Next
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import GPSCoOrdinates - Insert Finish'," & SafeSQL(NavCompanyName) & ",''" & ")")
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import GPSCoOrdinates - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try

    End Sub


    Public Sub ImportUOM()
        Dim dtr As SqlDataReader
        dtr = ReadNavRecord("Select  * from  UOM ")
        While dtr.Read
            If IsExists("Select ItemNo from UOM where ItemNo=" & SafeSQL(dtr("ItemNo").ToString.Trim) & " and Uom=" & SafeSQL(dtr("UOM").ToString.Trim)) = False Then
                ExecuteSQLAnother("Insert into UOM(ItemNo, Uom , BaseQty) Values (" & SafeSQL(dtr("ItemNo").ToString) & "," & SafeSQL(dtr("UOM").ToString) & "," & SafeSQL(dtr("BaseQty").ToString) & ")")
            Else
                ExecuteSQLAnother("Update UOM set BaseQty=" & SafeSQL(dtr("BaseQty").ToString) & " Where ItemNo=" & SafeSQL(dtr("ItemNo").ToString.Trim) & " and UOM=" & SafeSQL(dtr("UOM").ToString.Trim))
            End If
        End While
        dtr.Close()
    End Sub

    Public Sub ImportPayMethod()

    End Sub


    Public Sub ImportPayterms()
        Try
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import PayTerms - Insert Start'," & SafeSQL(NavCompanyName) & ",''" & ")")
            Dim dtr As SqlDataReader
            dtr = ReadRecord("Select Distinct PaymentTerms from Customer")
            While dtr.Read
                If IsExists("Select Code from PayTerms where code=" & SafeSQL(dtr("PaymentTerms"))) = False Then
                    ExecuteSQLAnother("Insert into PayTerms(Code , Description, DueDateCalc, DisDateCalc, DiscountPercent,Active,DTG) Values (" & SafeSQL(dtr("PaymentTerms").ToString) & "," & SafeSQL(dtr("PaymentTerms").ToString) & "," & SafeSQL(dtr("PaymentTerms").ToString + "D") & "," & SafeSQL("0D") & ",0,1," & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")) & ")")
                Else
                    ExecuteSQLAnother("Update PayTerms set Description=" & SafeSQL(dtr("PaymentTerms").ToString) & ", DueDateCalc =" & SafeSQL(dtr("PaymentTerms").ToString + "D") & " Where Code=" & SafeSQL(dtr("PaymentTerms").ToString))
                End If
            End While
            dtr.Close()
            dtr.Dispose()
            dtr = Nothing
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import PayTerms - Insert Finish'," & SafeSQL(NavCompanyName) & ",''" & ")")
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import PayTerms - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try

    End Sub


    Public Sub ImportPriceGroup()
        Try
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import PriceGroup - Insert Start'," & SafeSQL(NavCompanyName) & ",''" & ")")
            Dim dtr As SqlDataReader
            dtr = ReadNavRecord("Select * from PriceGroup where IsNull(IsRead,0) = 0")
            While dtr.Read
                If IsExists("Select Code from PriceGroup where code=" & SafeSQL(dtr("Code"))) = False Then
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import PriceGroup - Insert PriceGroup'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("Code").ToString) & ")")

                    ExecuteSQLAnother("Insert into PriceGroup(Code , Description) Values (" & SafeSQL(dtr("Code").ToString) & "," & SafeSQL(dtr("Description").ToString) & ")")
                Else
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import PriceGroup - Update PriceGroup'," & SafeSQL(NavCompanyName) & "," & SafeSQL(dtr("Code").ToString) & ")")

                    ExecuteSQLAnother("Update PriceGroup set Description=" & SafeSQL(dtr("Description").ToString))
                End If
                ExecuteNavAnotherSQL("Update PriceGroup set IsRead = 1  where Code = " & SafeSQL(dtr("Code")))
            End While
            dtr.Close()
            dtr.Dispose()
            dtr = Nothing
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import PriceGroup - Insert Finish'," & SafeSQL(NavCompanyName) & ",''" & ")")
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import PriceGroup - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try

    End Sub

    Public Sub ImportInvoice()
        Dim sQry As String = ""
        Try

            Dim dtr As SqlDataReader
            Dim arrList As New ArrayList
            dtr = ReadNavRecord("Select * from InvoiceErp")
            While dtr.Read
                If arrList.Contains(dtr("InvNo").ToString) = False Then arrList.Add(dtr("InvNo").ToString)
            End While
            dtr.Close()
            dtr.Dispose()
            For i = 0 To arrList.Count - 1
                'Header 
                Try
                    If IsExists("Select InvNo from Invoice where InvNo = " & SafeSQL(arrList(i))) = True Then
                        'Update
                        Try
                            dtr = ReadNavRecord("Select * from InvoiceErp where InvNo =  " & SafeSQL(arrList(i)) & " Order by InvoiceErp.InvNo")
                            While dtr.Read
                                ExecuteSQL("Update Invoice set PaidAmt = " & dtr("PaidAmt") & " where InvNo = " & SafeSQL(arrList(i)))
                            End While
                            dtr.Close()
                        Catch ex As Exception
                            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import InvoiceErp - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                        End Try

                    Else
                        'Insert
                        dtr = ReadNavRecord("Select * from InvoiceErp where InvNo =  " & SafeSQL(arrList(i)) & " Order by InvoiceErp.InvNo")
                        While dtr.Read

                            sQry = "Insert into Invoice([InvNo] ,[InvDt] ,[OrdNo] ,[CustId] ,[AgentId] ,[SubTotal] ,[GstAmt] ,[TotalAmt] ,[PaidAmt] ,[PayTerms] ,[Void] ,[GST] ,[SalesUnit] ,[DeliveryDate] ,[Remarks] ,[Discount]) VALUES " _
                                & "(" & SafeSQL(dtr("InvNo").ToString) & "," & SafeSQL(Format(dtr("InvDt"), "yyyy-MM-dd")) & "," & SafeSQL(dtr("OrdNo").ToString) & "," & SafeSQL(dtr("CustId").ToString) _
                                & "," & SafeSQL(dtr("AgentId").ToString) & "," & SafeSQL(dtr("SubTotal").ToString) & "," & SafeSQL(dtr("GstAmt").ToString) & "," & SafeSQL(dtr("TotalAmt").ToString) _
                                & "," & SafeSQL(dtr("PaidAmt").ToString) & "," & SafeSQL(dtr("PayTerms").ToString) & "," & SafeSQL(dtr("Void").ToString) & "," & SafeSQL(dtr("Gst").ToString) & "," & SafeSQL(dtr("SalesUnit").ToString) & "," & SafeSQL(dtr("DeliveryDate").ToString) & "," & SafeSQL(dtr("Remarks").ToString) & "," & SafeSQL(dtr("Discount").ToString) & ")"

                            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Import InvoiceErp - " & SafeSQL(arrList(i))) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                            ExecuteSQL(sQry)

                        End While
                        dtr.Close()

                        dtr = ReadNavRecord("Select * from InvItemErp where InvNo =  " & SafeSQL(arrList(i)) & " Order by InvItemErp.InvNo")
                        While dtr.Read

                            sQry = "Insert into InvItem ([InvNo] ,[ItemNo] ,[UOM] ,[Qty] ,[Price] ,[Discount] ,[SubAmt] ,[SalesType] ,[ReasonCode] ,[LineNo] ,[DisPer] ,[PriceGroup] ,[AttachedLineNo]) VALUES " _
                                    & "(" & SafeSQL(dtr("InvNo").ToString) & "," & SafeSQL("ItemNo") & "," & SafeSQL(dtr("UOM").ToString) & "," & SafeSQL(dtr("Qty").ToString) _
                                    & "," & SafeSQL(dtr("Price").ToString) & "," & SafeSQL(dtr("Discount").ToString) & "," & SafeSQL(dtr("SubAmt").ToString) & "," & SafeSQL(dtr("SalesType").ToString) _
                                    & "," & SafeSQL(dtr("ReasonCode").ToString) & "," & SafeSQL(dtr("LineNo").ToString) & "," & SafeSQL(dtr("DisPer").ToString) & "," & SafeSQL(dtr("PriceGroup").ToString) & "," & SafeSQL(dtr("AttachedLineNo").ToString) & ")"

                            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Import InvItemErp - " & SafeSQL(arrList(i))) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                            ExecuteSQL(sQry)

                        End While
                        dtr.Close()
                        sQry = "Update invitem set bulkuom =(Select bulkuom from item where ItemNo = invitem.itemno)"
                        ExecuteSQL(sQry)
                        sQry = "Update Invitem set BulkQty = qty/(Select BaseQty from uom where itemno = invitem.itemno and uom = invitem.bulkuom)"
                        ExecuteSQL(sQry)
                        sQry = "Update Invitem set LooseUom = (Select LooseUom from item where ItemNo = invitem.itemno)"
                        ExecuteSQL(sQry)
                        sQry = "Update Invitem set LooseQty = (qty % (Select BaseQty from uom where itemno = invitem.itemno and uom = invitem.bulkuom))/(Select BaseQty from uom where itemno = invitem.itemno and uom = invitem.looseUom)"
                        ExecuteSQL(sQry)
                        sQry = "Update Invitem Set SubAmtBefDIs = (Qty * Price)"
                        ExecuteSQL(sQry)

                    End If

                Catch ex As Exception
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import InvoiceErp - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try
            Next

        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import InvoiceErp - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try

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
        Dim sQry As String
        Try
            Dim dtr As SqlDataReader
            Dim arrList As New ArrayList
            dtr = ReadNavRecord("Select * from CreditNoteErp")
            While dtr.Read
                If arrList.Contains(dtr("CreditNoteNo").ToString) = False Then arrList.Add(dtr("CreditNoteNo").ToString)
            End While
            dtr.Close()
            dtr.Dispose()
            For i = 0 To arrList.Count - 1
                'Header 
                Try
                    If IsExists("Select CreditNoteNo from CreditNote where CreditNoteNo = " & SafeSQL(arrList(i))) = True Then
                        'Update
                        dtr = ReadNavRecord("Select * from CreditNoteErp where CreditNoteNo =  " & SafeSQL(arrList(i)) & " Order by CreditNoteErp.CreditNoteNo")
                        While dtr.Read
                            ExecuteSQL("Update CreditNote set PaidAmt = " & dtr("PaidAmt") & " where CreditNoteNo =" & SafeSQL(arrList(i)))
                        End While
                        dtr.Close()
                    Else
                        'Insert
                        dtr = ReadNavRecord("Select * from CreditNoteErp where CreditNoteNo =  " & SafeSQL(arrList(i)) & " Order by CreditNoteErp.CreditNoteNo")
                        While dtr.Read

                            sQry = "Insert into CreditNote([CreditNoteNo] ,[CreditDate] ,[CustNo] ,[GoodsReturnNo] ,[SalesPersonCode] ,[SubTotal] ,[GST] ,[TotalAmt] ,[PaidAmt] ,[Remarks] ,[Payterms] ,[Void]) VALUES " _
                                & "(" & SafeSQL(dtr("CreditNoteNo").ToString) & "," & SafeSQL(Format(dtr("CreditDate"), "yyyy-MM-dd")) & ",''," & SafeSQL(dtr("GoodsReturnNo").ToString) & "," & SafeSQL(dtr("SalesPersonCode").ToString) _
                                & "," & SafeSQL(dtr("SubTotal").ToString) & "," & SafeSQL(dtr("GST").ToString) & "," & SafeSQL(dtr("TotalAmt").ToString) & "," & SafeSQL(dtr("PaidAmt").ToString) _
                                & "," & SafeSQL(dtr("Remarks").ToString) & "," & SafeSQL(dtr("Payterms").ToString) & "," & SafeSQL(dtr("Void").ToString) & ")"

                            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Import CreditNote - " & SafeSQL(arrList(i))) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                            ExecuteSQL(sQry)

                        End While
                        dtr.Close()

                        dtr = ReadNavRecord("Select * from CreditNoteDetErp where CreditNoteNo =  " & SafeSQL(arrList(i)) & " Order by CreditNoteDetErp.CreditNoteNo")
                        While dtr.Read

                            sQry = "Insert into CreditNoteDet ([CreditNoteNo] ,[ItemNo] ,[Uom] ,[Baseuom] ,[Price] ,[Qty] ,[Amt] ,[LineNo] ,[Disper] ,[AttachedLineNo] ,[SalesType]) VALUES " _
                                    & "(" & SafeSQL(dtr("CreditNoteNo").ToString) & "," & SafeSQL("ItemNo") & "," & SafeSQL(dtr("UOM").ToString) & "," & SafeSQL(dtr("BaseUOM").ToString) _
                                    & "," & SafeSQL(dtr("Price").ToString) & "," & SafeSQL(dtr("Qty").ToString) & "," & SafeSQL(dtr("Amt").ToString) & "," & SafeSQL(dtr("LineNo").ToString) _
                                    & "," & SafeSQL(dtr("DisPer").ToString) & "," & SafeSQL(dtr("AttachedLineNo").ToString) & "," & SafeSQL(dtr("SalesType").ToString) & ")"

                            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate()," & SafeSQL("Import CreditNoteDet - " & SafeSQL(arrList(i))) & "," & SafeSQL(NavCompanyName) & "," & SafeSQL(sQry) & ")")

                            ExecuteSQL(sQry)

                        End While
                        dtr.Close()

                    End If
                Catch ex As Exception
                    ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import InvoiceErp - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
                End Try
            Next
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Import CreditNote - Insert Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try
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
        dtr = ReadRecordAnother("Select CustNo from Customer where CustNo = " & SafeSQL(sCustNo))
        If Not dtr Is Nothing Then
            bAns = 1
            dtr.Close()
        Else
            bAns = 0
        End If
        Return bAns
    End Function
    Private Function IsCustomerExistsInContacts(ByVal sCustNo As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = ReadRecordAnother("Select CustNo from Contacts where CustNo = " & SafeSQL(sCustNo))
        If Not dtr Is Nothing Then
            bAns = 1
            dtr.Close()
        Else
            bAns = 0
        End If
        Return bAns
    End Function
    Private Function IsNewCustomerExistsInNav(ByVal sNewCustNo As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = ReadRecord("Select CustNo from Customer where CustNo = " & SafeSQL(sNewCustNo))
        If Not dtr Is Nothing Then
            bAns = 1
            dtr.Close()
        Else
            bAns = 0
        End If
        Return bAns
    End Function
    Private Function IsCustomerExistsInSimplrNewCustomer(ByVal sCustNo As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = ReadNavRecordAnother("Select CustNo from SimplrNewCustomer where CustNo = " & SafeSQL(sCustNo))
        bAns = dtr.Read
        dtr.Close()
        Return bAns
    End Function
    Private Function IsCustomerExistsInSalesOrder(ByVal sCustNo As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = ReadRecord("Select CustId from orderhdr where CustId = " & SafeSQL(sCustNo))
        bAns = dtr.Read
        dtr.Close()
        Return bAns
    End Function
    Private Function IsCustomerExistsInInvoice(ByVal sCustNo As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = ReadRecord("Select CustId from Invoice where CustId = " & SafeSQL(sCustNo))
        bAns = dtr.Read
        dtr.Close()
        Return bAns
    End Function
    Private Function IsCustomerExistsInPayment(ByVal sCustNo As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = ReadRecord("Select CustId from Receipt where CustId = " & SafeSQL(sCustNo))
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
    Private Function IsDOExists(ByVal sDo As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = ReadRecord("Select OrdNo from DeliveryOrderHdr where OrdNo = " & SafeSQL(sDo))
        bAns = dtr.Read
        dtr.Close()
        Return bAns
    End Function
    Private Function IsDOItemExists(ByVal sDo As String, ByVal itemno As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = ReadRecord("Select OrdNo from DeliveryOrdItem where OrdNo = " & SafeSQL(sDo) & " and ItemNo = " & SafeSQL(itemno))
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


    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub


    Public Sub loadCombo()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Code, Name from SalesAgent, MDT where MDT.AgentID=SalesAgent.Code")

        aAgent.Clear()
        aAgent.Add(New ComboValues("ALL", "ALL"))
        While dtr.Read()
            aAgent.Add(New ComboValues(dtr("Code").ToString, dtr("Name").ToString))
            '    iSelIndex = iIndex
            'End If
            'iIndex = iIndex + 1
        End While
        dtr.Close()

    End Sub

    Private Sub btnEx_Click(sender As Object, e As EventArgs) Handles btnEx.Click
        Try
            ExportNewCustomer()
            ExportInvoices()
            ExportSalesOrder()
            ExportCreditMemo()
            ExportPayment()
            ExportStockOrder()
            ExportCustVisit()
            ExportBank()
            ExportStockInItem()
            ExportExchange()
            ExportException()
            ExportItemTrans()
            ExportReturn()
        Catch ex As Exception
            ExecuteSQLAnother("Insert into ErrorLog(DTG, FunctionName, CompanyName, ErrorText) values (GetDate(),'Export Error'," & SafeSQL(NavCompanyName) & "," & SafeSQL(ex.Message) & ")")
        End Try
    End Sub

    Private Sub ImportStatus_HandleDestroyed(sender As Object, e As EventArgs) Handles Me.HandleDestroyed

    End Sub
End Class