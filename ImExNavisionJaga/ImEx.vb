Imports System
Imports System.Resources
Imports System.Globalization
Imports System.Threading
Imports System.Reflection
Imports System.Data.SqlClient
Imports System.Data.Odbc
Imports SalesInterface.MobileSales

Public Class Imex
    Implements ISalesBase
    Public sCustPostGroup, sGenPostGroup, sGSTPostGroup, sGenJournalTemplate, sGenJournalBatch, sWorkSheetTemplate, sJournalBatch, sItemJnlTemplate, sItemJnlBatch As String
#Region "Localization Variants"
    Private rMgr As ResourceManager
    Private cInfo As CultureInfo
    Private sLang As String
#End Region
    Dim bStatus As Boolean = False
    Private Structure DelCust
        Dim CustID As String
        Dim PrGroup As String
    End Structure
    Dim i As Integer
    Dim cnt As Integer = 0
    Private Structure ArrItemPrice
        Dim ItemCode As String
        Dim SalesCode As String
        Dim MaxDate As Date
        Dim SalesType As Integer
        Dim MinQty As Double
        Dim VariantCode As String
        Dim sUOM As String
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

    Private Sub btnIm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnIm.Click
        bStatus = True
        btnIm.Enabled = False
        btnEx.Enabled = False
        btnImpInvoice.Enabled = False
        ConnectNavDB()
        ConnectAnotherDB()
        ImportData()
        bStatus = False
        'Dim frm As New SelectList
        'frm.frm = Me
        'frm.ShowDialog()
        DisConnect()
        MsgBox("Completed")
    End Sub
    Public Sub ImportDataAuto()
        ' bStatus = True
        Dim sStatus As String = ""
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Status from System")
        If dtr.Read = True Then
            sStatus = dtr("Status").ToString
        End If
        dtr.Close()
        If sStatus = "Completed" Then
            ' DisconnectDB()
            Exit Sub
        End If
        ConnectNavDB()
        ConnectAnotherDB()
        ImportData()
        bStatus = False
        'DisConnect()
        DisconnectNavDB()
        DisconnectAnotherDB()
        ExecuteSQL("Update System set LastImExDate = " & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")))
    End Sub
    Public Sub UpdateSystem()
        Dim dtr As OdbcDataReader
        dtr = ReadNavRecord("Select * from ""Sales & Receivables Setup"" ")
        If dtr.Read Then
            ExecuteSQL("Update System set GenJournalTemplate=" & SafeSQL(dtr("Journal Template Name").ToString) & ", GenJournalBatch=" & SafeSQL(dtr("Journal Batch Name").ToString))
        End If
        dtr.Close()
        'dtr = ReadNavRecord("Select * from ""Purchases & Payables Setup"" ")
        'If dtr.Read Then
        '    ExecuteSQL("Update System set WorkSheetTemplate=" & SafeSQL(dtr("POS Worksheet Template Name").ToString) & ", JournalBatch=" & SafeSQL(dtr("POS Journal Batch Name").ToString))
        'End If
        'dtr.Close()
        '   MsgBox("1")
        dtr = ReadNavRecord("Select * from ""Inventory Setup"" ")
        If dtr.Read Then
            ExecuteSQL("Update System set ItemJnlTemplate=" & SafeSQL(dtr("POS Item Jnl_ Template Name").ToString) & ", ItemJnlBatch=" & SafeSQL(dtr("POS Item Jnl_ Batch Name").ToString))
        End If
        dtr.Close()
    End Sub
    Public Sub DisConnect()
        DisconnectNavDB()
        DisconnectAnotherDB()
        'btnIm.Enabled = True
        'btnEx.Enabled = True
        'btnImpInvoice.Enabled = True
        ExecuteSQL("Update System set LastImExDate = " & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")))
        'MsgBox(rMgr.GetString("Msg_ImpComp"))
    End Sub
    Public Sub GetImagetick(ByVal cnt As Integer)
        Dim fname As String = Application.StartupPath & "\" & "tick.bmp"
        Dim imgTick As New System.Drawing.Bitmap(fname)
        dgvStatus.Rows(cnt).Cells(1).Value = imgTick
        dgvStatus.Refresh()
        System.Threading.Thread.Sleep(10)
    End Sub
    Private Sub GetImagecross(ByVal Name As String)
        Dim fname As String = Application.StartupPath & "\" & "cross.bmp"
        Dim imgTick As New System.Drawing.Bitmap(fname)
        dgvStatus.Rows(cnt).Cells(1).Value = imgTick
        dgvStatus.Refresh()
        System.Threading.Thread.Sleep(10)
        cnt = cnt + 1
    End Sub

    Private Sub GetImagetickIm(ByVal Name As String)
        Dim fname As String = Application.StartupPath & "\" & "tick.bmp"
        Dim imgTick As New System.Drawing.Bitmap(fname)
        dgvStatus.Rows(cnt).Cells(1).Value = imgTick
        dgvStatus.Refresh()
        System.Threading.Thread.Sleep(10)
        cnt = cnt + 1
    End Sub

    Public Sub ImportData()
        ExecuteSQL("Update System set Status='Not Completed'")
        ImportCustomer()
        ImportCategory()
        ImportProduct()
        ImportItemPrice()
        'ImportItemPricePromotion()
        'ImportCustDiscGroup()
        'ImportCustDiscPromotion()
        ImportCustInvDiscount()
        ImportLocation()
        ImportSalesAgent()
        ImportShipMethod()
        ImportPayMethod()
        ImportPayterms()
        ImportPriceGroup()
        ImportUOM()
        ImportZone()
        ImportCreditNote()
        ImportInvoice()
        ImportCustProd()
        ExecuteSQL("Update System set Status='Completed'")
    End Sub

    Public Sub ImportCustProd()
        Dim dtr As SqlDataReader
        ExecuteSQL("Delete from CustProd")
        dtr = ReadRecord("Select Distinct Invoice.CustID, Customer.PriceGroup, InvItem.ItemNo, ShortDesc, Item.BaseUOM as UOM, Item.ItemName from InvItem, Invoice, Item, Customer where Customer.CustNo = Invoice.CustId and Invoice.InvNo = InvItem.InvNo and InvItem.ItemNo = Item.ItemNo")
        While dtr.Read
            Dim dPr As Double = 0
            Dim dItemPr As Double
            dItemPr = GetPrice(dtr("ItemNo").ToString, dtr("CustId").ToString, dtr("PriceGroup").ToString, dtr("UOM").ToString)
            ExecuteSQLAnother("Insert into CustProd (CustId, ItemNo, Description, ItemName, Uom, Price) Values (" & SafeSQL(dtr("CustID")) & " , " & SafeSQL(dtr("ItemNo")) & ", " & SafeSQL(dtr("ShortDesc")) & "," & SafeSQL(dtr("ItemName").ToString) & ", " & SafeSQL(dtr("Uom")) & ", " & dItemPr & " )")
        End While
        dtr.Close()
    End Sub
    Private Function IsCustProdExists(ByVal sItemNo As String, ByVal sCustNo As String, ByVal sUOM As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean = False
        dtr = ReadRecordAnother("Select UOM from CustProd where ItemNo = " & SafeSQL(sItemNo) & " and CustID = " & SafeSQL(sCustNo))
        If dtr.Read = True Then
            If dtr("UOM").ToString = "CTN" Then
                bAns = True
            Else
                bAns = False
            End If
        End If
        dtr.Close()
        Return bAns
    End Function
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
    Public Sub ImportCreditNote()
        Dim dtr As OdbcDataReader
        Dim rs As SqlDataReader
        Dim arrList As New ArrayList
        Dim sHistoryunit As String = ""
        Dim iHistoryNo As Integer = 0
        ExecuteSQL("Update CreditNote set PaidAmt=TotalAmt")
        dtr = ReadNavRecord("Select A.No_, A.""Document Date"", A.""Posting Date"", A.""Sell-to Customer No_"", A.""Salesperson Code"", A.Amount, A.""Amount Including GST"", B.""Remaining Amount"", A.""Payment Terms Code"", B.""Remaining Amt_ (LCY)"", B.""Original Amt_ (LCY)"", B.""Sales (LCY)"" from ""Sales Cr_Memo Header"" A, ""Cust_ Ledger Entry"" B Where A.No_ = B.""Document No_"" and B.""Document Type"" = 3  order by A.No_")
        While dtr.Read
            Dim dDisAmt As Double = 0
            rs = ReadRecord("Select PaidAmt from CreditNote where CreditNoteNo= " & SafeSQL(dtr("No_")))
            If rs.Read = True Then
                rs.Close()
                ExecuteSQL("Update CreditNote set PaidAmt = " & -1 * (dtr("Original Amt_ (LCY)") - dtr("Remaining Amount")) & " where CreditNoteNo =" & SafeSQL(dtr("No_")))
            Else
                Dim aCr As New ArrCrNote
                aCr.CrNo = dtr("No_").ToString
                aCr.CrDate = dtr("Posting Date")
                ' aCr.OrdNo = dtr("Order No_").ToString
                aCr.CustID = dtr("Sell-to Customer No_").ToString
                aCr.AgentID = dtr("SalesPerson Code").ToString
                aCr.Discount = dDisAmt
                'aCr.CurCode = dtr("Currency code").ToString
                'aCr.CurExRate = dtr("Currency Factor")

                aCr.GSTAmt = -1 * (dtr("Original Amt_ (LCY)") - dtr("Amount"))
                aCr.Subtotal = -1 * dtr("Amount")
                aCr.TotalAmt = -1 * dtr("Original Amt_ (LCY)")
                aCr.PaidAmt = -1 * dtr("Remaining Amount")
                aCr.payterms = dtr("Payment Terms Code").ToString
                arrList.Add(aCr)
                rs.Close()
                ExecuteSQL("Insert into CreditNote(CreditNoteNo, CreditDate, CustNo, GoodsReturnNo, SalesPersonCode, Discount, SubTotal, GST, TotalAmt, PaidAmt, Void, Exported,DTG) Values (" & SafeSQL(dtr("No_").ToString) & "," & SafeSQL(Format(dtr("Document Date"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("Sell-to Customer No_").ToString) & ",''," & SafeSQL(dtr("Salesperson Code").ToString) & "," & dDisAmt & "," & (dtr("Amount")) & "," & CStr(((-1 * dtr("Original Amt_ (LCY)")) - (dtr("Amount")))) & "," & -1 * dtr("Original Amt_ (LCY)") & "," & CStr(-1 * ((-1 * dtr("Original Amt_ (LCY)")) - (-1 * dtr("Remaining Amount")))) & ",0,1," & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")) & ")")
            End If
            'End If
            'End If
        End While
        ImportCrNoteItem(arrList)
        dtr = ReadNavRecord("Select ""Document No_"", ""Posting Date"", ""Sell-to Customer No_"", ""Salesperson Code"", Amount, ""Remaining Amount"",  ""Remaining Amt_ (LCY)"", ""Original Amt_ (LCY)"" from ""Cust_ Ledger Entry"" Where ""Document Type"" = 3 and ""Source Code""='SALESJNL' order by ""Document No_""")
        While dtr.Read
            'If dtr("Posting Date") > DateAdd(DateInterval.Year, -1, Date.Now) Then
            If dtr("Remaining Amt_ (LCY)") < 0 Then
                Dim dDisAmt As Double = 0
                rs = ReadRecord("Select PaidAmt from CreditNote where CreditNoteNo= " & SafeSQL(dtr("Document No_")))
                If rs.Read = True Then
                    rs.Close()
                    ExecuteSQL("Update CreditNote set PaidAmt = " & dtr("Original Amt_ (LCY)") - dtr("Remaining Amount") & " where CreditNoteNo =" & SafeSQL(dtr("Document No_")))
                Else
                    'Dim aCr As New ArrCrNote
                    'aCr.CrNo = dtr("Document No_").ToString
                    'aCr.CrDate = dtr("Posting Date")
                    '' aCr.OrdNo = dtr("Order No_").ToString
                    'aCr.CustID = dtr("Sell-to Customer No_").ToString
                    'aCr.AgentID = dtr("SalesPerson Code").ToString
                    'aCr.Discount = dDisAmt
                    ''aCr.CurCode = dtr("Currency code").ToString
                    ''aCr.CurExRate = dtr("Currency Factor")
                    'aCr.GSTAmt = 0 '-1 * (dtr("Original Amt_ (LCY)") - dtr("Amount"))
                    'aCr.Subtotal = -1 * dtr("Amount")
                    'aCr.TotalAmt = -1 * dtr("Original Amt_ (LCY)")
                    'aCr.PaidAmt = -1 * dtr("Remaining Amount")
                    'aCr.payterms = ""
                    'arrList.Add(aCr)
                    rs.Close()
                    ExecuteSQL("Insert into CreditNote(CreditNoteNo, CreditDate, CustNo, GoodsReturnNo, SalesPersonCode, Discount, SubTotal, GST, TotalAmt, PaidAmt, Void, Exported,DTG) Values (" & SafeSQL(dtr("Document No_").ToString) & "," & SafeSQL(Format(dtr("Posting Date"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("Sell-to Customer No_").ToString) & ",''," & SafeSQL(dtr("Salesperson Code").ToString) & ",0," & (-1 * dtr("Amount")) & "," & CStr(((-1 * dtr("Original Amt_ (LCY)")) - (-1 * dtr("Amount")))) & "," & (-1 * dtr("Original Amt_ (LCY)")) & "," & CStr((-1 * dtr("Original Amt_ (LCY)")) - (-1 * dtr("Remaining Amount"))) & ",0,1," & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")) & ")")
                End If
                '    End If
            End If
        End While
        'ImportCrNoteItem(arrList)

        'dtr = ReadNavRecord("SELECT A.""Amount (LCY)"" As ReceiptAmt, A.""Applies-to Doc_ No_"" As AppliedDocNo FROM ""Gen_ Journal Line"" A, ""Sales Invoice Header"" B WHERE A.""Applies-to Doc_ Type"" = 2 AND A.""Applies-to Doc_ No_"" <> '' AND A.""Journal Template Name"" = 'CASH RECEI' AND A.""Account Type"" = 1 and A.""Applies-to Doc_ No_"" = B.No_")
        'While dtr.Read
        '    ExecuteSQL("Update CreditNote set PaidAmt = PaidAmt - " & dtr("ReceiptAmt") & " where CreditNoteNo =" & SafeSQL(dtr("AppliedDocNo")))
        'End While
        'dtr = ReadNavRecord("SELECT ""Cust_ Ledger Entry"".""Document No_"" As AppliedDocNo, ""Cust_ Ledger Entry"".""Amount to Apply"" As ReceiptAmt FROM ""Gen_ Journal Line"",""Cust_ Ledger Entry""  WHERE ""Cust_ Ledger Entry"".""Applies-to ID"" = ""Gen_ Journal Line"".""Applies-to ID"" AND ""Gen_ Journal Line"".""Applies-to Doc_ Type"" = 0 AND ""Gen_ Journal Line"".""Applies-to Doc_ No_"" <> '' AND ""Gen_ Journal Line"".""Journal Template Name"" = 'CASH RECEI' AND ""Gen_ Journal Line"".""Account Type"" = 1")
        'While dtr.Read
        '    ExecuteSQL("Update CreditNote set PaidAmt = PaidAmt - " & dtr("ReceiptAmt") & " where CreditNoteNo =" & SafeSQL(dtr("AppliedDocNo")))
        'End While
    End Sub

    Public Sub ImportCrNoteItem(ByVal arrList As ArrayList)
        Dim dtr As OdbcDataReader
        'Dim sInvNo As String = ""
        'Dim dGST As Double = 0
        'Dim dRAmt As Double = 0
        'dtr = ReadNavRecord("Select count(*) as Cnt from ""Sales Invoice Line"" order by ""Document No_""")
        'If dtr.Read Then
        '    MsgBox(dtr("Cnt"))
        'End If
        dtr = ReadNavRecord("Select * from ""Sales Cr_Memo Line"" order by ""Document No_""")
        While dtr.Read
            If dtr("No_").ToString.Trim <> "" Then
                For iIndex As Integer = 0 To arrList.Count - 1
                    Dim aPr As ArrCrNote
                    aPr = arrList(iIndex)
                    If dtr("Document No_").ToString = aPr.CrNo Then
                        ExecuteSQL("Insert into CreditNoteDet(CreditNoteNo, ItemNo, UOM, BaseUOM, Price, Qty, Amt) Values (" & SafeSQL(dtr("Document No_").ToString) & ", " & SafeSQL(dtr("No_").ToString) & ", " & SafeSQL(dtr("Unit of Measure Code").ToString) & ",'', " & dtr("Unit Price").ToString & "," & dtr("Quantity").ToString & "," & dtr("Amount").ToString & ")")
                        Exit For
                    End If
                Next
            End If
        End While
        'dtr.Close()
    End Sub

    '62820555

    Public Sub ImportInvoice()
        Dim dtr As OdbcDataReader
        Dim rs As SqlDataReader
        Dim arrList As New ArrayList
        Dim sHistoryunit As String = ""
        Dim iHistoryNo As Integer = 0
        ExecuteSQL("Update Invoice set PaidAmt=TotalAmt")
        rs = ReadRecord("Select HistoryUnit, Historyno from system")
        If rs.Read = True Then
            sHistoryunit = rs("HistoryUnit")
            iHistoryNo = rs("HistoryNo")
        End If
        rs.Close()
        dtr = ReadNavRecord("Select A.No_, A.""Document Date"", A.""Order No_"", A.""Sell-to Customer No_"", A.""Salesperson Code"", A.Amount, A.""Amount Including GST"", B.""Remaining Amount"", A.""Payment Terms Code"", A.""Currency Code"", A.""Currency Factor"", B.""Remaining Amt_ (LCY)"", B.""Original Amt_ (LCY)"", B.""Sales (LCY)"", A.""Ship-to Code"" from ""Sales Invoice Header"" A, ""Cust_ Ledger Entry"" B Where A.No_ = B.""Document No_"" and B.""Document Type"" = 2 and A.""Document Date"" > {d " & SafeSQL(Format(DateAdd(DateInterval.Year, -1, Date.Now), "yyyy-MM-dd")) & "} and (A.""Document Date"" > {d " & SafeSQL(Format(DateAdd(DateInterval.Month, iHistoryNo * -1, Date.Now), "yyyy-MM-dd")) & "} Or B.""Remaining Amt_ (LCY)"" > 0)   order by A.No_")
        While dtr.Read
            Dim dDisAmt As Double = 0
            rs = ReadRecord("Select PaidAmt from Invoice where InvNo= " & SafeSQL(dtr("No_")))
            If rs.Read = True Then
                rs.Close()
                ExecuteSQL("Update Invoice set PaidAmt = " & dtr("Original Amt_ (LCY)") - dtr("Remaining Amount") & " where InvNo =" & SafeSQL(dtr("No_")))
            Else
                Dim aInv As New ArrInvoice
                aInv.InvNo = dtr("No_").ToString
                aInv.InvDate = dtr("Document Date")
                aInv.OrdNo = dtr("Order No_").ToString
                '    If dtr("Ship-to Code").ToString <> "" Then
                'aInv.CustID = dtr("Sell-to Customer No_").ToString & "-" & dtr("Ship-to Code").ToString
                'Else
                aInv.CustID = dtr("Sell-to Customer No_").ToString
                'End If
                aInv.AgentID = dtr("SalesPerson Code").ToString
                aInv.Discount = dDisAmt
                aInv.CurCode = dtr("Currency code").ToString
                aInv.CurExRate = dtr("Currency Factor")
                aInv.GSTAmt = dtr("Original Amt_ (LCY)") - dtr("Amount")
                aInv.Subtotal = dtr("Amount")
                aInv.TotalAmt = dtr("Original Amt_ (LCY)")
                aInv.PaidAmt = dtr("Remaining Amount")
                aInv.payterms = dtr("Payment Terms Code").ToString
                arrList.Add(aInv)
                rs.Close()
                ExecuteSQL("Insert into Invoice(InvNo, InvDt, DONo, DoDt ,OrdNo, CustId, AgentID, Discount, SubTotal, GSTAmt, TotalAmt, PaidAmt, Void, PrintNo, PayTerms, CurCode, CurExRate, PONo, Exported,DTG) Values (" & SafeSQL(dtr("No_").ToString) & "," & SafeSQL(Format(dtr("Document Date"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("No_").ToString) & "," & SafeSQL(Format(dtr("Document Date"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("Order No_").ToString) & "," & SafeSQL(aInv.CustID) & "," & SafeSQL(dtr("Salesperson Code").ToString) & "," & dDisAmt & "," & dtr("Amount").ToString & "," & CStr(dtr("Original Amt_ (LCY)") - dtr("Amount")) & "," & dtr("Original Amt_ (LCY)").ToString & "," & CStr(dtr("Original Amt_ (LCY)") - dtr("Remaining Amount")) & ",0,1," & SafeSQL(dtr("Payment Terms Code").ToString) & "," & SafeSQL(dtr("Currency Code").ToString) & "," & dtr("Currency factor").ToString & ",'',1," & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")) & ")")
            End If
            '    End If
            'End If
        End While
        'MsgBox("1")
        'dtr.Close()

        ImportInvItem(arrList)
        'MsgBox("2")
        'Unposted receipts 1 to 1


        'dtr = ReadNavRecord("SELECT A.""Amount (LCY)"" As ReceiptAmt, A.""Applies-to Doc_ No_"" As AppliedDocNo FROM ""Gen_ Journal Line"" A, ""Sales Invoice Header"" B WHERE A.""Applies-to Doc_ Type"" = 'Invoice' AND A.""Applies-to Doc_ No_"" <> '' AND A.""Journal Template Name"" = 'CASH RECEI' AND A.""Account Type"" = 'Customer' and A.""Applies-to Doc_ No_"" = B.No_")
        'While dtr.Read
        '    ExecuteSQL("Update Invoice set PaidAmt = PaidAmt - " & dtr("ReceiptAmt") & " where InvNo =" & SafeSQL(dtr("AppliedDocNo")))
        'End While
        'MsgBox("3")
        'dtr.Close()
        'Unposted receipts 1 to many
        'dtr = ReadNavRecord("SELECT ""Cust_ Ledger Entry"".""Document No_"" As AppliedDocNo, ""Cust_ Ledger Entry"".""Amount to Apply"" As ReceiptAmt FROM ""Gen_ Journal Line"",""Cust_ Ledger Entry""  WHERE ""Cust_ Ledger Entry"".""Applies-to ID"" = ""Gen_ Journal Line"".""Applies-to ID"" AND ""Gen_ Journal Line"".""Applies-to Doc_ Type"" = '' AND ""Gen_ Journal Line"".""Applies-to Doc_ No_"" = '' AND ""Gen_ Journal Line"".""Journal Template Name"" = 'CASH RECEI' AND ""Gen_ Journal Line"".""Account Type"" = 'Customer'")
        'While dtr.Read
        '    ExecuteSQL("Update Invoice set PaidAmt = PaidAmt - " & dtr("ReceiptAmt") & " where InvNo =" & SafeSQL(dtr("AppliedDocNo")))
        'End While
        ''dtr.Close()
        'MsgBox("4")

        ' dtr = ReadNavRecord("Select A.No_, A.""Document Date"", A.""Order No_"", A.""Sell-to Customer No_"", A.""Salesperson Code"", A.Amount, A.""Amount Including GST"" - A.Amount as GSTAmt, A.""Amount Including GST"", A.""Amount Including GST"" - B.""Remaining Amount"" as PayAmt, A.""Payment Terms Code"", A.""Currency Code"", A.""Currency Factor"", B.""Remaining Amt_ (LCY)"" from ""Sales Invoice Header"" A, ""Cust_ Ledger Entry"" B Where A.No_ = B.""Document No_"" and A.""Sell-to Customer No_"" = B.""Customer No_"" and A.""Payment Method Code"" = 'VANDEL'")
        'dtr = ReadNavRecord("Select * from ""Sales Invoice Header"" Where ""Payment Method Code"" = 'VANDEL'")
        'While dtr.Read
        'If IsInvoiceExists(dtr("No_").ToString) = False Then
        'Dim dtr1 As OdbcDataReader
        'Dim dDisAmt As Double = 0
        'dtr1 = ReadNavRecord("Select ""Line No_"", No_, ""Unit of Measure"", Quantity, ""Unit Price"", ""Line Discount %"", ""Line Discount Amount"", Amount, ""Amount Including GST"" - Amount as GstAmt, Quantity, ""Unit of Measure Code"", ""Quantity (Base)""  from ""Sales Invoice Line"" Where ""Document No_"" = " & SafeSQL(dtr("No_").ToString))
        'While dtr1.Read
        'ExecuteSQL("Insert into InvItem(InvNo, [LineNo], ItemNo, UOM, Qty, Foc, Price, DisPer, DisPr, Discount, SubAmt, GstAmt, DeliQty, BaseUOM, BaseQty) Values (" & SafeSQL(dtr("No_").ToString) & ", " & dtr1("Line No_").ToString & ", " & SafeSQL(dtr1("No_").ToString) & ", " & SafeSQL(dtr1("Unit of Measure Code").ToString) & ", " & dtr1("Quantity").ToString & ", 0, " & dtr1("Unit Price").ToString & ", " & dtr1("Line Discount %").ToString & ", 0, " & dtr1("Line Discount Amount") & ", " & dtr1("Amount").ToString & ", " & dtr1("GstAMt").ToString & ", " & dtr1("Quantity").ToString & ", " & SafeSQL(dtr1("Unit of Measure Code").ToString) & ", " & dtr1("Quantity (Base)").ToString & ")")
        '    dDisAmt = dDisAmt + CDbl(dtr1("Line Discount Amount").ToString)
        'End While
        'dtr1.Close()
        'Dim sLC As String
        'Dim dLC As Double
        'sLC = dtr("Currency Code").ToString
        'dLC = dtr("Currency Factor")
        'Dim invstr As String = "Insert into Invoice(InvNo, InvDt, OrdNo, CustId, AgentID, Discount, SubTotal, GSTAmt, TotalAmt, PaidAmt, PayTerms, CurCode, CurExRate, PONo, Exported,DTG) Values (" & SafeSQL(dtr("No_").ToString) & "," & SafeSQL(Format(dtr("Document Date"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("Order No_").ToString) & "," & SafeSQL(dtr("Sell-to Customer No_").ToString) & "," & SafeSQL(dtr("Salesperson Code").ToString) & "," & dDisAmt & "," & dtr("Amount").ToString & "," & dtr("GstAmt").ToString & "," & dtr("Amount Including GST").ToString & "," & dtr("PayAMt").ToString & "," & SafeSQL(dtr("Payment Terms Code").ToString) & "," & SafeSQL(sLC) & "," & CStr(dLC) & ",'',1," & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")) & ")"
        'ExecuteSQL(invstr)
        'End If
        'End While
        'dtr.Close()


        'Dim dtr As OdbcDataReader
        'Dim rs As SqlDataReader
        'Dim arrList As New ArrayList
        'Dim sHistoryunit As String = ""
        'Dim iHistoryNo As Integer = 0
        'ExecuteSQL("Delete from Invoice")
        'ExecuteSQL("Delete from InvItem")
        'rs = ReadRecord("Select HistoryUnit, Historyno from system")
        'If rs.Read = True Then
        '    sHistoryunit = rs("HistoryUnit")
        '    iHistoryNo = rs("HistoryNo")
        'End If
        'rs.Close()
        'dtr = ReadNavRecord("Select A.No_, A.""Document Date"", A.""Order No_"", A.""Sell-to Customer No_"", A.""Salesperson Code"", A.Amount, A.""Amount Including GST"", B.""Remaining Amount"", A.""Payment Terms Code"", A.""Currency Code"", A.""Currency Factor"", B.""Remaining Amt_ (LCY)"", B.""Original Amt_ (LCY)"", B.""Sales (LCY)"", A.""Ship-to Code"" from ""Sales Invoice Header"" A, ""Cust_ Ledger Entry"" B Where A.No_ = B.""Document No_"" and B.""Document Type"" = 'Invoice' and A.""Document Date"" > {d " & SafeSQL(Format(DateAdd(DateInterval.Year, -1, Date.Now), "yyyy-MM-dd")) & "} and (A.""Document Date"" > {d " & SafeSQL(Format(DateAdd(DateInterval.Month, iHistoryNo * -1, Date.Now), "yyyy-MM-dd")) & "} Or B.""Remaining Amt_ (LCY)"" > 0)   order by A.No_")
        'dtr = ReadNavRecord("Select A.No_, A.""Document Date"", A.""Posting Date"", A.""Order No_"", A.""Sell-to Customer No_"", A.""Salesperson Code"", A.Amount, A.""Amount Including GST"", B.""Remaining Amount"", A.""Payment Terms Code"", A.""Currency Code"", A.""Currency Factor"", B.""Remaining Amt_ (LCY)"", B.""Original Amt_ (LCY)"", B.""Sales (LCY)"" from ""Sales Invoice Header"" A, ""Cust_ Ledger Entry"" B Where A.No_ = B.""Document No_"" and B.""Document Type"" = 2  and  (A.""Document Date"" > " & SafeSQL(Format(DateAdd(DateInterval.Month, (-1) * iHistoryNo, Date.Now), "yyyy-MM-dd")) & " or ""Original Amt_ (LCY)"" >  (""Original Amt_ (LCY)"" - ""Remaining Amount"")) order by A.No_")
        'While dtr.Read
        '    '            If dtr("Document Date") > DateAdd(DateInterval.Year, -1, Date.Now) Then
        '    'If dtr("Document Date") > DateAdd(DateInterval.Month, iHistoryNo * -1, Date.Now) Or dtr("Remaining Amt_ (LCY)") > 0 Then
        '    Dim dDisAmt As Double = 0
        '    '        rs = ReadRecord("Select PaidAmt from Invoice where InvNo= " & SafeSQL(dtr("No_")))
        '    '       If rs.Read = True Then
        '    'rs.Close()
        '    '      ExecuteSQL("Update Invoice set PaidAmt = " & dtr("Original Amt_ (LCY)") - dtr("Remaining Amount") & " where InvNo =" & SafeSQL(dtr("No_")))
        '    '     Else
        '    Dim aInv As New ArrInvoice
        '    aInv.InvNo = dtr("No_").ToString
        '    aInv.InvDate = dtr("Posting Date")
        '    aInv.OrdNo = dtr("Order No_").ToString
        '    aInv.CustID = dtr("Sell-to Customer No_").ToString
        '    aInv.AgentID = dtr("SalesPerson Code").ToString
        '    aInv.Discount = dDisAmt
        '    aInv.CurCode = dtr("Currency code").ToString
        '    aInv.CurExRate = dtr("Currency Factor")
        '    aInv.GSTAmt = dtr("Original Amt_ (LCY)") - dtr("Amount")
        '    aInv.Subtotal = dtr("Amount")
        '    aInv.TotalAmt = dtr("Original Amt_ (LCY)")
        '    aInv.PaidAmt = dtr("Remaining Amount")
        '    aInv.payterms = dtr("Payment Terms Code").ToString
        '    arrList.Add(aInv)
        '    rs.Close()
        '    ExecuteSQL("Insert into Invoice(InvNo, InvDt, DONo, DoDt ,OrdNo, CustId, AgentID, Discount, SubTotal, GSTAmt, TotalAmt, PaidAmt, Void, PrintNo, PayTerms, CurCode, CurExRate, PONo, Exported,DTG) Values (" & SafeSQL(dtr("No_").ToString) & "," & SafeSQL(Format(dtr("Posting Date"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("No_").ToString) & "," & SafeSQL(Format(dtr("Posting Date"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("Order No_").ToString) & "," & SafeSQL(dtr("Sell-to Customer No_").ToString) & "," & SafeSQL(dtr("Salesperson Code").ToString) & "," & dDisAmt & "," & dtr("Amount").ToString & "," & CStr(dtr("Original Amt_ (LCY)") - dtr("Amount")) & "," & dtr("Original Amt_ (LCY)").ToString & "," & CStr(dtr("Original Amt_ (LCY)") - dtr("Remaining Amount")) & ",0,1," & SafeSQL(dtr("Payment Terms Code").ToString) & "," & SafeSQL(dtr("Currency Code").ToString) & "," & dtr("Currency factor").ToString & ",'',1," & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")) & ")")
        '    '    End If
        '    '   End If
        '    '  End If
        'End While
        'ImportInvItem(arrList)
        dtr = ReadNavRecord("SELECT A.""Amount (LCY)"" As ReceiptAmt, A.""Applies-to Doc_ No_"" As AppliedDocNo FROM ""Gen_ Journal Line"" A, ""Sales Invoice Header"" B WHERE A.""Applies-to Doc_ Type"" = 2 AND A.""Applies-to Doc_ No_"" <> '' AND A.""Journal Template Name"" = 'CASH RECEI' AND A.""Account Type"" = 1 and A.""Applies-to Doc_ No_"" = B.No_")
        While dtr.Read
            ExecuteSQL("Update Invoice set PaidAmt = PaidAmt - " & dtr("ReceiptAmt") & " where InvNo =" & SafeSQL(dtr("AppliedDocNo")))
        End While
        dtr = ReadNavRecord("SELECT ""Cust_ Ledger Entry"".""Document No_"" As AppliedDocNo, ""Cust_ Ledger Entry"".""Amount to Apply"" As ReceiptAmt FROM ""Gen_ Journal Line"",""Cust_ Ledger Entry""  WHERE ""Cust_ Ledger Entry"".""Applies-to ID"" = ""Gen_ Journal Line"".""Applies-to ID"" AND ""Gen_ Journal Line"".""Applies-to Doc_ Type"" = 0 AND ""Gen_ Journal Line"".""Applies-to Doc_ No_"" <> '' AND ""Gen_ Journal Line"".""Journal Template Name"" = 'CASH RECEI' AND ""Gen_ Journal Line"".""Account Type"" = 1")
        While dtr.Read
            ExecuteSQL("Update Invoice set PaidAmt = PaidAmt - " & dtr("ReceiptAmt") & " where InvNo =" & SafeSQL(dtr("AppliedDocNo")))
        End While
    End Sub

    Public Sub ImportInvoiceEH()
        Dim dtr As OdbcDataReader
        Dim rs As SqlDataReader
        Dim arrList As New ArrayList
        dtr = ReadNavRecord("Select A.No_, A.""Document Date"", A.""Order No_"", A.""Sell-to Customer No_"", A.""Salesperson Code"", A.Amount, A.""Amount Including GST"", B.""Remaining Amount"", A.""Payment Terms Code"", A.""Currency Code"", A.""Currency Factor"", B.""Remaining Amt_ (LCY)"", B.""Original Amt_ (LCY)"", B.""Sales (LCY)"" from ""Sales Invoice Header"" A, ""Cust_ Ledger Entry"" B Where A.No_ = B.""Document No_"" and B.""Document Type"" = 'Invoice' order by A.No_")
        While dtr.Read
            Dim dDisAmt As Double = 0
            rs = ReadRecord("Select PaidAmt from Invoice where InvNo= " & SafeSQL(dtr("No_")))
            If rs.Read = True Then
                rs.Close()
                ExecuteSQL("Update Invoice set PaidAmt = " & dtr("Original Amt_ (LCY)") - dtr("Remaining Amount") & " where InvNo =" & SafeSQL(dtr("No_")))
            Else
                Dim aInv As New ArrInvoice
                aInv.InvNo = dtr("No_").ToString
                aInv.InvDate = dtr("Document Date")
                aInv.OrdNo = dtr("Order No_").ToString
                aInv.CustID = dtr("Sell-to Customer No_").ToString
                aInv.AgentID = dtr("SalesPerson Code").ToString
                aInv.Discount = dDisAmt
                aInv.CurCode = dtr("Currency code").ToString
                aInv.CurExRate = dtr("Currency Factor")
                aInv.GSTAmt = dtr("Original Amt_ (LCY)") - dtr("Amount")
                aInv.Subtotal = dtr("Amount")
                aInv.TotalAmt = dtr("Original Amt_ (LCY)")
                aInv.PaidAmt = dtr("Remaining Amount")
                aInv.payterms = dtr("Payment Terms Code").ToString
                arrList.Add(aInv)
                rs.Close()
                ExecuteSQL("Insert into Invoice(InvNo, InvDt, DONo, DoDt ,OrdNo, CustId, AgentID, Discount, SubTotal, GSTAmt, TotalAmt, PaidAmt, Void, PrintNo, PayTerms, CurCode, CurExRate, PONo, Exported,DTG) Values (" & SafeSQL(dtr("No_").ToString) & "," & SafeSQL(Format(dtr("Document Date"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("No_").ToString) & "," & SafeSQL(Format(dtr("Document Date"), "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(dtr("Order No_").ToString) & "," & SafeSQL(dtr("Sell-to Customer No_").ToString) & "," & SafeSQL(dtr("Salesperson Code").ToString) & "," & dDisAmt & "," & dtr("Amount").ToString & "," & CStr(dtr("Original Amt_ (LCY)") - dtr("Amount")) & "," & dtr("Original Amt_ (LCY)").ToString & "," & CStr(dtr("Original Amt_ (LCY)") - dtr("Remaining Amount")) & ",0,1," & SafeSQL(dtr("Payment Terms Code").ToString) & "," & SafeSQL(dtr("Currency Code").ToString) & "," & dtr("Currency factor").ToString & ",'',1," & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")) & ")")
            End If
        End While
        dtr.Close()
        ImportInvItem(arrList)
        dtr = ReadNavRecord("SELECT A.""Amount (LCY)"" As ReceiptAmt, A.""Applies-to Doc_ No_"" As AppliedDocNo FROM ""Gen_ Journal Line"" A, ""Sales Invoice Header"" B WHERE A.""Applies-to Doc_ Type"" = 'Invoice' AND A.""Applies-to Doc_ No_"" <> '' AND A.""Journal Template Name"" = 'CASH RECEI' AND A.""Account Type"" = 'Customer' and A.""Applies-to Doc_ No_"" = B.No_")
        While dtr.Read
            ExecuteSQL("Update Invoice set PaidAmt = PaidAmt - " & dtr("ReceiptAmt") & " where InvNo =" & SafeSQL(dtr("AppliedDocNo")))
        End While
        dtr.Close()
        dtr = ReadNavRecord("SELECT ""Cust_ Ledger Entry"".""Document No_"" As AppliedDocNo, ""Cust_ Ledger Entry"".""Amount to Apply"" As ReceiptAmt FROM ""Gen_ Journal Line"",""Cust_ Ledger Entry""  WHERE ""Cust_ Ledger Entry"".""Applies-to ID"" = ""Gen_ Journal Line"".""Applies-to ID"" AND ""Gen_ Journal Line"".""Applies-to Doc_ Type"" = '' AND ""Gen_ Journal Line"".""Applies-to Doc_ No_"" = '' AND ""Gen_ Journal Line"".""Journal Template Name"" = 'CASH RECEI' AND ""Gen_ Journal Line"".""Account Type"" = 'Customer'")
        While dtr.Read
            ExecuteSQL("Update Invoice set PaidAmt = PaidAmt - " & dtr("ReceiptAmt") & " where InvNo =" & SafeSQL(dtr("AppliedDocNo")))
        End While
        dtr.Close()
    End Sub
    Private Function IsInvoiceExists(ByVal sInvNo As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = ReadRecord("Select InvNo from Invoice where InvNo = " & SafeSQL(sInvNo))
        bAns = dtr.Read
        dtr.Close()
        Return bAns
    End Function

    Public Sub ImportProduct()
        Dim dtr As OdbcDataReader
        ExecuteSQL("Update Item set Active = 0")
        dtr = ReadNavRecord("Select * from Item")
        Dim icnt As Integer = 1
        While dtr.Read
            If dtr("No_").ToString <> "" Then
                'If dtr("Gen_ Prod_ Posting Group").ToString = "FG" Or dtr("Gen_ Prod_ Posting Group").ToString = "RAW" Then
                Dim sCatcode As String
                If dtr("Item Category Code").ToString <> "" Then
                    sCatcode = dtr("Item Category Code").ToString
                Else
                    sCatcode = "STD"
                End If
                If IsItemExists(dtr("No_")) Then
                    ExecuteSQL("Update Item Set Description = " & SafeSQL(dtr("Description").ToString) & ", ShortDesc = " & SafeSQL(dtr("Search Description").ToString) & ", BaseUOM = " & SafeSQL(dtr("Base Unit of Measure").ToString) & ", ItemName=" & SafeSQL(dtr("Search Description").ToString) & ", UnitPrice = " & dtr("Unit Price") & ", VendorNo = " & SafeSQL(dtr("Vendor No_").ToString) & ", VendorItemNo = " & SafeSQL(dtr("Vendor Item No_").ToString) & ", AllowInvDiscount = " & IIf(dtr("Allow Invoice Disc_"), 1, 0) & ", Active = 1, Category =  " & SafeSQL(sCatcode) & " where ItemNo = " & SafeSQL(dtr("No_").ToString))
                Else
                    ExecuteSQL("Insert into Item(ItemNo, Description, ItemName, ShortDesc, ChineseDesc, BaseUOM, UnitPrice, VendorNo, VendorItemNo, AllowInvDiscount, Active, Category, DisplayNo, Barcode, ItemPostGroup, Brand, CompanyNo, SubCategory, ToPDA, ChangePr) Values (" & SafeSQL(dtr("No_").ToString) & "," & SafeSQL(dtr("Description").ToString) & "," & SafeSQL(dtr("Search Description").ToString) & "," & SafeSQL(dtr("Search Description").ToString) & ",''," & SafeSQL(dtr("Base Unit of Measure").ToString) & ", " & dtr("Unit Price") & ", " & SafeSQL(dtr("Vendor No_").ToString) & ", " & SafeSQL(dtr("Vendor Item No_").ToString) & ", " & IIf(dtr("Allow Invoice Disc_"), 1, 0) & ",1, " & SafeSQL(sCatcode) & ",0,'','','','','',1,0)")
                End If
                'End If
            End If
            icnt += 1
        End While
        ExecuteSQL("Insert into Category(Code , Description) Values ('STD','STD')")
        ExecuteSQL("Update Item set Active=0 where Description like 'DO NOT%'")
        '  dtr.Close()
    End Sub
    Public Sub ImportProductEH()
        Dim dtr As OdbcDataReader
        Dim arr As New ArrayList
        Dim sLTime As String = ""
        Dim sSLTime As String = ""
        Dim IsActive As Boolean = False
        ExecuteSQL("Update Item set Active = 0")
        dtr = ReadNavRecord("Select * from Item")
        While dtr.Read
            arr.Add(dtr("No_").ToString)
            If dtr("Safety Lead Time").ToString = "" Or IsDBNull(dtr("Safety Lead Time")) = True Then
                sSLTime = "0D"
            Else
                sSLTime = dtr("Safety Lead Time").ToString
            End If
            If dtr("Lead Time Calculation").ToString = "" Or IsDBNull(dtr("Lead Time Calculation")) = True Then
                sLTime = "0D"
            Else
                sLTime = dtr("Lead Time Calculation").ToString
            End If
            '  If CBool(dtr("Is-Sale Item").ToString) = True Then
            IsActive = True
            'Else
            IsActive = False
            'End If
            If IsItemExists(dtr("No_")) Then
                ExecuteSQL("Update Item Set Description = " & SafeSQL(dtr("Description").ToString) & ", ShortDesc = " & SafeSQL(dtr("Description 2").ToString) & ", BaseUOM = " & SafeSQL(dtr("Base Unit of Measure").ToString) & ", ItemName=" & SafeSQL(dtr("No_ 2").ToString) & ", UnitPrice = " & dtr("Unit Price") & ", VendorNo = " & SafeSQL(dtr("Vendor No_").ToString) & ", VendorItemNo = " & SafeSQL(dtr("Vendor Item No_").ToString) & ", AllowInvDiscount = " & IIf(dtr("Allow Invoice Disc_"), 1, 0) & ", Active = " & IIf(IsActive, 1, 0) & " , Category =  " & SafeSQL(dtr("Item category code").ToString) & ", LeadTime =  " & SafeSQL(sLTime) & ", SafetyLeadTime =  " & SafeSQL(sSLTime) & " where ItemNo = " & SafeSQL(dtr("No_").ToString))
            Else
                ExecuteSQL("Insert into Item(ItemNo, Description, ItemName, ShortDesc, ChineseDesc, BaseUOM, UnitPrice, VendorNo, VendorItemNo, AllowInvDiscount, Active, Category, DisplayNo, VendorName, Barcode, ItemPostGroup, ItemImage, ProductionBOM, AssemblyBOM, LeadTime, SafetyLeadTime) Values (" & SafeSQL(dtr("No_").ToString) & "," & SafeSQL(dtr("Description").ToString) & "," & SafeSQL(dtr("No_ 2").ToString) & "," & SafeSQL(dtr("Description 2").ToString) & ",''," & SafeSQL(dtr("Base Unit of Measure").ToString) & ", " & dtr("Unit Price") & ", " & SafeSQL(dtr("Vendor No_").ToString) & ", " & SafeSQL(dtr("Vendor Item No_").ToString) & ", " & IIf(dtr("Allow Invoice Disc_"), 1, 0) & "," & IIf(IsActive, 1, 0) & "," & SafeSQL(dtr("Item category code").ToString) & ",1,'','','','',0,0, " & SafeSQL(sLTime) & ", " & SafeSQL(sSLTime) & ")")
            End If
        End While
        dtr.Close()
        'For iIndex As Integer = 0 To arr.Count - 1
        '    dtr = ReadNavRecord("Select * from ""Nonstock Item"" where ""Item No_"" = " & SafeSQL(arr(iIndex)))
        '    If dtr.Read Then
        '        ExecuteSQL("Update Item Set NonStockItem = 1 where ItemNo = " & SafeSQL(arr(iIndex)))
        '    Else
        '        ExecuteSQL("Update Item Set NonStockItem = 0 where ItemNo = " & SafeSQL(arr(iIndex)))
        '    End If
        '    dtr.Close()
        '    'dtr = ReadNavRecord("Select * from ""BOM Component"" where ""Parent Item No_"" = " & SafeSQL(arr(iIndex)))
        '    'If dtr.Read Then
        '    '    ExecuteSQL("Update Item Set AssemblyBOM = 1 where ItemNo = " & SafeSQL(arr(iIndex)))
        '    'Else
        '    '    ExecuteSQL("Update Item Set AssemblyBOM = 0 where ItemNo = " & SafeSQL(arr(iIndex)))
        '    'End If
        '    'dtr.Close()
        'Next
    End Sub
    Public Sub ImportAssemblyBOM()
        Dim dtr As OdbcDataReader
        ExecuteSQL("Delete from AssemblyBOM")
        dtr = ReadNavRecord("Select ""Parent Item No_"", ""Line No_"", ""No_"", Description, ""Quantity per"" from ""BOM Component""")
        While dtr.Read
            ExecuteSQL("Insert into AssemblyBOM(ParentItemNo , [LineNo], ItemNo, Description, Qty) Values (" & SafeSQL(dtr("Parent Item No_").ToString) & "," & dtr("Line No_") & "," & SafeSQL(dtr("No_").ToString) & "," & SafeSQL(dtr("Description").ToString) & "," & SafeSQL(dtr("Quantity per").ToString) & ")")
            ExecuteSQL("Update Item set AssemblyBOM=1 where ItemNo=" & SafeSQL(dtr("No_").ToString))
        End While
        dtr.Close()
    End Sub

    Public Sub ImportCountry()
        Dim dtr As OdbcDataReader
        ExecuteSQL("Delete from Country")
        dtr = ReadNavRecord("Select Code, Name from Country")
        While dtr.Read
            ExecuteSQL("Insert into Country(Code , Name) Values (" & SafeSQL(dtr("Code").ToString) & "," & SafeSQL(dtr("Name").ToString) & ")")
        End While
        dtr.Close()
    End Sub

    Public Sub ImportCategory()
        Dim dtr As OdbcDataReader
        ExecuteSQL("Delete from Category")
        dtr = ReadNavRecord("Select * from ""Item Category""")
        While dtr.Read
            If dtr("Code").ToString <> "" Then
                ExecuteSQL("Insert into Category(Code , Description) Values (" & SafeSQL(dtr("Code").ToString) & "," & SafeSQL(dtr("Description").ToString) & ")")
            End If
        End While
        dtr.Close()
    End Sub

    Public Sub ImportCurrency()
        Dim dtr As OdbcDataReader
        ExecuteSQL("Delete from CurrencyRate")
        dtr = ReadNavRecord("Select ""Currency Code"", Max(""Starting Date"") as MaxDate from ""Currency Exchange Rate"" Group by ""Currency Code""")
        While dtr.Read
            Dim dtr1 As OdbcDataReader
            dtr1 = ReadNavRecord("Select ""Exchange Rate Amount"", Description from ""Currency Exchange Rate"", Currency Where ""Currency Code"" = Code and ""Currency Code"" = " & SafeSQL(dtr("Currency Code").ToString) & " and ""Starting Date"" = '" & Format(dtr("MaxDate"), "yyyy-MM-dd HH:mm:ss") & "'")
            If dtr1.Read Then
                ExecuteSQL("Insert into CurrencyRate(Code , StartDate, ExchangeRate, Description) Values (" & SafeSQL(dtr("Currency Code").ToString) & ",'" & Format(dtr("MaxDate"), "dd/MMM/yyyy") & "'," & dtr1("Exchange Rate Amount").ToString & "," & SafeSQL(dtr1("Description").ToString) & ")")
            End If
            dtr1.Close()
        End While
        dtr.Close()
        dtr = ReadNavRecord("Select ""LCY Code"" from ""General Ledger Setup""")
        If dtr.Read Then
            ExecuteSQL("Insert into CurrencyRate(Code , StartDate, ExchangeRate, Description) Values (" & SafeSQL(dtr("LCY Code").ToString) & ",'01/01/2000',1,'Local Currency')")
        End If
        dtr.Close()
    End Sub

    Public Sub ImportCustInvDiscount()
        Dim dtr As OdbcDataReader
        ExecuteSQL("Delete from CustInvDiscount")
        dtr = ReadNavRecord("Select Code, ""Minimum Amount"",""Discount %"" from ""Cust_ Invoice Disc_""")
        While dtr.Read
            ExecuteSQL("Insert into CustInvDiscount(Code , MinAmount, DiscountPercent) Values (" & SafeSQL(dtr("Code").ToString) & "," & dtr("Minimum Amount").ToString & "," & dtr("Discount %").ToString & ")")
        End While
        dtr.Close()
    End Sub

    Public Sub ImportCustomer()
        'Dim dt As DateTime
        'Dim dtr As OdbcDataReader
        'dt = Date.Now
        'ExecuteSQL("Update Customer set Active = 0")
        'Dim icnt As Integer = 1
        'Dim sCustNo As String
        ''dtr = ReadNavRecord("select A.*, B.code, B.Name as ShipName, B.address as shipadd1, B.""Address 2"" as shipadd2, B.City as shipcity, B.""Post Code"" as shippost from Customer A, ""Ship-to Address"" B where A.No_ = B.""Customer No_")
        'dtr = ReadNavRecord("select ""Customer No_"", Code, Name as ShipName, Address as shipadd1, ""Address 2"" as shipadd2, City as shipcity, ""Post Code"" as shippost from ""Ship-to Address""")
        'While dtr.Read
        '    If dtr("ShipName").ToString.Trim <> "" Then
        '        sCustNo = dtr("Customer No_").ToString '& "-" & dtr("Code").ToString
        '        If IsCustomerExists(sCustNo) Then
        '            ExecuteSQL("Update Customer Set CustName = " & SafeSQL(dtr("ShipName").ToString) & ", ChineseName = " & SafeSQL(dtr("ShipName").ToString) & ", Address = " & SafeSQL(dtr("ShipAdd1").ToString) & ", Address2 = " & SafeSQL(dtr("ShipAdd2").ToString) & ", Address3 = " & SafeSQL("") & ", Address4 = " & SafeSQL("") & ", PostCode = " & SafeSQL(dtr("ShipPost").ToString) & ", City = " & SafeSQL(dtr("ShipCity").ToString) & ", Active = 1 Where CustNo = " & SafeSQL(sCustNo))
        '        Else
        '            ExecuteSQL("Insert into Customer(CustNo, CustName, ChineseName, Address, Address2, Address3, Address4, PostCode, City, AcCustCode, AcBranchRef, Active) Values (" & SafeSQL(dtr("Customer No_").ToString & "-" & dtr("Code").ToString) & "," & SafeSQL(dtr("ShipName").ToString) & "," & SafeSQL(dtr("ShipName").ToString) & "," & SafeSQL(dtr("ShipAdd1").ToString) & ", " & SafeSQL(dtr("ShipAdd2").ToString) & ", " & SafeSQL("") & ", " & SafeSQL("") & "," & _
        '                        SafeSQL(dtr("ShipPost").ToString) & "," & SafeSQL(dtr("ShipCity").ToString) & "," & SafeSQL(dtr("Customer No_").ToString) & "," & SafeSQL(dtr("Code").ToString) & ",1)")
        '        End If
        '        icnt += 1
        '    End If
        'End While
        ''dtr.Close()
        'dtr = ReadNavRecord("Select * from Customer where ""E-mail"" not like '%ZZ'")
        '' where No_ not in (Select ""Customer No_"" from ""Ship-to Address"") and ""Location code"" like 'WHSE-1'")
        'Dim sdCustNo As String = ""
        'While dtr.Read
        '    sdCustNo = dtr("No_").ToString
        '    If dtr("No_").ToString <> "" Then
        '        If dtr("Blocked") <> 3 Then
        '            Dim sBillto As String = dtr("Bill-to Customer No_").ToString
        '            If sBillto.Trim = "" Then sBillto = dtr("No_").ToString
        '            If IsCustomerExists(sdCustNo) = False Then
        '                ExecuteSQL("Insert into Customer(CustNo, CustName, ChineseName, Address, Address2, Address3, Address4, PostCode, City, CountryCode, Phone, ContactPerson, Balance, CreditLimit, ProvisionalBalance, ZoneCode, FaxNo, Email, Website, ICPartner, PriceGroup, PaymentTerms, PaymentMethod, ShipmentMethod, SalesAgent, ShipAgent, Category, Dimension1, Dimension2, [Bill-toNo], InvDisGroup, Location, CurrencyCode, Active, CustPostGroup, DisplayNo, CustType, Photo, MemberType, GSTCustGroup, Exported, SearchName, SupplierCode, ShipName, ShipAddr, ShipAddr2, ShipAddr3, ShipAddr4, ShipPost, ShipCity, [Print],GSTType, Basket, MDTNo, ExportExclude, AcCustCode, AcBranchRef, Remarks, CommissionCode, Rebate) Values (" & SafeSQL(dtr("No_").ToString) & "," & SafeSQL(dtr("Search Name").ToString) & "," & SafeSQL(dtr("Name 2").ToString) & "," & SafeSQL(dtr("Address").ToString) & ", " & SafeSQL(dtr("Address 2").ToString) & ", " & SafeSQL("") & ", " & SafeSQL("") & "," & _
        '                                                SafeSQL(dtr("Post Code").ToString) & "," & SafeSQL(dtr("City").ToString) & "," & SafeSQL(dtr("Country Code").ToString) & "," & SafeSQL(dtr("Phone No_").ToString) & "," & SafeSQL(dtr("Contact").ToString) & ", " & dtr("Balance").ToString & ", " & dtr("Credit Limit (LCY)").ToString & ", 0," & SafeSQL("") & ", " & SafeSQL(dtr("Fax No_").ToString) & ", " & SafeSQL(dtr("E-mail").ToString) & ", " & SafeSQL(dtr("Home Page").ToString) & ", " & SafeSQL("") & ", " & SafeSQL(dtr("Customer Price Group").ToString) & _
        '                                                ", " & SafeSQL(dtr("Payment Terms Code").ToString) & ", " & SafeSQL(dtr("Payment Method Code").ToString) & ", " & SafeSQL(dtr("Shipment Method Code").ToString) & ", " & SafeSQL(dtr("Salesperson Code").ToString) & ", " & SafeSQL(dtr("Shipping Agent Code").ToString) & ", " & SafeSQL("") & ", " & SafeSQL(dtr("Global Dimension 1 Filter").ToString) & ", " & SafeSQL(dtr("Global Dimension 2 Filter").ToString) & "," & SafeSQL(sBillto) & ", " & SafeSQL(dtr("Invoice Disc_ Code").ToString) & ", " & SafeSQL(dtr("Location Code").ToString) & ", " & SafeSQL(dtr("Currency Code").ToString) & ",1,''," & icnt & ",'Customer','','','',1," & SafeSQL(dtr("Search Name").ToString) & ",''," & SafeSQL(dtr("Search Name").ToString) & "," & SafeSQL(dtr("Address").ToString) & ", " & SafeSQL(dtr("Address 2").ToString) & ", " & SafeSQL("") & ", " & SafeSQL("") & "," & _
        '                                                SafeSQL(dtr("Post Code").ToString) & "," & SafeSQL(dtr("City").ToString) & ",'Invoice','Inclusive',0,'',0," & SafeSQL(dtr("No_").ToString) & ",'','','',0)")
        '            Else
        '                ExecuteSQL("Update Customer Set CustName = " & SafeSQL(dtr("Search Name").ToString) & ", ChineseName = " & SafeSQL(dtr("Name 2").ToString) & ", Address = " & SafeSQL(dtr("Address").ToString) & ", Address2 = " & SafeSQL(dtr("Address 2").ToString) & ", PostCode = " & SafeSQL(dtr("Post Code").ToString) & ", City = " & SafeSQL(dtr("City").ToString) & ", CountryCode = " & SafeSQL(dtr("Country Code").ToString) & ", Phone = " & SafeSQL(dtr("Phone No_").ToString) & ", ContactPerson = " & SafeSQL(dtr("Contact").ToString) & ", Balance = " & dtr("Balance").ToString & ", CreditLimit = " & dtr("Credit Limit (LCY)").ToString & ", ZoneCode = " & SafeSQL("") & ", FaxNo = " & SafeSQL(dtr("Fax No_").ToString) & ", email = " & SafeSQL(dtr("E-mail").ToString) & ", Website = " & SafeSQL(dtr("Home Page").ToString) & ", ICPartner = " & SafeSQL("") & ", " & _
        '                        "PriceGroup= " & SafeSQL(dtr("Customer Price Group").ToString) & ", PaymentTerms = " & SafeSQL(dtr("Payment Terms Code").ToString) & ", PaymentMethod = " & SafeSQL(dtr("Payment Method Code").ToString) & ", ShipmentMethod = " & SafeSQL(dtr("Shipment Method Code").ToString) & ", SalesAgent = " & SafeSQL(dtr("Salesperson Code").ToString) & ", ShipAgent = " & SafeSQL(dtr("Shipping Agent Code").ToString) & ", Category = " & SafeSQL("") & ", Dimension1 = " & SafeSQL(dtr("Global Dimension 1 Filter").ToString) & ", Dimension2 = " & SafeSQL(dtr("Global Dimension 2 Filter").ToString) & ", [Bill-toNo] = " & SafeSQL(sBillto) & ", InvDisGroup = " & SafeSQL(dtr("Invoice Disc_ Code").ToString) & ", Location = " & SafeSQL(dtr("Location Code").ToString) & ", CurrencyCode = " & SafeSQL(dtr("Currency Code").ToString) & ",SearchName=" & SafeSQL(dtr("Search Name").ToString) & ", Active = 1 Where CustNo = " & SafeSQL(dtr("No_").ToString))
        '            End If
        '            ExecuteSQL("Update Customer Set [Bill-toNo] = " & SafeSQL(sBillto) & ", CountryCode = " & SafeSQL(dtr("Country Code").ToString) & ", Phone = " & SafeSQL(dtr("Phone No_").ToString) & ", ContactPerson = " & SafeSQL(dtr("Contact").ToString) & ", Balance = " & dtr("Balance").ToString & ", CreditLimit = " & dtr("Credit Limit (LCY)").ToString & ", ProvisionalBalance = 0, ZoneCode = '', FaxNo = " & SafeSQL(dtr("Fax No_").ToString) & ", Email = " & SafeSQL(dtr("E-mail").ToString) & ", Website = " & SafeSQL(dtr("Home Page").ToString) & ", ICPartner = '', PriceGroup = " & SafeSQL(dtr("Customer Price Group").ToString) & ", PaymentTerms = " & SafeSQL(dtr("Payment Terms Code").ToString) & ", PaymentMethod = " & SafeSQL(dtr("Payment Method Code").ToString) & ", ShipmentMethod = " & SafeSQL(dtr("Shipment Method Code").ToString) & ", SalesAgent = " & SafeSQL(dtr("Salesperson Code").ToString) & ", ShipAgent = " & SafeSQL(dtr("Shipping Agent Code").ToString) & ", Category = '', Dimension1 = '', Dimension2 = '', InvDisGroup = " & SafeSQL(dtr("Invoice Disc_ Code").ToString) & ", Location = " & SafeSQL(dtr("Location Code").ToString) & ", CurrencyCode = " & SafeSQL(dtr("Currency Code").ToString) & ", CustPostGroup = '', CustType = 'Customer', Photo = '', MemberType = '', GSTCustGroup = '', Exported = 1, SupplierCode = '', ShipName = " & SafeSQL(dtr("Search Name").ToString) & ", ShipAddr = " & SafeSQL(dtr("Address").ToString) & ", ShipAddr2 = " & SafeSQL(dtr("Address 2").ToString) & ", ShipAddr3 = '', ShipAddr4 = '', ShipPost = " & SafeSQL(dtr("Post Code").ToString) & ", ShipCity = " & SafeSQL(dtr("City").ToString) & ", [Print] = 'Invoice',GSTType = 'Inclusive', Basket = 0, MDTNo = '', ExportExclude = 0, Remarks = '', CommissionCode = '', Rebate = '' Where AcCustCode = " & SafeSQL(dtr("No_").ToString))
        '        Else
        '            ExecuteSQL("Update Customer Set Active = 0 Where AcCustCode = " & SafeSQL(dtr("No_").ToString))
        '        End If
        '        'Dim iBranch As Integer = HasBranch(dtr("No_").ToString)
        '        ''If iBranch = 0 Then
        '        ''    ExecuteSQL("Insert into Customer(CustNo, CustName, ChineseName, Address, Address2, Address3, Address4, PostCode, City, CountryCode, Phone, ContactPerson, Balance, CreditLimit, ProvisionalBalance, ZoneCode, FaxNo, Email, Website, ICPartner, PriceGroup, PaymentTerms, PaymentMethod, ShipmentMethod, SalesAgent, ShipAgent, Category, Dimension1, Dimension2, [Bill-toNo], InvDisGroup, Location, CurrencyCode, Active, CustPostGroup, DisplayNo, CustType, Photo, MemberType, GSTCustGroup, Exported, SearchName, SupplierCode, ShipName, ShipAddr, ShipAddr2, ShipAddr3, ShipAddr4, ShipPost, ShipCity, [Print],GSTType, Basket, MDTNo, ExportExclude, AcCustCode, AcBranchRef, Remarks, CommissionCode, Rebate) Values (" & SafeSQL(dtr("No_").ToString) & "," & SafeSQL(dtr("Search Name").ToString) & "," & SafeSQL(dtr("Name 2").ToString) & "," & SafeSQL(dtr("Address").ToString) & ", " & SafeSQL(dtr("Address 2").ToString) & ", " & SafeSQL("") & ", " & SafeSQL("") & "," & _
        '        ''                SafeSQL(dtr("Post Code").ToString) & "," & SafeSQL(dtr("City").ToString) & "," & SafeSQL(dtr("Country Code").ToString) & "," & SafeSQL(dtr("Phone No_").ToString) & "," & SafeSQL(dtr("Contact").ToString) & ", " & dtr("Balance").ToString & ", " & dtr("Credit Limit (LCY)").ToString & ", 0," & SafeSQL("") & ", " & SafeSQL(dtr("Fax No_").ToString) & ", " & SafeSQL(dtr("E-mail").ToString) & ", " & SafeSQL(dtr("Home Page").ToString) & ", " & SafeSQL("") & ", " & SafeSQL(dtr("Customer Price Group").ToString) & _
        '        ''                ", " & SafeSQL(dtr("Payment Terms Code").ToString) & ", " & SafeSQL(dtr("Payment Method Code").ToString) & ", " & SafeSQL(dtr("Shipment Method Code").ToString) & ", " & SafeSQL(dtr("Salesperson Code").ToString) & ", " & SafeSQL(dtr("Shipping Agent Code").ToString) & ", " & SafeSQL("") & ", " & SafeSQL(dtr("Global Dimension 1 Filter").ToString) & ", " & SafeSQL(dtr("Global Dimension 2 Filter").ToString) & "," & SafeSQL(dtr("Bill-to Customer No_").ToString) & ", " & SafeSQL(dtr("Invoice Disc_ Code").ToString) & ", " & SafeSQL(dtr("Location Code").ToString) & ", " & SafeSQL(dtr("Currency Code").ToString) & ",1,''," & icnt & ",'Customer','','','',1,'SName',''," & SafeSQL(dtr("Search Name").ToString) & "," & SafeSQL(dtr("Address").ToString) & ", " & SafeSQL(dtr("Address 2").ToString) & ", " & SafeSQL("") & ", " & SafeSQL("") & "," & _
        '        ''                SafeSQL(dtr("Post Code").ToString) & "," & SafeSQL(dtr("City").ToString) & ",'Invoice','Inclusive',0,'',0," & SafeSQL(dtr("No_").ToString) & ",'','','',0)")
        '        ''Else
        '        '    If iBranch = 1 Then
        '        '        ExecuteSQL("Update Customer Set CustName = " & SafeSQL(dtr("Search Name").ToString) & ", ChineseName = " & SafeSQL(dtr("Name 2").ToString) & ", Address = " & SafeSQL(dtr("Address").ToString) & ", Address2 = " & SafeSQL(dtr("Address 2").ToString) & ", PostCode = " & SafeSQL(dtr("Post Code").ToString) & ", City = " & SafeSQL(dtr("City").ToString) & ", CountryCode = " & SafeSQL(dtr("Country Code").ToString) & ", Phone = " & SafeSQL(dtr("Phone No_").ToString) & ", ContactPerson = " & SafeSQL(dtr("Contact").ToString) & ", Balance = " & dtr("Balance").ToString & ", CreditLimit = " & dtr("Credit Limit (LCY)").ToString & ", ZoneCode = " & SafeSQL("") & ", FaxNo = " & SafeSQL(dtr("Fax No_").ToString) & ", email = " & SafeSQL(dtr("E-mail").ToString) & ", Website = " & SafeSQL(dtr("Home Page").ToString) & ", ICPartner = " & SafeSQL("") & ", " & _
        '        '                        "PriceGroup= " & SafeSQL(dtr("Customer Price Group").ToString) & ", PaymentTerms = " & SafeSQL(dtr("Payment Terms Code").ToString) & ", PaymentMethod = " & SafeSQL(dtr("Payment Method Code").ToString) & ", ShipmentMethod = " & SafeSQL(dtr("Shipment Method Code").ToString) & ", SalesAgent = " & SafeSQL(dtr("Salesperson Code").ToString) & ", ShipAgent = " & SafeSQL(dtr("Shipping Agent Code").ToString) & ", Category = " & SafeSQL("") & ", Dimension1 = " & SafeSQL(dtr("Global Dimension 1 Filter").ToString) & ", Dimension2 = " & SafeSQL(dtr("Global Dimension 2 Filter").ToString) & ", [Bill-toNo] = " & SafeSQL(dtr("Bill-to Customer No_").ToString) & ", InvDisGroup = " & SafeSQL(dtr("Invoice Disc_ Code").ToString) & ", Location = " & SafeSQL(dtr("Location Code").ToString) & ", CurrencyCode = " & SafeSQL(dtr("Currency Code").ToString) & ", Active = 1 Where CustNo = " & SafeSQL(dtr("No_").ToString))
        '        '    ElseIf iBranch = 2 Then
        '        '        'Dim iCustName As Integer = UpdateCustName(sdCustNo)
        '        '        'If iCustName = 1 Then
        '        '        '    ExecuteSQL("Update Customer Set CustName=" & SafeSQL(dtr("Search Name").ToString) & ", CountryCode = " & SafeSQL(dtr("Country Code").ToString) & ", Phone = " & SafeSQL(dtr("Phone No_").ToString) & ", ContactPerson = " & SafeSQL(dtr("Contact").ToString) & ", Balance = " & dtr("Balance").ToString & ", CreditLimit = " & dtr("Credit Limit (LCY)").ToString & ", ProvisionalBalance = 0, ZoneCode = '', FaxNo = " & SafeSQL(dtr("Fax No_").ToString) & ", Email = " & SafeSQL(dtr("E-mail").ToString) & ", Website = " & SafeSQL(dtr("Home Page").ToString) & ", ICPartner = '', PriceGroup = " & SafeSQL(dtr("Customer Price Group").ToString) & ", PaymentTerms = " & SafeSQL(dtr("Payment Terms Code").ToString) & ", PaymentMethod = " & SafeSQL(dtr("Payment Method Code").ToString) & ", ShipmentMethod = " & SafeSQL(dtr("Shipment Method Code").ToString) & ", SalesAgent = " & SafeSQL(dtr("Salesperson Code").ToString) & ", ShipAgent = " & SafeSQL(dtr("Shipping Agent Code").ToString) & ", Category = '', Dimension1 = '', Dimension2 = '', [Bill-toNo] = " & SafeSQL(dtr("Bill-to Customer No_").ToString) & ", InvDisGroup = " & SafeSQL(dtr("Invoice Disc_ Code").ToString) & ", Location = " & SafeSQL(dtr("Location Code").ToString) & ", CurrencyCode = " & SafeSQL(dtr("Currency Code").ToString) & ", Active = 1, CustPostGroup = '', CustType = 'Customer', Photo = '', MemberType = '', GSTCustGroup = '', Exported = 1, SupplierCode = '', ShipName = '', ShipAddr = " & SafeSQL(dtr("Address").ToString) & ", ShipAddr2 = " & SafeSQL(dtr("Address 2").ToString) & ", ShipAddr3 = '', ShipAddr4 = '', ShipPost = " & SafeSQL(dtr("Post Code").ToString) & ", ShipCity = " & SafeSQL(dtr("City").ToString) & ", [Print] = 'Invoice',GSTType = 'Inclusive', Basket = 0, MDTNo = '', ExportExclude = 0, Remarks = '', CommissionCode = '', Rebate = '' Where AcCustCode = " & SafeSQL(dtr("No_").ToString))
        '        '        'Else
        '        '    ExecuteSQL("Update Customer Set CountryCode = " & SafeSQL(dtr("Country Code").ToString) & ", Phone = " & SafeSQL(dtr("Phone No_").ToString) & ", ContactPerson = " & SafeSQL(dtr("Contact").ToString) & ", Balance = " & dtr("Balance").ToString & ", CreditLimit = " & dtr("Credit Limit (LCY)").ToString & ", ProvisionalBalance = 0, ZoneCode = '', FaxNo = " & SafeSQL(dtr("Fax No_").ToString) & ", Email = " & SafeSQL(dtr("E-mail").ToString) & ", Website = " & SafeSQL(dtr("Home Page").ToString) & ", ICPartner = '', PriceGroup = " & SafeSQL(dtr("Customer Price Group").ToString) & ", PaymentTerms = " & SafeSQL(dtr("Payment Terms Code").ToString) & ", PaymentMethod = " & SafeSQL(dtr("Payment Method Code").ToString) & ", ShipmentMethod = " & SafeSQL(dtr("Shipment Method Code").ToString) & ", SalesAgent = " & SafeSQL(dtr("Salesperson Code").ToString) & ", ShipAgent = " & SafeSQL(dtr("Shipping Agent Code").ToString) & ", Category = '', Dimension1 = '', Dimension2 = '', [Bill-toNo] = " & SafeSQL(dtr("Bill-to Customer No_").ToString) & ", InvDisGroup = " & SafeSQL(dtr("Invoice Disc_ Code").ToString) & ", Location = " & SafeSQL(dtr("Location Code").ToString) & ", CurrencyCode = " & SafeSQL(dtr("Currency Code").ToString) & ", Active = 1, CustPostGroup = '', CustType = 'Customer', Photo = '', MemberType = '', GSTCustGroup = '', Exported = 1, SupplierCode = '', ShipName = '', ShipAddr = " & SafeSQL(dtr("Address").ToString) & ", ShipAddr2 = " & SafeSQL(dtr("Address 2").ToString) & ", ShipAddr3 = '', ShipAddr4 = '', ShipPost = " & SafeSQL(dtr("Post Code").ToString) & ", ShipCity = " & SafeSQL(dtr("City").ToString) & ", [Print] = 'Invoice',GSTType = 'Inclusive', Basket = 0, MDTNo = '', ExportExclude = 0, Remarks = '', CommissionCode = '', Rebate = '' Where AcCustCode = " & SafeSQL(dtr("No_").ToString))
        '        '        'End If
        '        '    End If
        '    End If
        '    icnt += 1
        'End While
        ''dtr.Close()
        'Dim dtrCust As SqlDataReader
        'Dim arr As New ArrayList
        ''Remove duplicate customer id
        'dtrCust = ReadRecord("Select * from Customer order by CustNo")
        'Dim sCust As String = ""
        'While dtrCust.Read
        '    If sCust = dtrCust("CustNo").ToString Then
        '        Dim ds As DelCust
        '        ds.CustID = dtrCust("CustNo").ToString
        '        ds.PrGroup = dtrCust("PriceGroup").ToString
        '        arr.Add(ds)
        '    End If
        '    sCust = dtrCust("CustNo").ToString
        'End While
        'dtrCust.Close()
        'For iindex As Integer = 0 To arr.Count - 1
        '    Dim ds As DelCust
        '    ds = arr(iindex)
        '    ExecuteSQL("Delete from Customer Where CustNo =" & SafeSQL(ds.CustID) & " and PriceGroup = " & SafeSQL(ds.PrGroup))
        'Next
        'ExecuteSQL("Update Customer Set Active=0 where CustName like '%Closed%'")
        'ExecuteSQL("Update Customer Set ShipName = CustName where (ShipName='' or ShipName is Null)")
        'ExecuteSQL("Update Customer Set SearchName = CustName where (SearchName='' or SearchName is Null)")


        Dim dt As DateTime
        Dim dtr As OdbcDataReader
        dt = Date.Now
        ExecuteSQL("Update Customer set Active = 0")
        Dim icnt As Integer = 1
        Dim sCustNo As String
        'dtr = ReadNavRecord("select A.*, B.code, B.Name as ShipName, B.address as shipadd1, B.""Address 2"" as shipadd2, B.City as shipcity, B.""Post Code"" as shippost from Customer A, ""Ship-to Address"" B where A.No_ = B.""Customer No_")
        dtr = ReadNavRecord("select ""Customer No_"", Code, Name as ShipName, Address as shipadd1, ""Address 2"" as shipadd2, City as shipcity, ""Post Code"" as shippost from ""Ship-to Address""")
        While dtr.Read
            sCustNo = dtr("Customer No_").ToString '& "-" & dtr("Code").ToString
            If IsCustomerExists(sCustNo) Then
                ExecuteSQL("Update Customer Set CustName = " & SafeSQL(dtr("ShipName").ToString) & ", ChineseName = " & SafeSQL(dtr("ShipName").ToString) & ", Address = " & SafeSQL(dtr("ShipAdd1").ToString) & ", Address2 = " & SafeSQL(dtr("ShipAdd2").ToString) & ", Address3 = " & SafeSQL("") & ", Address4 = " & SafeSQL("") & ", PostCode = " & SafeSQL(dtr("ShipPost").ToString) & ", City = " & SafeSQL(dtr("ShipCity").ToString) & ", Active = 1 Where CustNo = " & SafeSQL(sCustNo))
            Else
                '& "-" & dtr("Code").ToString)
                ExecuteSQL("Insert into Customer(CustNo, CustName, ChineseName, Address, Address2, Address3, Address4, PostCode, City, AcCustCode, AcBranchRef, Active) Values (" & SafeSQL(dtr("Customer No_").ToString) & "," & SafeSQL(dtr("ShipName").ToString) & "," & SafeSQL(dtr("ShipName").ToString) & "," & SafeSQL(dtr("ShipAdd1").ToString) & ", " & SafeSQL(dtr("ShipAdd2").ToString) & ", " & SafeSQL("") & ", " & SafeSQL("") & "," & _
                            SafeSQL(dtr("ShipPost").ToString) & "," & SafeSQL(dtr("ShipCity").ToString) & "," & SafeSQL(dtr("Customer No_").ToString) & "," & SafeSQL(dtr("Code").ToString) & ",1)")
            End If
            icnt += 1
        End While
        'dtr.Close()

        dtr = ReadNavRecord("Select * from Customer where Blocked <> 1") 'where ""City"" not like '%ZZ'")
        ' where No_ not in (Select ""Customer No_"" from ""Ship-to Address"") and ""Location code"" like 'WHSE-1'")
        Dim sdCustNo As String = ""
        While dtr.Read
            Dim sBillto As String
            sdCustNo = dtr("No_").ToString
            If dtr("Bill-to Customer No_").ToString = "" Then
                sBillto = dtr("No_").ToString
            Else
                sBillto = dtr("Bill-to Customer No_").ToString
            End If

            If dtr("No_").ToString <> "" Then
                Dim iBranch As Integer = HasBranch(dtr("No_").ToString)
                If iBranch = 0 Then
                    ExecuteSQL("Insert into Customer(CustNo, CustName, ChineseName, Address, Address2, Address3, Address4, PostCode, City, CountryCode, Phone, ContactPerson, Balance, CreditLimit, ProvisionalBalance, ZoneCode, FaxNo, Email, Website, ICPartner, PriceGroup, PaymentTerms, PaymentMethod, ShipmentMethod, SalesAgent, ShipAgent, Category, Dimension1, Dimension2, [Bill-toNo], InvDisGroup, Location, CurrencyCode, Active, CustPostGroup, DisplayNo, CustType, Photo, MemberType, GSTCustGroup, Exported, SearchName, SupplierCode, ShipName, ShipAddr, ShipAddr2, ShipAddr3, ShipAddr4, ShipPost, ShipCity, [Print],GSTType, Basket, MDTNo, ExportExclude, AcCustCode, AcBranchRef, Remarks, CommissionCode, Rebate, CustDiscGroup, AcBillRef, BillMultiple) Values (" & SafeSQL(dtr("No_").ToString) & "," & SafeSQL(dtr("Name").ToString) & "," & SafeSQL(dtr("Name 2").ToString) & "," & SafeSQL(dtr("Address").ToString) & ", " & SafeSQL(dtr("Address 2").ToString) & ", " & SafeSQL("") & ", " & SafeSQL("") & "," & _
                                SafeSQL(dtr("Post Code").ToString) & "," & SafeSQL(dtr("City").ToString) & "," & SafeSQL(dtr("Country Code").ToString) & "," & SafeSQL(dtr("Phone No_").ToString) & "," & SafeSQL(dtr("Contact").ToString) & ", " & dtr("Balance").ToString & ", " & dtr("Credit Limit (LCY)").ToString & ", 0," & SafeSQL(dtr("Customer Disc_ Group").ToString) & ", " & SafeSQL(dtr("Fax No_").ToString) & ", " & SafeSQL(dtr("E-mail").ToString) & ", " & SafeSQL(dtr("Home Page").ToString) & ", " & SafeSQL("") & ", " & SafeSQL(dtr("Customer Price Group").ToString) & _
                                ", " & SafeSQL(dtr("Payment Terms Code").ToString) & ", " & SafeSQL(dtr("Payment Method Code").ToString) & ", " & SafeSQL(dtr("Shipment Method Code").ToString) & ", " & SafeSQL(dtr("Salesperson Code").ToString) & ", " & SafeSQL(dtr("Shipping Agent Code").ToString) & ", " & SafeSQL("") & ", " & SafeSQL(dtr("Global Dimension 1 Filter").ToString) & ", " & SafeSQL(dtr("Global Dimension 2 Filter").ToString) & "," & SafeSQL(sBillto) & ", " & SafeSQL(dtr("Invoice Disc_ Code").ToString) & ", " & SafeSQL(dtr("Location Code").ToString) & ", " & SafeSQL(dtr("Currency Code").ToString) & ",1,''," & icnt & ",'Customer','','','',1," & SafeSQL(dtr("Search Name").ToString) & ",''," & SafeSQL(dtr("Search Name").ToString) & "," & SafeSQL(dtr("Address").ToString) & ", " & SafeSQL(dtr("Address 2").ToString) & ", " & SafeSQL("") & ", " & SafeSQL("") & "," & _
                                SafeSQL(dtr("Post Code").ToString) & "," & SafeSQL(dtr("City").ToString) & ",'Invoice','Inclusive',0,'',0," & SafeSQL(dtr("No_").ToString) & ",'','','',0," & SafeSQL(dtr("Customer Disc_ Group").ToString) & ",'',0)")
                ElseIf iBranch = 1 Then
                    ExecuteSQL("Update Customer Set CustName = " & SafeSQL(dtr("Name").ToString) & ", ChineseName = " & SafeSQL(dtr("Name 2").ToString) & ", Address = " & SafeSQL(dtr("Address").ToString) & ", Address2 = " & SafeSQL(dtr("Address 2").ToString) & ",SearchName=" & SafeSQL(dtr("Search Name").ToString) & ", PostCode = " & SafeSQL(dtr("Post Code").ToString) & ", City = " & SafeSQL(dtr("City").ToString) & ", CountryCode = " & SafeSQL(dtr("Country Code").ToString) & ", Phone = " & SafeSQL(dtr("Phone No_").ToString) & ", ContactPerson = " & SafeSQL(dtr("Contact").ToString) & ", Balance = " & dtr("Balance").ToString & ", CreditLimit = " & dtr("Credit Limit (LCY)").ToString & ", ZoneCode = " & SafeSQL(dtr("Customer Disc_ Group").ToString) & ", FaxNo = " & SafeSQL(dtr("Fax No_").ToString) & ", email = " & SafeSQL(dtr("E-mail").ToString) & ", Website = " & SafeSQL(dtr("Home Page").ToString) & ", ICPartner = " & SafeSQL("") & ", " & _
                                    "PriceGroup= " & SafeSQL(dtr("Customer Price Group").ToString) & ", PaymentTerms = " & SafeSQL(dtr("Payment Terms Code").ToString) & ", PaymentMethod = " & SafeSQL(dtr("Payment Method Code").ToString) & ", CustDiscGroup = " & SafeSQL(dtr("Customer Disc_ Group").ToString) & ", ShipmentMethod = " & SafeSQL(dtr("Shipment Method Code").ToString) & ", SalesAgent = " & SafeSQL(dtr("Salesperson Code").ToString) & ", ShipAgent = " & SafeSQL(dtr("Shipping Agent Code").ToString) & ", Category = " & SafeSQL("") & ", Dimension1 = " & SafeSQL(dtr("Global Dimension 1 Filter").ToString) & ", Dimension2 = " & SafeSQL(dtr("Global Dimension 2 Filter").ToString) & ", [Bill-toNo] = " & SafeSQL(sBillto) & ", InvDisGroup = " & SafeSQL(dtr("Invoice Disc_ Code").ToString) & ", Location = " & SafeSQL(dtr("Location Code").ToString) & ", CurrencyCode = " & SafeSQL(dtr("Currency Code").ToString) & ", Active = 1 Where CustNo = " & SafeSQL(dtr("No_").ToString))
                ElseIf iBranch = 2 Then
                    Dim iCustName As Integer = UpdateCustName(sdCustNo)

                    If iCustName = 1 Then
                        ExecuteSQL("Update Customer Set CustName=" & SafeSQL(dtr("Name").ToString) & ",SearchName=" & SafeSQL(dtr("Search Name").ToString) & ", CountryCode = " & SafeSQL(dtr("Country Code").ToString) & ", Phone = " & SafeSQL(dtr("Phone No_").ToString) & ", CustDiscGroup = " & SafeSQL(dtr("Customer Disc_ Group").ToString) & ", ContactPerson = " & SafeSQL(dtr("Contact").ToString) & ", Balance = " & dtr("Balance").ToString & ", CreditLimit = " & dtr("Credit Limit (LCY)").ToString & ", ProvisionalBalance = 0, ZoneCode = " & SafeSQL(dtr("Customer Disc_ Group").ToString) & ", FaxNo = " & SafeSQL(dtr("Fax No_").ToString) & ", Email = " & SafeSQL(dtr("E-mail").ToString) & ", Website = " & SafeSQL(dtr("Home Page").ToString) & ", ICPartner = '', PriceGroup = " & SafeSQL(dtr("Customer Price Group").ToString) & ", PaymentTerms = " & SafeSQL(dtr("Payment Terms Code").ToString) & ", PaymentMethod = " & SafeSQL(dtr("Payment Method Code").ToString) & ", ShipmentMethod = " & SafeSQL(dtr("Shipment Method Code").ToString) & ", SalesAgent = " & SafeSQL(dtr("Salesperson Code").ToString) & ", ShipAgent = " & SafeSQL(dtr("Shipping Agent Code").ToString) & ", Category = '', Dimension1 = '', Dimension2 = '', [Bill-toNo] = " & SafeSQL(sBillto) & ", InvDisGroup = " & SafeSQL(dtr("Invoice Disc_ Code").ToString) & ", Location = " & SafeSQL(dtr("Location Code").ToString) & ", CurrencyCode = " & SafeSQL(dtr("Currency Code").ToString) & ", Active = 1, CustPostGroup = '', CustType = 'Customer', Photo = '', MemberType = '', GSTCustGroup = '', Exported = 1, SupplierCode = '', ShipName = '', ShipAddr = " & SafeSQL(dtr("Address").ToString) & ", ShipAddr2 = " & SafeSQL(dtr("Address 2").ToString) & ", ShipAddr3 = '', ShipAddr4 = '', ShipPost = " & SafeSQL(dtr("Post Code").ToString) & ", ShipCity = " & SafeSQL(dtr("City").ToString) & ", [Print] = 'Invoice',GSTType = 'Inclusive', Basket = 0, MDTNo = '', ExportExclude = 0, Remarks = '', CommissionCode = '', Rebate = '' Where AcCustCode = " & SafeSQL(dtr("No_").ToString))
                    Else
                        ExecuteSQL("Update Customer Set CountryCode = " & SafeSQL(dtr("Country Code").ToString) & ", Phone = " & SafeSQL(dtr("Phone No_").ToString) & ", ContactPerson = " & SafeSQL(dtr("Contact").ToString) & ", CustDiscGroup = " & SafeSQL(dtr("Customer Disc_ Group").ToString) & ", Balance = " & dtr("Balance").ToString & ", CreditLimit = " & dtr("Credit Limit (LCY)").ToString & ", ProvisionalBalance = 0, ZoneCode = " & SafeSQL(dtr("Customer Disc_ Group").ToString) & ", FaxNo = " & SafeSQL(dtr("Fax No_").ToString) & ", Email = " & SafeSQL(dtr("E-mail").ToString) & ", Website = " & SafeSQL(dtr("Home Page").ToString) & ", ICPartner = '', PriceGroup = " & SafeSQL(dtr("Customer Price Group").ToString) & ", PaymentTerms = " & SafeSQL(dtr("Payment Terms Code").ToString) & ", PaymentMethod = " & SafeSQL(dtr("Payment Method Code").ToString) & ", ShipmentMethod = " & SafeSQL(dtr("Shipment Method Code").ToString) & ", SalesAgent = " & SafeSQL(dtr("Salesperson Code").ToString) & ", ShipAgent = " & SafeSQL(dtr("Shipping Agent Code").ToString) & ", Category = '', Dimension1 = '', Dimension2 = '', [Bill-toNo] = " & SafeSQL(sBillto) & ", InvDisGroup = " & SafeSQL(dtr("Invoice Disc_ Code").ToString) & ", Location = " & SafeSQL(dtr("Location Code").ToString) & ", CurrencyCode = " & SafeSQL(dtr("Currency Code").ToString) & ", Active = 1, CustPostGroup = '', CustType = 'Customer', Photo = '', MemberType = '', GSTCustGroup = '', Exported = 1, SupplierCode = '', ShipName = '', ShipAddr = " & SafeSQL(dtr("Address").ToString) & ", ShipAddr2 = " & SafeSQL(dtr("Address 2").ToString) & ", ShipAddr3 = '', ShipAddr4 = '', ShipPost = " & SafeSQL(dtr("Post Code").ToString) & ", ShipCity = " & SafeSQL(dtr("City").ToString) & ", [Print] = 'Invoice',GSTType = 'Inclusive', Basket = 0, MDTNo = '', ExportExclude = 0, Remarks = '', CommissionCode = '', Rebate = '' Where AcCustCode = " & SafeSQL(dtr("No_").ToString))
                    End If
                End If
            End If
            icnt += 1
        End While
        'dtr.Close()
        Dim dtrCust As SqlDataReader
        Dim arr As New ArrayList
        'Remove duplicate customer id
        dtrCust = ReadRecord("Select * from Customer order by CustNo")
        Dim sCust As String = ""
        While dtrCust.Read
            If sCust = dtrCust("CustNo") Then
                Dim ds As DelCust
                ds.CustID = dtrCust("CustNo")
                ds.PrGroup = dtrCust("PriceGroup")
                arr.Add(ds)
            End If
            sCust = dtrCust("CustNo")
        End While
        dtrCust.Close()
        'dtr = ReadNavRecord("Select * from Customer where Blocked <> 0")
        '' where No_ not in (Select ""Customer No_"" from ""Ship-to Address"") and ""Location code"" like 'WHSE-1'")
        'While dtr.Read
        '    ExecuteSQL("Update Customer set Active =0 where CustNo=" & SafeSQL(dtr("No_").ToString))
        'End While
        'dtr.Close()
        For iindex As Integer = 0 To arr.Count - 1
            Dim ds As DelCust
            ds = arr(iindex)
            '        ExecuteSQL("Delete from Customer Where CustNo =" & SafeSQL(ds.CustID) & " and PriceGroup = " & SafeSQL(ds.PrGroup))
        Next
        ExecuteSQL("Update Customer Set ShipName = CustName where (ShipName='' or ShipName is Null)")
        ExecuteSQL("Update Customer Set Active=0 where CustName like '%Closed%'")

        'ExecuteSQL("Update Customer Set Active=0 where CustName like '%Closed%'")
        ExecuteSQL("Update Customer Set SearchName = CustName where (SearchName='' or SearchName is Null)")
        ''MsgBox(Format(dt, "mm:ss") & vbCrLf & Format(Date.Now, "mm:ss"))
    End Sub

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

    Public Sub ImportCustomerEH()
        Dim dt As DateTime
        Dim dtr As OdbcDataReader
        dt = Date.Now
        ExecuteSQL("Update Customer set Active = 0")
        dtr = ReadNavRecord("Select * from Customer")
        Dim icnt As Integer = 1
        While dtr.Read
            If dtr("No_").ToString <> "" Then
                If IsCustomerExists(dtr("No_").ToString) Then
                    If dtr("Customer Type").ToString = "Customer" Then
                        ExecuteSQL("Update Customer Set CustName = " & SafeSQL(dtr("Name").ToString) & ", ChineseName = " & SafeSQL("") & ", Address = " & SafeSQL(dtr("Address").ToString) & ", Address2 = " & SafeSQL(dtr("Address 2").ToString) & ", Address3 = " & SafeSQL(dtr("Address 3").ToString) & ", Address4 = " & SafeSQL(dtr("Address 4").ToString) & ", PostCode = " & SafeSQL(dtr("Post Code").ToString) & ", City = " & SafeSQL(dtr("City").ToString) & ", CountryCode = " & SafeSQL(dtr("Country Code").ToString) & ", Phone = " & SafeSQL(dtr("Phone No_").ToString) & ", ContactPerson = " & SafeSQL(dtr("Contact").ToString) & ", Balance = " & dtr("Balance").ToString & ", CreditLimit = " & dtr("Credit Limit (LCY)").ToString & ", ZoneCode = " & SafeSQL(dtr("Service Zone Code").ToString) & ", FaxNo = " & SafeSQL(dtr("Fax No_").ToString) & ", email = " & SafeSQL(dtr("E-mail").ToString) & ", Website = " & SafeSQL(dtr("Home Page").ToString) & ", ICPartner = " & SafeSQL(dtr("IC Partner Code").ToString) & ", " & _
                                        "PriceGroup= " & SafeSQL(dtr("Customer Price Group").ToString) & ", PaymentTerms = " & SafeSQL(dtr("Payment Terms Code").ToString) & ", PaymentMethod = " & SafeSQL(dtr("Payment Method Code").ToString) & ", ShipmentMethod = " & SafeSQL(dtr("Shipment Method Code").ToString) & ", SalesAgent = " & SafeSQL(dtr("Salesperson Code").ToString) & ", ShipAgent = " & SafeSQL(dtr("Shipping Agent Code").ToString) & ", Category = " & SafeSQL(dtr("IC Partner Code").ToString) & ", Dimension1 = " & SafeSQL(dtr("Global Dimension 1 Filter").ToString) & ", Dimension2 = " & SafeSQL(dtr("Global Dimension 2 Filter").ToString) & ", [Bill-toNo] = " & SafeSQL("") & ", InvDisGroup = " & SafeSQL(dtr("Invoice Disc_ Code").ToString) & ", Location = " & SafeSQL(dtr("Location Code").ToString) & ", CurrencyCode = " & SafeSQL(dtr("Currency Code").ToString) & ", Active = 1, Exported = 1 Where CustNo = " & SafeSQL(dtr("No_").ToString))
                    Else
                        ExecuteSQL("Update Customer Set CustName = " & SafeSQL(dtr("Name").ToString) & ", ChineseName = " & SafeSQL("") & ", Address = " & SafeSQL(dtr("Address").ToString) & ", Address2 = " & SafeSQL(dtr("Address 2").ToString) & ", Address3 = " & SafeSQL(dtr("Address 3").ToString) & ", Address4 = " & SafeSQL(dtr("Address 4").ToString) & ", PostCode = " & SafeSQL(dtr("Post Code").ToString) & ", City = " & SafeSQL(dtr("City").ToString) & ", CountryCode = " & SafeSQL(dtr("Country Code").ToString) & ", Phone = " & SafeSQL(dtr("Phone No_").ToString) & ", ContactPerson = " & SafeSQL(dtr("Contact").ToString) & ", Balance = " & dtr("Balance").ToString & ", CreditLimit = " & dtr("Credit Limit (LCY)").ToString & ", ZoneCode = " & SafeSQL(dtr("Service Zone Code").ToString) & ", FaxNo = " & SafeSQL(dtr("Fax No_").ToString) & ", email = " & SafeSQL(dtr("E-mail").ToString) & ", Website = " & SafeSQL(dtr("Home Page").ToString) & ", ICPartner = " & SafeSQL(dtr("IC Partner Code").ToString) & ", " & _
                                                                "PriceGroup= " & SafeSQL(dtr("Customer Price Group").ToString) & ", PaymentTerms = " & SafeSQL(dtr("Payment Terms Code").ToString) & ", PaymentMethod = " & SafeSQL(dtr("Payment Method Code").ToString) & ", ShipmentMethod = " & SafeSQL(dtr("Shipment Method Code").ToString) & ", SalesAgent = " & SafeSQL(dtr("Salesperson Code").ToString) & ", ShipAgent = " & SafeSQL(dtr("Shipping Agent Code").ToString) & ", Category = " & SafeSQL(dtr("IC Partner Code").ToString) & ", Dimension1 = " & SafeSQL(dtr("Global Dimension 1 Filter").ToString) & ", Dimension2 = " & SafeSQL(dtr("Global Dimension 2 Filter").ToString) & ", [Bill-toNo] = " & SafeSQL(dtr("Bill-to Customer No_").ToString) & ", InvDisGroup = " & SafeSQL(dtr("Invoice Disc_ Code").ToString) & ", Location = " & SafeSQL(dtr("Location Code").ToString) & ", CurrencyCode = " & SafeSQL(dtr("Currency Code").ToString) & ", ExpiryDate = " & SafeSQL(Format(dtr("Expiry Date"), "yyyyMMdd")) & ", Active = 1, Exported = 1 Where CustNo = " & SafeSQL(dtr("No_").ToString))
                    End If
                Else
                    If dtr("Customer Type").ToString = "Customer" Then
                        ExecuteSQL("Insert into Customer(CustNo, CustName, ChineseName, Address, Address2, Address3, Address4, PostCode, City, CountryCode, Phone, ContactPerson, Balance, CreditLimit, ProvisionalBalance, ZoneCode, FaxNo, Email, Website, ICPartner, PriceGroup, PaymentTerms, PaymentMethod, ShipmentMethod, SalesAgent, ShipAgent, Category, Dimension1, Dimension2, [Bill-toNo], InvDisGroup, Location, CurrencyCode, Active, CustPostGroup, DisplayNo, Basket, CustType, Photo, Exported) Values (" & SafeSQL(dtr("No_").ToString) & "," & SafeSQL(dtr("Name").ToString) & "," & SafeSQL("") & "," & SafeSQL(dtr("Address").ToString) & ", " & SafeSQL(dtr("Address 2").ToString) & ", " & SafeSQL(dtr("Address 3").ToString) & ", " & SafeSQL(dtr("Address 4").ToString) & "," & _
                                                            SafeSQL(dtr("Post Code").ToString) & "," & SafeSQL(dtr("City").ToString) & "," & SafeSQL(dtr("Country Code").ToString) & "," & SafeSQL(dtr("Phone No_").ToString) & "," & SafeSQL(dtr("Contact").ToString) & ", " & dtr("Balance").ToString & ", " & dtr("Credit Limit (LCY)").ToString & ", 0," & SafeSQL(dtr("Service Zone Code").ToString) & ", " & SafeSQL(dtr("Fax No_").ToString) & ", " & SafeSQL(dtr("E-mail").ToString) & ", " & SafeSQL(dtr("Home Page").ToString) & ", " & SafeSQL(dtr("IC Partner Code").ToString) & ", " & SafeSQL(dtr("Customer Price Group").ToString) & _
                                                            ", " & SafeSQL(dtr("Payment Terms Code").ToString) & ", " & SafeSQL(dtr("Payment Method Code").ToString) & ", " & SafeSQL(dtr("Shipment Method Code").ToString) & ", " & SafeSQL(dtr("Salesperson Code").ToString) & ", " & SafeSQL(dtr("Shipping Agent Code").ToString) & ", " & SafeSQL(dtr("IC Partner Code").ToString) & ", " & SafeSQL(dtr("Global Dimension 1 Filter").ToString) & ", " & SafeSQL(dtr("Global Dimension 2 Filter").ToString) & "," & SafeSQL(dtr("Bill-to Customer No_").ToString) & ", " & SafeSQL(dtr("Invoice Disc_ Code").ToString) & ", " & SafeSQL(dtr("Location Code").ToString) & ", " & SafeSQL(dtr("Currency Code").ToString) & ",1,''," & icnt & ",0,'Customer','', 1)")
                    Else
                        ExecuteSQL("Insert into Customer(CustNo, CustName, ChineseName, Address, Address2, Address3, Address4, PostCode, City, CountryCode, Phone, ContactPerson, Balance, CreditLimit, ProvisionalBalance, ZoneCode, FaxNo, Email, Website, ICPartner, PriceGroup, PaymentTerms, PaymentMethod, ShipmentMethod, SalesAgent, ShipAgent, Category, Dimension1, Dimension2, [Bill-toNo], InvDisGroup, Location, CurrencyCode, Active, CustPostGroup, DisplayNo, Basket, CustType, Photo, ExpiryDate, Exported) Values (" & SafeSQL(dtr("No_").ToString) & "," & SafeSQL(dtr("Name").ToString) & "," & SafeSQL("") & "," & SafeSQL(dtr("Address").ToString) & ", " & SafeSQL(dtr("Address 2").ToString) & ", " & SafeSQL(dtr("Address 3").ToString) & ", " & SafeSQL(dtr("Address 4").ToString) & "," & _
                                    SafeSQL(dtr("Post Code").ToString) & "," & SafeSQL(dtr("City").ToString) & "," & SafeSQL(dtr("Country Code").ToString) & "," & SafeSQL(dtr("Phone No_").ToString) & "," & SafeSQL(dtr("Contact").ToString) & ", " & dtr("Balance").ToString & ", " & dtr("Credit Limit (LCY)").ToString & ", 0," & SafeSQL(dtr("Service Zone Code").ToString) & ", " & SafeSQL(dtr("Fax No_").ToString) & ", " & SafeSQL(dtr("E-mail").ToString) & ", " & SafeSQL(dtr("Home Page").ToString) & ", " & SafeSQL(dtr("IC Partner Code").ToString) & ", " & SafeSQL(dtr("Customer Price Group").ToString) & _
                                    ", " & SafeSQL(dtr("Payment Terms Code").ToString) & ", " & SafeSQL(dtr("Payment Method Code").ToString) & ", " & SafeSQL(dtr("Shipment Method Code").ToString) & ", " & SafeSQL(dtr("Salesperson Code").ToString) & ", " & SafeSQL(dtr("Shipping Agent Code").ToString) & ", " & SafeSQL(dtr("IC Partner Code").ToString) & ", " & SafeSQL(dtr("Global Dimension 1 Filter").ToString) & ", " & SafeSQL(dtr("Global Dimension 2 Filter").ToString) & "," & SafeSQL(dtr("Bill-to Customer No_").ToString) & ", " & SafeSQL(dtr("Invoice Disc_ Code").ToString) & ", " & SafeSQL(dtr("Location Code").ToString) & ", " & SafeSQL(dtr("Currency Code").ToString) & ",1,''," & icnt & ",0,'Member', ''," & SafeSQL(Format(dtr("Expiry Date"), "yyyyMMdd")) & ", 1)")
                    End If
                End If
            End If
            icnt += 1
        End While
        dtr.Close()
    End Sub

    Public Sub ImportItemVariant()
        Dim dtr As OdbcDataReader
        ExecuteSQL("Delete from ItemVariant")
        dtr = ReadNavRecord("Select * from ""Item Variant""")
        While dtr.Read
            If dtr("Code").ToString <> "" Then
                ExecuteSQL("Insert into ItemVariant(ItemNo, Code, Description) Values (" & SafeSQL(dtr("Item No_").ToString) & "," & SafeSQL(dtr("Code").ToString) & "," & SafeSQL(dtr("Description").ToString) & ")")
            End If
        End While
        dtr.Close()
    End Sub

    Public Sub ImportReservationQty()
        Dim dtr As OdbcDataReader
        ExecuteSQL("Delete from Reservation")
        dtr = ReadNavRecord("Select * from ""Reservation Entry""")
        While dtr.Read
            If dtr("Qty_ to Handle (Base)") > 0 Then
                ExecuteSQL("Insert into Reservation(ItemNo, Location, VariantCode, Qty, ReserveType, ShipmentDate, ExpectedRcptDate) Values (" & SafeSQL(dtr("Item No_").ToString) & "," & SafeSQL(dtr("Location Code").ToString) & "," & SafeSQL(dtr("Variant Code").ToString) & "," & dtr("Quantity (Base)") & "," & SafeSQL(dtr("Reservation Status").ToString) & "," & SafeSQL(Format(dtr("Shipment Date"), "yyyyMMdd")) & "," & SafeSQL(Format(dtr("Expected Receipt Date"), "yyyyMMdd")) & ")")
            End If
        End While
        dtr.Close()
    End Sub

    Public Sub ImportPOLine()
        Dim dtr As OdbcDataReader
        ExecuteSQL("Delete from POLine")
        dtr = ReadNavRecord("Select * from ""Purchase Line""")
        While dtr.Read
            If dtr("Outstanding Qty_ (Base)") > 0 Then
                ExecuteSQL("Insert into POLine(ItemNo, Location, VariantCode, Qty, PlannedRcptDate) Values (" & SafeSQL(dtr("No_").ToString) & "," & SafeSQL(dtr("Location Code").ToString) & "," & SafeSQL(dtr("Variant Code").ToString) & "," & dtr("Outstanding Qty_ (Base)") & "," & SafeSQL(Format(dtr("Planned Receipt Date"), "yyyyMMdd")) & ")")
            End If
        End While
        dtr.Close()
    End Sub

    Public Sub ImportLocation()
        Dim dtr As OdbcDataReader
        ExecuteSQL("Delete from Location")
        dtr = ReadNavRecord("Select * from Location")
        While dtr.Read
            If dtr("Code").ToString <> "" Then
                ExecuteSQL("Insert into Location(Code, Name) Values (" & SafeSQL(dtr("Code").ToString) & "," & SafeSQL(dtr("Name").ToString) & ")")
            End If
        End While
        dtr.Close()
    End Sub

    Public Sub ImportSalesAgent()
        Dim dtr As OdbcDataReader
        ExecuteSQL("Update SalesAgent Set Active = 0")
        dtr = ReadNavRecord("Select * from ""Salesperson/Purchaser""")
        While dtr.Read
            If dtr("Code").ToString <> "" Then
                If IsAgentExists(dtr("Code").ToString) = False Then
                    ExecuteSQL("Insert into SalesAgent(Code, Name, Phone,Email, Password, UserID, Access, Active, SalesTarget) Values (" & SafeSQL(dtr("Code").ToString) & "," & SafeSQL(dtr("Name").ToString) & "," & SafeSQL(dtr("Phone No_").ToString) & "," & SafeSQL(dtr("E-Mail").ToString) & "," & SafeSQL(dtr("Code").ToString) & "," & SafeSQL(dtr("Code").ToString) & ",1,1,0)")
                Else
                    ExecuteSQL("Update SalesAgent Set Name = " & SafeSQL(dtr("Name").ToString) & ", Phone =" & SafeSQL(dtr("Phone No_").ToString) & ", Email=" & SafeSQL(dtr("E-Mail").ToString) & ", Active = 1 Where Code = " & SafeSQL(dtr("Code").ToString))
                End If
            End If
        End While
        dtr.Close()
    End Sub

    Public Sub ImportPaymentTerms()

    End Sub

    Public Sub ImportItemPricePromotion()
        Dim dtr As OdbcDataReader
        Dim LastPromoNo As Double = 0
        Dim PrePromoNo As String = ""
        Dim LenPromoNo As Integer = 0
        Dim sPrNo As String = ""
        'ExecuteSQL("Delete from ItemPr")
        ExecuteSQL("Delete from PromoOffer Where PromoId in (Select PromoId from Promotion where PromoName = 'Special Nav Price')")
        ExecuteSQL("Delete from PromoCondition Where PromoId in (Select PromoId from Promotion where PromoName = 'Special Nav Price')")
        ExecuteSQL("Delete from PromoApply Where PromoId in (Select PromoId from Promotion where PromoName = 'Special Nav Price')")
        ExecuteSQL("Delete from Promotion where PromoName = 'Special Nav Price'")
        Dim dtrSQL As SqlDataReader
        Dim sMDT As String = ""
        dtrSQL = ReadRecord("Select MDT.MDTNO, LastPrNo, PrePrNo, LenPrNo from MDT, System Where MDT.MDTNO = System.MDTNO")
        If dtrSQL.Read Then
            LastPromoNo = dtrSQL("LastPrNo")
            PrePromoNo = dtrSQL("PrePrNo")
            LenPromoNo = dtrSQL("LenPrNo")
            sMDT = dtrSQL("MDTNo").ToString
        End If
        dtrSQL.Close()
        'Dim sVariantCode As String
        dtr = ReadNavRecord("Select * from ""Sales Price""")
        While dtr.Read
            LastPromoNo = LastPromoNo + 1
            sPrNo = PrePromoNo & StrDup(LenPromoNo - Len(PrePromoNo) - Len(CStr(LastPromoNo)), "0") & CStr(LastPromoNo)
            ImportNavSpecialPrice(sPrNo, dtr("Sales Code"), dtr("FOC Item No_").ToString, dtr("Unit of Measure Code").ToString, dtr("FOC Quantity"), Date.Now, DateAdd(DateInterval.Month, 1, Date.Now), dtr("Minimum Quantity"), dtr("Maximum Quantity"))
        End While
        dtr.Close()
        ExecuteSQL("Update MDT Set LastPrNo = " & LastPromoNo & " where MDTNO = " & SafeSQL(sMDT))
    End Sub

    Private Sub ImportNavSpecialPrice(ByVal PromoNo As String, ByVal PriceGroup As String, ByVal ItemNo As String, ByVal UOM As String, ByVal SpPrice As Double, ByVal FromDate As Date, ByVal ToDate As Date, ByVal MinQty As Double, ByVal MaxQty As Double)
        ExecuteSQL("Insert into Promotion(PromoID, PromoName, APType, FromDate, ToDate,Multiply,PromoType, Priority, Entitle, EntitleType, CATBased, ItemCondition,Event, Attachment, AllItems) Values (" & SafeSQL(PromoNo) & ",'Special Nav Price','Price Group'," & SafeSQL(FromDate) & "," & SafeSQL(ToDate) & ", 'Standard','Item Promotion',0,99,'Per Invoice',0,0,'','',0)")
        ExecuteSQL("Insert into PromoCondition(PromoID, ItemId, UOM, MinQty, MaxQty, LineType, IsRequired) Values (" & SafeSQL(PromoNo) & "," & SafeSQL(ItemNo) & "," & SafeSQL(UOM) & "," & MinQty & "," & MaxQty & ",'Item','Required')")
        ExecuteSQL("Insert into PromoOffer(PromoID, ItemId, UOM, FocQty, DisPrice, Discount, LineType) Values (" & SafeSQL(PromoNo) & "," & SafeSQL(ItemNo) & "," & SafeSQL(UOM) & ", " & SpPrice & ", 0, 0, 'Item')")
        ExecuteSQL("Insert into PromoApply(PromoID, Id) Values (" & SafeSQL(PromoNo) & "," & SafeSQL(PriceGroup) & ")")
    End Sub

    Public Sub ImportCustDiscPromotion()
        Dim dtr As OdbcDataReader
        Dim LastPromoNo As Double = 0
        Dim PrePromoNo As String = ""
        Dim LenPromoNo As Integer = 0
        Dim sPrNo As String = ""
        'ExecuteSQL("Delete from ItemPr")
        'MsgBox("1")
        ExecuteSQL("Delete from PromoOffer Where PromoId in (Select PromoId from Promotion where PromoName = 'Special Disc Promotion')")
        ExecuteSQL("Delete from PromoCondition Where PromoId in (Select PromoId from Promotion where PromoName = 'Special Disc Promotion')")
        ExecuteSQL("Delete from PromoApply Where PromoId in (Select PromoId from Promotion where PromoName = 'Special Disc Promotion')")
        ExecuteSQL("Delete from Promotion where PromoName = 'Special Disc Promotion'")
        'MsgBox("2")
        Dim dtrSQL As SqlDataReader
        Dim sMDT As String = ""
        dtrSQL = ReadRecord("Select MDT.MDTNO, LastPrNo, PrePrNo, LenPrNo from MDT, System Where MDT.MDTNO = System.MDTNO")
        If dtrSQL.Read Then
            '   MsgBox("3")
            LastPromoNo = dtrSQL("LastPrNo")
            '  MsgBox(LastPromoNo)
            PrePromoNo = dtrSQL("PrePrNo").ToString
            ' MsgBox(PrePromoNo)
            LenPromoNo = dtrSQL("LenPrNo")
            'MsgBox(LenPromoNo)
            sMDT = dtrSQL("MDTNo").ToString
            'MsgBox(sMDT)
        End If
        dtrSQL.Close()
        'MsgBox("3")
        'Dim sVariantCode As String
        Dim sCode As String = ""
        Dim sUOM As String = ""
        Dim sType As String = ""
        Dim dDis As Double = 0
        Dim dMin As Double = 0
        Dim arr As New ArrayList
        dtr = ReadNavRecord("Select * from ""Sales Line Discount"" order by Code, ""Unit Of Measure Code"", ""Sales Type"", ""Minimum Quantity"", ""Line Discount %""")
        'MsgBox("1")
        While dtr.Read
            '         MsgBox(dtr("Code").ToString & "," & dtr("Unit Of Measure Code").ToString & "," & dtr("Line Discount %") & "," & dtr("Minimum Quantity") & "," & dtr("Sales Type").ToString)
            If sCode <> dtr("Code").ToString Or sUOM <> dtr("Unit Of Measure Code").ToString Or dDis <> dtr("Line Discount %") Or sType <> dtr("Sales Type").ToString Or dMin <> dtr("Minimum Quantity") Then
                '        MsgBox("1")
                If sCode <> "" Then
                    '           MsgBox("2")
                    LastPromoNo = LastPromoNo + 1
                    '          MsgBox("3")
                    sPrNo = PrePromoNo & StrDup(LenPromoNo - Len(PrePromoNo) - Len(CStr(LastPromoNo)), "0") & CStr(LastPromoNo)
                    '           MsgBox("4")
                    'If sType = "Customer Disc. Group" Then sType = "Zone"
                    '        MsgBox("5")
                    ImportNavSpecialDiscount(sPrNo, sCode, sUOM, dDis, Date.Now, DateAdd(DateInterval.Month, 1, Date.Now), dMin, 9999, sType)
                    '       MsgBox("6")
                    For iIndex As Integer = 0 To arr.Count - 1
                        ImportNavSpecialDiscountApply(sPrNo, arr(iIndex))
                    Next
                    '      MsgBox("7")
                End If
                ' MsgBox("2: " & dtr("Line Discount %") & "," & dtr("Minimum Quantity"))
                sCode = dtr("Code").ToString
                sUOM = dtr("Unit Of Measure Code").ToString
                dDis = dtr("Line Discount %")
                If dtr("Sales Type") = 0 Then
                    sType = "Customer"
                Else
                    sType = "Zone"
                End If
                dMin = dtr("Minimum Quantity")
                arr.Clear()
            End If
            arr.Add(dtr("Sales Code").ToString)
        End While
        dtr.Close()
        'ExecuteSQL("Delete from Itempr where PriceGroup not in (Select PriceGroup from customer where Active=1) and SalesType='Customer Price Group'")
        ExecuteSQL("Update MDT Set LastPrNo = " & LastPromoNo & " where MDTNO = " & SafeSQL(sMDT))
    End Sub
    Private Sub ImportNavSpecialDiscount(ByVal PromoNo As String, ByVal ItemNo As String, ByVal UOM As String, ByVal SpDisc As Double, ByVal FromDate As Date, ByVal ToDate As Date, ByVal MinQty As Double, ByVal MaxQty As Double, ByVal SalesType As String)
        ExecuteSQL("Insert into Promotion(PromoID, PromoName, APType, FromDate, ToDate,Multiply,PromoType, Priority, Entitle, EntitleType, CATBased, ItemCondition,Event, Attachment, AllItems) Values (" & SafeSQL(PromoNo) & ",'Special Disc Promotion', " & SafeSQL(SalesType) & "," & SafeSQL(FromDate) & "," & SafeSQL(ToDate) & ", 'Standard','Item Promotion',0,9999,'Per Month',0,0,'','',0)")
        ExecuteSQL("Insert into PromoCondition(PromoID, ItemId, UOM, MinQty, MaxQty, LineType, IsRequired) Values (" & SafeSQL(PromoNo) & "," & SafeSQL(ItemNo) & "," & SafeSQL(UOM) & "," & MinQty & "," & MaxQty & ",'Item','Required')")
        ExecuteSQL("Insert into PromoOffer(PromoID, ItemId, UOM, FocQty, DisPrice, Discount, LineType) Values (" & SafeSQL(PromoNo) & "," & SafeSQL(ItemNo) & "," & SafeSQL(UOM) & ",0, 0," & SpDisc & ",'Item')")
        'ExecuteSQL("Insert into PromoApply(PromoID, Id) Values (" & SafeSQL(PromoNo) & "," & SafeSQL(Zone) & ")")
    End Sub

    Private Sub ImportNavSpecialDiscountApply(ByVal PromoNo As String, ByVal Zone As String)
        ExecuteSQL("Insert into PromoApply(PromoID, Id) Values (" & SafeSQL(PromoNo) & "," & SafeSQL(Zone) & ")")
    End Sub

    Public Sub ImportBOM()
        Dim dtr As OdbcDataReader
        Dim arrList As New ArrayList
        ExecuteSQL("Delete from BOM")
        dtr = ReadNavRecord("Select ""No_"", Description,""Unit of Measure Code"" from ""Production BOM Header""")
        While dtr.Read
            If dtr("No_").ToString = "" Then
            Else
                ExecuteSQL("Insert into BOM (BOMNo, Description, ItemNo, UOM, Active) Values (" & SafeSQL(dtr("No_").ToString) & "," & SafeSQL(dtr("Description").ToString) & ",''," & SafeSQL(dtr("Unit of Measure Code").ToString) & ",1)")
                arrList.Add(dtr("No_"))
            End If
        End While
        dtr.Close()
        ImportBOMDet(arrList)
    End Sub
    Public Sub ImportItemPrice()
        'MsgBox("ItemPr")
        Dim dtr As OdbcDataReader
        Dim arrList As New ArrayList
        ExecuteSQL("Delete from ItemPr")
        'dtr = ReadNavRecord("Select A.""Item No_"", A.""Sales Code"",Max(A.""Starting Date"") as MaxDate, A.""Sales Type"", A.""Minimum Quantity"" from ""Sales Price""  A, Customer B where B.""Customer Price Group"" = A.""Sales Code"" group by A.""Item No_"", A.""Sales Code"", A.""Sales Type"", A.""Minimum Quantity""")
        dtr = ReadNavRecord("Select ""Item No_"", ""Sales Code"",Max(""Starting Date"") as MaxDate, ""Sales Type"", ""Minimum Quantity"", ""Unit of Measure Code"", ""Variant Code"" from ""Sales Price""  group by ""Item No_"", ""Sales Code"", ""Sales Type"",""Minimum Quantity"", ""Unit of Measure Code"", ""Variant Code""")
        While dtr.Read
            If dtr("Item No_").ToString = "" Or dtr("Sales Code").ToString = "" Then
            Else
                'MsgBox(dtr("Sales Type"))
                '        If IsPriceGroupExists(dtr("Sales Code").ToString) Then
                Dim aPr As ArrItemPrice
                aPr.ItemCode = dtr("Item No_").ToString
                aPr.MaxDate = dtr("MaxDate")
                aPr.SalesCode = dtr("Sales Code").ToString
                aPr.SalesType = dtr("Sales Type")
                aPr.MinQty = dtr("Minimum Quantity")
                aPr.VariantCode = dtr("Variant Code").ToString
                aPr.sUOM = dtr("Unit of Measure Code").ToString
                arrList.Add(aPr)
                'End If
                'If IsCustExists(dtr("Sales Code").ToString) Then
                '    Dim aPr As ArrItemPrice
                '    aPr.ItemCode = dtr("Item No_").ToString
                '    aPr.MaxDate = dtr("MaxDate")
                '    aPr.SalesCode = dtr("Sales Code").ToString
                '    aPr.SalesType = dtr("Sales Type")
                '    aPr.MinQty = dtr("Minimum Quantity")
                '    aPr.VariantCode = dtr("Variant Code").ToString
                '    aPr.sUOM = dtr("Unit of Measure Code").ToString
                '    arrList.Add(aPr)
                'End If
            End If
        End While
        '   MsgBox("1")
        'dtr.Close()
        ImportItemPrDet(arrList)
        'ExecuteSQL("Update ItemPr set SalesType='Customer Price Group' where SalesType='1'")
        'ExecuteSQL("Update ItemPr set SalesType='Customer' where SalesType='0'")
        'ExecuteSQL("Update ItemPr set SalesType='All Customer' where SalesType='2'")
        ExecuteSQL("Delete from ItemPr where UnitPrice=0")
    End Sub
    Private Function IsPriceGroupExists(ByVal sCode As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = ReadRecord("Select PriceGroup from Customer where PriceGroup = " & SafeSQL(sCode))
        bAns = dtr.Read
        dtr.Close()
        Return bAns
    End Function
    Private Function IsCustExists(ByVal sCode As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = ReadRecord("Select CustNo from Customer where CustNo = " & SafeSQL(sCode))
        bAns = dtr.Read
        dtr.Close()
        Return bAns
    End Function
    Private Function IsCustDiscGroupExists(ByVal sCode As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = ReadRecord("Select CustDiscGroup from Customer where CustDiscGroup = " & SafeSQL(sCode))
        bAns = dtr.Read
        dtr.Close()
        Return bAns
    End Function
    Public Sub ImportCustDiscItemGroup()
        Dim dtr As OdbcDataReader
        Dim arrList As New ArrayList
        ExecuteSQL("Delete from CustDiscGroupItem")
        'dtr = ReadNavRecord("Select A.""Item No_"", A.""Sales Code"",Max(A.""Starting Date"") as MaxDate, A.""Sales Type"", A.""Minimum Quantity"" from ""Sales Price""  A, Customer B where B.""Customer Price Group"" = A.""Sales Code"" group by A.""Item No_"", A.""Sales Code"", A.""Sales Type"", A.""Minimum Quantity""")
        dtr = ReadNavRecord("Select ""Code"", ""Sales Code"",Max(""Starting Date"") as MaxDate, ""Sales Type"", ""Minimum Quantity"", ""Unit of Measure Code"", ""Variant Code"" from ""Sales Line Discount""  group by ""Code"", ""Sales Code"", ""Sales Type"",""Minimum Quantity"", ""Unit of Measure Code"", ""Variant Code""")
        While dtr.Read
            If dtr("Code").ToString = "" Or dtr("Sales Code").ToString = "" Then
            Else
                If IsCustDiscGroupExists(dtr("Sales Code").ToString) Then
                    Dim aPr As ArrDiscGroup
                    aPr.ItemCode = dtr("Code").ToString
                    aPr.MaxDate = dtr("MaxDate")
                    aPr.SalesCode = dtr("Sales Code").ToString
                    aPr.SalesType = dtr("Sales Type").ToString
                    aPr.MinQty = dtr("Minimum Quantity")
                    aPr.VariantCode = dtr("Variant Code").ToString
                    aPr.sUOM = dtr("Unit of Measure Code").ToString
                    arrList.Add(aPr)
                End If
            End If
        End While
        'dtr.Close()
        ImportCustDiscItemDet(arrList)
        'ExecuteSQL("Delete from ItemPr where UnitPrice=0")
    End Sub
    Public Sub ImportCustDiscItemDet(ByVal arrList As ArrayList)
        Dim dtr As OdbcDataReader
        dtr = ReadNavRecord("Select * from ""Sales Line Discount""")
        While dtr.Read
            i = i + 1
            For iIndex As Integer = 0 To arrList.Count - 1
                Dim aPr As ArrDiscGroup
                aPr = arrList(iIndex)
                'dtr("Unit of Measure Code") = aPr.sUOM And
                If dtr("Code").ToString.Trim = aPr.ItemCode.Trim And dtr("Sales Code").ToString.Trim = aPr.SalesCode.Trim And dtr("Sales Type").ToString.Trim = aPr.SalesType.Trim And dtr("Minimum Quantity") = aPr.MinQty Then
                    ExecuteSQL("Insert into CustDiscGroupItem(DiscountGroup, ItemNo, DisPer, SalesType,Minqty, VariantCode, UOM) Values (" & SafeSQL(aPr.SalesCode.Trim) & "," & SafeSQL(aPr.ItemCode.Trim) & "," & dtr("Line Discount %").ToString & "," & SafeSQL(aPr.SalesType.Trim) & "," & aPr.MinQty & "," & SafeSQL(aPr.VariantCode) & "," & SafeSQL(aPr.sUOM) & ")")
                    '  ExecuteSQL("Insert into ItemPr(PriceGroup, ItemNo, UnitPrice, SalesType,Minqty, VariantCode, UOM) Values (" & SafeSQL(dtr("Sales Code").ToString) & "," & SafeSQL(dtr("Item No_").ToString) & "," & dtr("Unit Price").ToString & "," & SafeSQL(dtr("Sales Type").ToString) & "," & "0" & "," & SafeSQL(sVariantCode) & "," & SafeSQL(dtr("Unit of Measure Code").ToString) & ")")
                    arrList.RemoveAt(iIndex)
                    Exit For
                End If
            Next
        End While
        i = 17
        'dtr.Close()
        'ExecuteSQL("Delete from Itempr where PriceGroup not in (Select custno from customer where Active=1) and SalesType='Customer'")
    End Sub

    Public Sub ImportItemPrDet(ByVal arrList As ArrayList)
        Dim dtr As OdbcDataReader
        Dim sType As String = ""
        dtr = ReadNavRecord("Select * from ""Sales Price""")
        While dtr.Read
            i = i + 1
            For iIndex As Integer = 0 To arrList.Count - 1
                Dim aPr As ArrItemPrice
                aPr = arrList(iIndex)
                'If aPr.SalesType = 0 Then MsgBox("1")
                'And dtr("Sales Type") = aPr.SalesType
                If dtr("Item No_").ToString.Trim = aPr.ItemCode.Trim And dtr("Sales Code").ToString.Trim = aPr.SalesCode.Trim And dtr("Minimum Quantity") = aPr.MinQty Then
                    ' If aPr.SalesType = 0 Then MsgBox(dtr("Sales Type") = aPr.SalesType)
                    If aPr.SalesType = 2 Then
                        sType = "All Customers"
                    ElseIf aPr.SalesType = 0 Then
                        sType = "Customer"
                    Else
                        sType = "Customer Price Group"
                    End If
                    ExecuteSQL("Insert into ItemPr(PriceGroup, ItemNo, UnitPrice, SalesType, Minqty, VariantCode, UOM) Values (" & SafeSQL(aPr.SalesCode.Trim) & "," & SafeSQL(aPr.ItemCode.Trim) & "," & dtr("Unit Price").ToString & "," & SafeSQL(sType) & "," & aPr.MinQty & "," & SafeSQL(aPr.VariantCode) & "," & SafeSQL(aPr.sUOM) & ")")
                    '  ExecuteSQL("Insert into ItemPr(PriceGroup, ItemNo, UnitPrice, SalesType,Minqty, VariantCode, UOM) Values (" & SafeSQL(dtr("Sales Code").ToString) & "," & SafeSQL(dtr("Item No_").ToString) & "," & dtr("Unit Price").ToString & "," & SafeSQL(dtr("Sales Type").ToString) & "," & "0" & "," & SafeSQL(sVariantCode) & "," & SafeSQL(dtr("Unit of Measure Code").ToString) & ")")
                    arrList.RemoveAt(iIndex)
                    Exit For
                End If
            Next
        End While
        i = 17
        'dtr.Close()
        'ExecuteSQL("Delete from Itempr where PriceGroup not in (Select custno from customer where Active=1) and SalesType='Customer'")
    End Sub
    Public Sub ImportBOMDet(ByVal arrList As ArrayList)
        Dim dtr As OdbcDataReader
        Dim sDesc As String = ""
        Dim sUOM As String = ""
        dtr = ReadNavRecord("Select * from ""Production BOM Line""")
        While dtr.Read
            For iIndex As Integer = 0 To arrList.Count - 1
                ExecuteSQL("Delete from BOmItem where BOMNo=" & SafeSQL(arrList.Item(iIndex)))
                If dtr("Production BOM No_").ToString = arrList.Item(iIndex) Then
                    If IsDBNull(dtr("Description")) = True Then sDesc = "" Else sDesc = dtr("Description")
                    If IsDBNull(dtr("Unit of Measure Code")) = True Then sUOM = "" Else sUOM = dtr("Unit of Measure Code")
                    ExecuteSQL("Insert into BOMItem (BOMNo, [LineNo], ItemNo, Description, UOM, Qty) Values (" & SafeSQL(dtr("Production BOM No_").ToString) & "," & dtr("Line No_") & "," & SafeSQL(dtr("No_").ToString) & "," & SafeSQL(sDesc) & "," & SafeSQL(sUOM) & "," & dtr("Quantity") & ")")
                    Exit For
                End If
            Next
        End While
        dtr.Close()
    End Sub

    Public Sub ImportInvItem(ByVal arrList As ArrayList)

        Dim dtr As OdbcDataReader
        For iIndex As Integer = 0 To arrList.Count - 1
            Dim aPr As ArrInvoice
            aPr = arrList(iIndex)
            dtr = ReadNavRecord("Select * from ""Sales Invoice Line"" where ""Document No_"" = " & SafeSQL(aPr.InvNo) & " order by ""Line No_""")
            While dtr.Read
                ExecuteSQL("Insert into InvItem(InvNo, [LineNo], ItemNo, UOM, Qty, Foc, Price, DisPer, DisPr, Discount, SubAmt, GstAmt, DeliQty, BaseUOM, BaseQty) Values (" & SafeSQL(dtr("Document No_").ToString) & ", " & dtr("Line No_").ToString & ", " & SafeSQL(dtr("No_").ToString) & ", " & SafeSQL(dtr("Unit of Measure Code").ToString) & ", " & dtr("Quantity").ToString & ", 0, " & dtr("Unit Price").ToString & ", " & dtr("Line Discount %").ToString & ", 0, " & dtr("Line Discount Amount") & ", " & dtr("Amount").ToString & ", " & dtr("Amount Including GST") - dtr("Amount") & ", " & dtr("Quantity").ToString & ", " & SafeSQL(dtr("Unit of Measure Code").ToString) & ", " & dtr("Quantity (Base)").ToString & ")")
            End While
        Next
        'Dim rs As SqlDataReader
        'Dim sHistoryunit As String = ""
        'Dim iHistoryNo As Integer = 0

        'rs = ReadRecord("Select HistoryUnit, Historyno from system")
        'If rs.Read = True Then
        '    sHistoryunit = rs("HistoryUnit")
        '    iHistoryNo = rs("HistoryNo")
        'End If
        'rs.Close()
        ''dtr = ReadNavRecord("Select C.* from ""Sales Invoice Line"" C, ""Sales Invoice Header"" A Where C.No_ = A.No_  and  (A.""Document Date"" > " & SafeSQL(Format(DateAdd(DateInterval.Month, (-1) * iHistoryNo, Date.Now), "yyyy-MM-dd")) & " or ""Original Amt_ (LCY)"" >  (""Original Amt_ (LCY)"" - ""Remaining Amount"")) order by C.No_")
        'dtr = ReadNavRecord("Select * from ""Sales Invoice Line"" order by ""Document No_""")
        'While dtr.Read
        '    If dtr("No_").ToString.Trim <> "" Then
        '        For iIndex As Integer = 0 To arrList.Count - 1
        '            Dim aPr As ArrInvoice
        '            aPr = arrList(iIndex)
        '            If dtr("Document No_").ToString = aPr.InvNo Then
        '                ExecuteSQL("Insert into InvItem(InvNo, [LineNo], ItemNo, UOM, Qty, Foc, Price, DisPer, DisPr, Discount, SubAmt, GstAmt, DeliQty, BaseUOM, BaseQty) Values (" & SafeSQL(dtr("Document No_").ToString) & ", " & dtr("Line No_").ToString & ", " & SafeSQL(dtr("No_").ToString) & ", " & SafeSQL(dtr("Unit of Measure Code").ToString) & ", " & dtr("Quantity").ToString & ", 0, " & dtr("Unit Price").ToString & ", " & dtr("Line Discount %").ToString & ", 0, " & dtr("Line Discount Amount") & ", " & dtr("Amount").ToString & ", " & dtr("Amount Including GST") - dtr("Amount") & ", " & dtr("Quantity").ToString & ", " & SafeSQL(dtr("Unit of Measure Code").ToString) & ", " & dtr("Quantity (Base)").ToString & ")")
        '                Exit For
        '            End If
        '        Next
        '    End If
        'End While
        ''dtr.Close()
    End Sub
    Public Sub ImportInvItemEH(ByVal arrList As ArrayList)
        Dim dtr As OdbcDataReader
        dtr = ReadNavRecord("Select * from ""Sales Invoice Line"" order by ""Document No_""")
        While dtr.Read
            For iIndex As Integer = 0 To arrList.Count - 1
                Dim aPr As ArrInvoice
                aPr = arrList(iIndex)
                If dtr("Document No_").ToString = aPr.InvNo Then
                    ExecuteSQL("Insert into InvItem(InvNo, [LineNo], ItemNo, UOM, Qty, Foc, Price, DisPer, DisPr, Discount, SubAmt, GstAmt, DeliQty, BaseUOM, BaseQty) Values (" & SafeSQL(dtr("Document No_").ToString) & ", " & dtr("Line No_").ToString & ", " & SafeSQL(dtr("No_").ToString) & ", " & SafeSQL(dtr("Unit of Measure Code").ToString) & ", " & dtr("Quantity").ToString & ", 0, " & dtr("Unit Price").ToString & ", " & dtr("Line Discount %").ToString & ", 0, " & dtr("Line Discount Amount") & ", " & dtr("Amount").ToString & ", " & dtr("Amount Including GST") - dtr("Amount") & ", " & dtr("Quantity").ToString & ", " & SafeSQL(dtr("Unit of Measure Code").ToString) & ", " & dtr("Quantity (Base)").ToString & ")")
                    Exit For
                End If
            Next
        End While
        dtr.Close()
    End Sub
    Public Sub ImportPayMethod()
        Dim dtr As OdbcDataReader
        Dim rs As SqlDataReader
        dtr = ReadNavRecord("Select Code, Description, ""Bal_ Account Type"",""Bal_ Account No_"" from ""Payment Method""")
        While dtr.Read
            rs = ReadRecord("Select Code from PayMethod where code=" & SafeSQL(dtr("Code")))
            If rs.Read = False Then
                ExecuteSQLAnother("Insert into PayMethod(Code , Description, BalAcType, BalAcNo,Active, PaymentType) Values (" & SafeSQL(dtr("Code").ToString) & "," & SafeSQL(dtr("Description").ToString) & "," & SafeSQL(dtr("Bal_ Account Type").ToString) & "," & SafeSQL(dtr("Bal_ Account No_").ToString) & ",1,0)")
            End If
            rs.Close()
        End While
        dtr.Close()
    End Sub

    Public Sub ImportPayterms()
        Dim dtr As OdbcDataReader
        Dim rs As SqlDataReader
        dtr = ReadNavRecord("Select Code, Description, ""Due Date Calculation"", ""Discount Date Calculation"", ""Discount %"" from ""Payment Terms""")
        While dtr.Read
            rs = ReadRecord("Select Code from Payterms where code=" & SafeSQL(dtr("Code")))
            If rs.Read = False Then
                ExecuteSQLAnother("Insert into PayTerms(Code , Description, DueDateCalc, DisDateCalc, DiscountPercent,Active) Values (" & SafeSQL(dtr("Code").ToString) & "," & SafeSQL(dtr("Description").ToString) & "," & SafeSQL(dtr("Due Date Calculation").ToString) & "," & SafeSQL(dtr("Discount Date Calculation").ToString) & "," & dtr("Discount %").ToString & ",1)")
            End If
            rs.Close()
        End While
        dtr.Close()
    End Sub

    Public Sub ImportPriceGroup()
        Dim dtr As OdbcDataReader
        ExecuteSQL("Delete from PriceGroup")
        dtr = ReadNavRecord("Select Code, Description, ""Allow Invoice Disc_"", ""Allow Line Disc_"" from ""Customer Price Group""")
        While dtr.Read
            ExecuteSQL("Insert into PriceGroup(Code , Description, InvoiceDiscount, LineDiscount) Values (" & SafeSQL(dtr("Code").ToString) & "," & SafeSQL(dtr("Description").ToString) & "," & IIf(dtr("Allow Invoice Disc_"), 1, 0) & "," & IIf(dtr("Allow Line Disc_"), 1, 0) & ")")
        End While
        dtr.Close()
    End Sub
    Public Sub ImportCustDiscGroup()
        Dim dtr As OdbcDataReader
        ExecuteSQL("Delete from CustDiscGroup")
        ExecuteSQL("Delete from Zone")
        dtr = ReadNavRecord("Select Code, Description from ""Customer Discount Group""")
        While dtr.Read
            ExecuteSQL("Insert into CustDiscGroup(Code , Description) Values (" & SafeSQL(dtr("Code").ToString) & "," & SafeSQL(dtr("Description").ToString) & ")")
            ExecuteSQL("Insert into Zone(Code , Description, DTG) Values (" & SafeSQL(dtr("Code").ToString) & "," & SafeSQL(dtr("Description").ToString) & "," & SafeSQL(Date.Now) & ")")
        End While
        dtr.Close()
    End Sub

    Public Sub ImportShipMethod()
        Dim dtr As OdbcDataReader
        ExecuteSQL("Delete from ShipMethod")
        dtr = ReadNavRecord("Select Code, Description from ""Shipment Method""")
        While dtr.Read
            ExecuteSQL("Insert into ShipMethod(Code , Description) Values (" & SafeSQL(dtr("Code").ToString) & "," & SafeSQL(dtr("Description").ToString) & ")")
        End While
        dtr.Close()
    End Sub

    Public Sub ImportUOM()
        Dim dtr As OdbcDataReader
        ExecuteSQL("Delete from UOM")
        dtr = ReadNavRecord("Select Item.""No_ 2"", ""Item No_"", Code, ""Qty_ per Unit of Measure"" from ""Item Unit of Measure"", Item Where ""Item Unit of Measure"".""Item No_"" = Item.No_ ")
        While dtr.Read
            ExecuteSQL("Insert into UOM(ItemNo, Uom , BaseQty) Values (" & SafeSQL(dtr("Item No_").ToString) & "," & SafeSQL(dtr("Code").ToString) & "," & SafeSQL(dtr("Qty_ per Unit of Measure").ToString) & ")")
        End While
        dtr.Close()
    End Sub

    Public Sub ImportZone()
        'Dim dtr As OdbcDataReader
        'ExecuteSQL("Delete from Zone")
        'dtr = ReadNavRecord("Select Distinct ""Service Zone Code"" from Customer")
        'While dtr.Read
        '    If dtr("Service Zone Code").ToString <> "" Then
        '        ExecuteSQL("Insert into Zone(Code , Description) Values (" & SafeSQL(dtr("Service Zone Code").ToString) & "," & SafeSQL(dtr("Service Zone Code").ToString) & ")")
        '    End If
        'End While
        'dtr.Close()
    End Sub

    Public Sub ImportInventory()
        Dim dtr As OdbcDataReader
        ExecuteSQL("Delete from GoodsInvn")
        'and (B.""No_ 2"" <> '' or B.""No_ 2"" is Not Null)
        dtr = ReadNavRecord("Select A.""Location Code"", A.""Item No_"", Sum(A.""Remaining Quantity"") as Qty, B.""Base Unit of Measure"" from ""Item Ledger Entry"" A, Item B Where A.""Item No_"" = B.""No_""  group by A.""Location Code"", A.""Item No_"", B.""Base Unit of Measure""")
        While dtr.Read
            ExecuteSQL("Insert into GoodsInvn(Location, ItemNo, Qty, UOM) Values (" & SafeSQL(dtr("Location Code").ToString) & "," & SafeSQL(dtr("Item No_").ToString) & "," & dtr("Qty") & "," & SafeSQL(dtr("Base Unit of Measure").ToString) & ")")
        End While
        dtr.Close()
    End Sub

    Private Function IsItemExists(ByVal sItemNo As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = ReadRecord("Select ItemNo from Item where ItemNo = " & SafeSQL(sItemNo))
        bAns = dtr.Read
        dtr.Close()
        Return bAns
    End Function

    Private Function IsAgentExists(ByVal sCode As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = ReadRecord("Select Code from SalesAgent where Code = " & SafeSQL(sCode))
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

    Private Function IsNavCustomerExists(ByVal sCustNo As String) As Boolean
        Dim dtr As OdbcDataReader
        Dim bAns As Boolean
        dtr = ReadNavRecord("Select * from Customer where No_ = " & SafeSQL(sCustNo))
        bAns = dtr.Read
        dtr.Close()
        Return bAns
    End Function

    Private Sub ExportCustomer()
        'Dim dtr As SqlDataReader
        'dtr = ReadRecord("Select * from Customer where Exported = 0")
        'While dtr.Read
        '    If IsNavCustomerExists(dtr("CustNo").ToString) Then
        '        If dtr("ExpiryDate").ToString = "" Then
        '            ExecuteNavSQL("Update Customer Set Name = " & SafeSQL(dtr("CustName").ToString) & ", Address = " & SafeSQL(dtr("Address").ToString) & ", ""Address 2"" = " & SafeSQL(dtr("Address2").ToString) & ", ""Address 3"" = " & SafeSQL(dtr("Address3").ToString) & ", ""Address 4"" = " & SafeSQL(dtr("Address4").ToString) & ", ""Post Code"" = " & SafeSQL(dtr("PostCode").ToString) & ", City = " & SafeSQL(dtr("City").ToString) & ", ""Country Code"" = " & SafeSQL(dtr("CountryCode").ToString) & ", ""Phone No_""= " & SafeSQL(dtr("Phone").ToString) & ", Contact = " & SafeSQL(dtr("ContactPerson").ToString) & ", Balance = " & dtr("Balance").ToString & ", ""Fax No_"" = " & SafeSQL(dtr("FaxNo").ToString) & ", ""E-mail"" = " & SafeSQL(dtr("Email").ToString) & ", ""Home Page"" = " & SafeSQL(dtr("Website").ToString) & ", ""Customer Price Group"" = " & SafeSQL(dtr("PriceGroup").ToString) & ", ""Salesperson Code"" = " & SafeSQL(dtr("SalesAgent").ToString) & " Where No_ = " & SafeSQL(dtr("CustNo").ToString))
        '        Else
        '            ExecuteNavSQL("Update Customer Set Name = " & SafeSQL(dtr("CustName").ToString) & ", Address = " & SafeSQL(dtr("Address").ToString) & ", ""Address 2"" = " & SafeSQL(dtr("Address2").ToString) & ", ""Address 3"" = " & SafeSQL(dtr("Address3").ToString) & ", ""Address 4"" = " & SafeSQL(dtr("Address4").ToString) & ", ""Post Code"" = " & SafeSQL(dtr("PostCode").ToString) & ", City = " & SafeSQL(dtr("City").ToString) & ", ""Country Code"" = " & SafeSQL(dtr("CountryCode").ToString) & ", ""Phone No_""= " & SafeSQL(dtr("Phone").ToString) & ", Contact = " & SafeSQL(dtr("ContactPerson").ToString) & ", Balance = " & dtr("Balance").ToString & ", ""Fax No_"" = " & SafeSQL(dtr("FaxNo").ToString) & ", ""E-mail"" = " & SafeSQL(dtr("Email").ToString) & ", ""Home Page"" = " & SafeSQL(dtr("Website").ToString) & ", ""Customer Price Group"" = " & SafeSQL(dtr("PriceGroup").ToString) & ", ""Salesperson Code"" = " & SafeSQL(dtr("SalesAgent").ToString) & ", ""Expiry Date"" = " & SafeSQL(Format(dtr("ExpiryDate"), "yyyy-MM-dd")) & " Where No_ = " & SafeSQL(dtr("CustNo").ToString))
        '        End If
        '    Else
        '        If dtr("ExpiryDate").ToString = "" Then
        '            ExecuteNavSQL("Insert into ""Customer"" (""No_"",""Name"",""Address"",""Address 2"",""Address 3"",""Address 4"",""Post Code"",""City"",""Country Code"",""Phone No_"",""Contact"",""Balance"",""Fax No_"",""E-mail"",""Customer Price Group"",""Salesperson Code"") Values (" & _
        '                            SafeSQL(dtr("CustNo").ToString) & "," & SafeSQL(dtr("CustName").ToString) & ", " & SafeSQL(dtr("Address").ToString) & ", " & SafeSQL(dtr("Address2").ToString) & ", " & SafeSQL(dtr("Address3").ToString) & ", " & SafeSQL(dtr("Address4").ToString) & ", " & SafeSQL(dtr("PostCode").ToString) & ", " & SafeSQL(dtr("City").ToString) & ", " & SafeSQL(dtr("CountryCode").ToString) & ", " & SafeSQL(dtr("Phone").ToString) & ", " & SafeSQL(dtr("ContactPerson").ToString) & ", " & dtr("Balance").ToString & ", " & SafeSQL(dtr("FaxNo").ToString) & ", " & SafeSQL(dtr("Email").ToString) & ", " & SafeSQL(dtr("PriceGroup").ToString) & ", " & SafeSQL(dtr("SalesAgent").ToString) & ")")
        '        Else
        '            ExecuteNavSQL("Insert into ""Customer"" (""No_"",""Name"",""Address"",""Address 2"",""Address 3"",""Address 4"",""Post Code"",""City"",""Country Code"",""Phone No_"",""Contact"",""Balance"",""Fax No_"",""E-mail"",""Customer Price Group"",""Salesperson Code"",""Expiry Date"") Values (" & _
        '                            SafeSQL(dtr("CustNo").ToString) & "," & SafeSQL(dtr("CustName").ToString) & ", " & SafeSQL(dtr("Address").ToString) & ", " & SafeSQL(dtr("Address2").ToString) & ", " & SafeSQL(dtr("Address3").ToString) & ", " & SafeSQL(dtr("Address4").ToString) & ", " & SafeSQL(dtr("PostCode").ToString) & ", " & SafeSQL(dtr("City").ToString) & ", " & SafeSQL(dtr("CountryCode").ToString) & ", " & SafeSQL(dtr("Phone").ToString) & ", " & SafeSQL(dtr("ContactPerson").ToString) & ", " & dtr("Balance").ToString & ", " & SafeSQL(dtr("FaxNo").ToString) & ", " & SafeSQL(dtr("Email").ToString) & ", " & SafeSQL(dtr("PriceGroup").ToString) & ", " & SafeSQL(dtr("SalesAgent").ToString) & ", " & SafeSQL(Format(dtr("ExpiryDate"), "yyyy-MM-dd")) & ")")
        '        End If
        '    End If
        'End While
        'dtr.Close()
        'ExecuteSQL("Update Customer Set Exported = 1 where Exported = 0")
    End Sub


    Private Sub ExportInvoices()
        Dim dtr1 As OdbcDataReader
        Dim sCurCode As String = ""
        dtr1 = ReadNavRecord("Select * from ""General Ledger Setup"" ")
        If dtr1.Read Then
            sCurCode = dtr1("LCY Code")
        End If
        dtr1.Close()
        Dim dExRate As Double = 0
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Customer.AcCustCode, InvNo, [Bill-toNo], InvDt, Invoice.PayTerms, ShipmentMethod, MDT.Location, Dimension1, Dimension2, customer.CustPostGroup, Invoice.CurCode, Invoice.CurExRate, PriceGroup, Invoice.TotalAmt, InvDisGroup, Invoice.AgentID, DONO, PaymentMethod, Invoice.Discount, System.GST, Invoice.Void, Invoice.PONO from Invoice, Customer, System, MDT where System.MDTNo=MDT.MDTNo and Invoice.CustId = Customer.CustNo and Invoice.Exported = 0 and (Invoice.Void=0 or Invoice.void is Null)")
        While dtr.Read
            If dtr("CurCode") <> "" Then
                sCurCode = dtr("CurCode")
                dExRate = dtr("CurExRate")
            Else
                dExRate = 1
            End If
            Dim sExtDocNo As String = ""
            If CBool(dtr("Void")) = True Then
                sExtDocNo = "VOIDED"
            End If
            '1-'Released', 
            'System.IO.File.AppendAllText("C:\Nav.txt", "Insert into ""POS Sales Header"" (""Document Type"",""Sell-to Customer No_"",""No_"",""Bill-to Customer No_"",""Order Date"",""Posting Date"",""Shipment Date"",""Posting Description"",""Payment Terms Code"",""Shipment Method Code"",""Location Code"",""Shortcut Dimension 1 Code"",""Shortcut Dimension 2 Code"",""Customer Posting Group"",""Currency Code"",""Currency Factor"",""Customer Price Group"",""Invoice Disc_ Code"",""Salesperson Code"",""Job No_"",""Shipping No_"",""Posting No_"",""Document Date"",""External Document No_"",""Payment Method Code"",""No_ Series"",""Posting No_ Series"",""Shipping No_ Series"",""Status"",""Invoice Discount Calculation"",""Invoice Discount Value"",""Allow Line Disc_"", ""Due Date"", ""Customer Disc_ Group"") Values (" & _
            '"2," & SafeSQL(dtr("AcCustCode").ToString) & "," & SafeSQL(dtr("InvNo").ToString) & "," & SafeSQL(dtr("AccustCode").ToString) & ",'" & Format(dtr("InvDt"), "yyyy-MM-dd") & "', " & _
            '"'" & Format(dtr("InvDt"), "yyyy-MM-dd") & "','" & Format(dtr("InvDt"), "yyyy-MM-dd") & "','Invoice " & dtr("InvNo").ToString & "'," & _
            'SafeSQL(dtr("PayTerms").ToString) & "," & SafeSQL(dtr("ShipmentMethod").ToString) & "," & SafeSQL(dtr("AgentID").ToString) & "," & SafeSQL("") & "," & SafeSQL(dtr("Dimension2").ToString) & "," & _
            'SafeSQL(dtr("CustPostGroup").ToString) & "," & SafeSQL(sCurCode) & "," & dExRate & "," & SafeSQL(dtr("PriceGroup").ToString) & "," & _
            'SafeSQL(dtr("InvDisGroup").ToString) & "," & SafeSQL(dtr("AgentID").ToString) & ",''," & SafeSQL(dtr("DONO").ToString) & "," & _
            'SafeSQL(dtr("InvNo").ToString) & ",'" & Format(dtr("InvDt"), "yyyy-MM-dd") & "'," & SafeSQL(sExtDocNo) & "," & SafeSQL("") & "," & _
            'SafeSQL(dtr("InvNo").ToString) & "," & SafeSQL(dtr("InvNo").ToString) & "," & SafeSQL(dtr("InvNo").ToString) & "," & _
            '"1,2," & dtr("Discount").ToString & ",1" & ",'" & Format(dtr("InvDt"), "yyyy-MM-dd") & "','')")


            ExecuteNavSQL("Insert into ""POS Sales Header"" (""Document Type"",""Sell-to Customer No_"",""No_"",""Bill-to Customer No_"",""Order Date"",""Posting Date"",""Shipment Date"",""Posting Description"",""Payment Terms Code"",""Shipment Method Code"",""Location Code"",""Shortcut Dimension 1 Code"",""Shortcut Dimension 2 Code"",""Customer Posting Group"",""Currency Code"",""Currency Factor"",""Customer Price Group"",""Invoice Disc_ Code"",""Salesperson Code"",""Job No_"",""Shipping No_"",""Posting No_"",""Document Date"",""External Document No_"",""Payment Method Code"",""No_ Series"",""Posting No_ Series"",""Shipping No_ Series"",""Status"",""Invoice Discount Calculation"",""Invoice Discount Value"",""Allow Line Disc_"", ""Due Date"", ""Customer Disc_ Group"") Values (" & _
            "1," & SafeSQL(dtr("AcCustCode").ToString) & "," & SafeSQL(dtr("InvNo").ToString) & "," & SafeSQL(dtr("AcCustCode").ToString) & ",'" & Format(dtr("InvDt"), "yyyy-MM-dd") & "', " & _
            "'" & Format(dtr("InvDt"), "yyyy-MM-dd") & "','" & Format(dtr("InvDt"), "yyyy-MM-dd") & "','Invoice " & dtr("InvNo").ToString & "'," & _
            SafeSQL(dtr("PayTerms").ToString) & "," & SafeSQL(dtr("ShipmentMethod").ToString) & "," & SafeSQL(dtr("AgentID").ToString) & "," & SafeSQL("") & "," & SafeSQL(dtr("Dimension2").ToString) & "," & _
            SafeSQL(dtr("CustPostGroup").ToString) & "," & SafeSQL(sCurCode) & "," & dExRate & "," & SafeSQL(dtr("PriceGroup").ToString) & "," & _
            SafeSQL(dtr("InvDisGroup").ToString) & "," & SafeSQL(dtr("AgentID").ToString) & ",''," & SafeSQL(dtr("InvNo").ToString) & "," & _
            SafeSQL(dtr("InvNo").ToString) & ",'" & Format(dtr("InvDt"), "yyyy-MM-dd") & "'," & SafeSQL(sExtDocNo) & "," & SafeSQL("") & "," & _
            SafeSQL(dtr("InvNo").ToString) & "," & SafeSQL(dtr("InvNo").ToString) & "," & SafeSQL(dtr("InvNo").ToString) & "," & _
            "1,2," & dtr("Discount").ToString & ",1" & ",'" & Format(dtr("InvDt"), "yyyy-MM-dd") & "','')")

            'Payment method replaced with blank dtr("PaymentMethod").ToString
        End While
        dtr.Close()
    End Sub

    Private Sub ExportInvItem()
        Dim dtr1 As SqlDataReader
        Dim iIndex As Integer = 0
        Dim sInvNo As String = ""
        Dim arr As New ArrayList
        Try
            dtr1 = ReadRecord("Select PriceGroup, Customer.AcCustCode, Invoice.AgentID, InvItem.Location, Invoice.Invdt, Invoice.Gst, Invoice.InvNo, [LineNo], InvItem.ItemNo, InvItem.Description, Item.BaseUOM, InvItem.Qty, InvItem.Price, InvItem.Discount, DisPer, DisPr, AllowInvDiscount, UOM.BaseQty, InvItem.Uom from InvItem, Item, UOM, Invoice, Customer where Invoice.CustId = Customer.CustNo and Invoice.InvNo = InvItem.InvNo and UOM.ItemNo = InvItem.ItemNo and Uom.Uom = InvItem.Uom and Item.ItemNo = InvItem.ItemNo and Invoice.Exported = 0 and (Invoice.Void=0 or Invoice.void is Null) order by Invoice.InvNo, InvItem.ItemNo")
            While dtr1.Read
                If sInvNo <> dtr1("InvNo") Then
                    sInvNo = dtr1("InvNo")
                    iIndex = 0
                End If

                iIndex = iIndex + 1
                '1-'Item'
                Dim dDisc, dDisPr As Double
                If IsDBNull(dtr1("Discount")) = True Then
                    dDisc = 0
                Else
                    dDisc = dtr1("Discount")
                End If
                If IsDBNull(dtr1("DisPer")) = True Then
                    dDisPr = 0
                Else
                    dDisPr = dtr1("DisPer")
                End If

                Dim sSQL As String = "Insert into ""POS Sales Line"" (""Document Type"",""Sell-to Customer No_"",""Document No_"",""Line No_"",""Type"",""No_"",""Location Code"",""Shipment Date"",""Description"",""Unit of Measure"",""Quantity"",""Unit Price"",""Allow Invoice Disc_"",""Customer Price Group"",""Job No_"",""Qty_ per Unit of Measure"",""Unit of Measure Code"",""Line Discount Amount"",""Line Discount %"",""Shortcut Dimension 1 Code"",""Shortcut Dimension 2 Code"") Values (" & _
                "1," & SafeSQL(dtr1("AcCustCode").ToString) & "," & SafeSQL(dtr1("InvNo").ToString) & "," & CStr(iIndex * 10000) & ",2," & _
                SafeSQL(dtr1("ItemNo").ToString) & "," & SafeSQL(dtr1("AgentID").ToString) & ",'" & Format(dtr1("InvDt"), "yyyy-MM-dd") & "'," & SafeSQL(dtr1("Description").ToString) & "," & SafeSQL(dtr1("BaseUOM").ToString) & "," & dtr1("Qty") & "," & dtr1("Price") & "," & _
                 IIf(dtr1("AllowInvDiscount"), 1, 0) & "," & SafeSQL(dtr1("PriceGroup").ToString) & ",''," & dtr1("BaseQty") & "," & SafeSQL(dtr1("Uom").ToString) & "," & dDisc & "," & dDisPr & "," & SafeSQL("") & ",'')"
                arr.Add(sSQL)

                'If Mid(dtr1("Description").ToString, 1, 8).ToUpper = "EXCHANGE" Then
                '    iIndex = iIndex + 1
                '    sSQL = "Insert into ""POS Sales Line"" (""Document Type"",""Sell-to Customer No_"",""Document No_"",""Line No_"",""Type"",""No_"",""Location Code"",""Shipment Date"",""Description"",""Unit of Measure"",""Quantity"",""Unit Price"",""Allow Invoice Disc_"",""Customer Price Group"",""Job No_"",""Qty_ per Unit of Measure"",""Unit of Measure Code"",""Line Discount Amount"",""Line Discount %"",""Shortcut Dimension 1 Code"",""Shortcut Dimension 2 Code"") Values (" & _
                '"1," & SafeSQL(dtr1("AcCustCode").ToString) & "," & SafeSQL(dtr1("InvNo").ToString) & "," & CStr(iIndex * 10000) & ",2," & _
                'SafeSQL(dtr1("ItemNo").ToString) & "," & SafeSQL(dtr1("Location").ToString) & ",'" & Format(dtr1("InvDt"), "yyyy-MM-dd") & "'," & SafeSQL(dtr1("Description").ToString) & "," & SafeSQL(dtr1("BaseUOM").ToString) & "," & dtr1("Qty") * -1 & "," & dtr1("Price") & "," & _
                ' IIf(dtr1("AllowInvDiscount"), 1, 0) & "," & SafeSQL(dtr1("PriceGroup").ToString) & ",''," & dtr1("BaseQty") & "," & SafeSQL(dtr1("Uom").ToString) & "," & dDisc & "," & dDisPr & "," & SafeSQL("") & ",'')"
                '    arr.Add(sSQL)
                'System.IO.File.AppendAllText("C:\Nav.txt", sSQL)
                'End If
            End While
            dtr1.Close()
            For iIndex = 0 To arr.Count - 1
                If iIndex Mod 10 = 0 Then
                    Dim dt As Date = DateAdd(DateInterval.Second, 3, Date.Now)
                    Do Until Date.Now > dt
                        Debug.Print("Test")
                    Loop
                End If
                ExecuteNavSQL(arr(iIndex).ToString)
            Next
            ExecuteSQL("Update Invoice Set Exported = 1 where Exported = 0")
        Catch ex As Exception
            dtr1.Close()
            MsgBox(sInvNo)
            MsgBox(ex.Message)
            Exit Sub
        End Try
        'dtr1 = ReadRecord("Select PriceGroup, Customer.AcCustCode, Invoice.AgentID, InvItem.Location, Invoice.Invdt, Invoice.Gst, Invoice.InvNo, [LineNo], InvItem.ItemNo, InvItem.Description, Item.BaseUOM, InvItem.Qty, InvItem.Price, InvItem.Discount, DisPer, DisPr, AllowInvDiscount, UOM.BaseQty, InvItem.Uom from InvItem, Item, UOM, Invoice, Customer where Invoice.CustId = Customer.CustNo and Invoice.InvNo = InvItem.InvNo and UOM.ItemNo = InvItem.ItemNo and Uom.Uom = InvItem.Uom and Item.ItemNo = InvItem.ItemNo and Invoice.Exported = 0 and (Invoice.Void=0 or Invoice.void is Null) order by Invoice.InvNo, InvItem.ItemNo")
        'While dtr1.Read
        '    If sInvNo <> dtr1("InvNo") Then
        '        sInvNo = dtr1("InvNo")
        '        iIndex = 0
        '    End If

        '    iIndex = iIndex + 1
        '    '1-'Item'
        '    Dim dDisc, dDisPr As Double
        '    If IsDBNull(dtr1("Discount")) = True Then
        '        dDisc = 0
        '    Else
        '        dDisc = dtr1("Discount")
        '    End If
        '    If IsDBNull(dtr1("DisPer")) = True Then
        '        dDisPr = 0
        '    Else
        '        dDisPr = dtr1("DisPer")
        '    End If

        '    Dim sSQL As String = "Insert into ""POS Sales Line"" (""Document Type"",""Sell-to Customer No_"",""Document No_"",""Line No_"",""Type"",""No_"",""Location Code"",""Shipment Date"",""Description"",""Unit of Measure"",""Quantity"",""Unit Price"",""Allow Invoice Disc_"",""Customer Price Group"",""Job No_"",""Qty_ per Unit of Measure"",""Unit of Measure Code"",""Line Discount Amount"",""Line Discount %"",""Shortcut Dimension 1 Code"",""Shortcut Dimension 2 Code"") Values (" & _
        '    "1," & SafeSQL(dtr1("AcCustCode").ToString) & "," & SafeSQL(dtr1("InvNo").ToString) & "," & CStr(iIndex * 10000) & ",2," & _
        '    SafeSQL(dtr1("ItemNo").ToString) & "," & SafeSQL(dtr1("AgentID").ToString) & ",'" & Format(dtr1("InvDt"), "yyyy-MM-dd") & "'," & SafeSQL(dtr1("Description").ToString) & "," & SafeSQL(dtr1("BaseUOM").ToString) & "," & dtr1("Qty") & "," & dtr1("Price") & "," & _
        '     IIf(dtr1("AllowInvDiscount"), 1, 0) & "," & SafeSQL(dtr1("PriceGroup").ToString) & ",''," & dtr1("BaseQty") & "," & SafeSQL(dtr1("Uom").ToString) & "," & dDisc & "," & dDisPr & "," & SafeSQL("") & ",'')"
        '    arr.Add(sSQL)

        '    If Mid(dtr1("Description").ToString, 1, 8).ToUpper = "EXCHANGE" Then
        '        iIndex = iIndex + 1
        '        sSQL = "Insert into ""POS Sales Line"" (""Document Type"",""Sell-to Customer No_"",""Document No_"",""Line No_"",""Type"",""No_"",""Location Code"",""Shipment Date"",""Description"",""Unit of Measure"",""Quantity"",""Unit Price"",""Allow Invoice Disc_"",""Customer Price Group"",""Job No_"",""Qty_ per Unit of Measure"",""Unit of Measure Code"",""Line Discount Amount"",""Line Discount %"",""Shortcut Dimension 1 Code"",""Shortcut Dimension 2 Code"") Values (" & _
        '    "1," & SafeSQL(dtr1("AcCustCode").ToString) & "," & SafeSQL(dtr1("InvNo").ToString) & "," & CStr(iIndex * 10000) & ",2," & _
        '    SafeSQL(dtr1("ItemNo").ToString) & "," & SafeSQL(dtr1("Location").ToString) & ",'" & Format(dtr1("InvDt"), "yyyy-MM-dd") & "'," & SafeSQL(dtr1("Description").ToString) & "," & SafeSQL(dtr1("BaseUOM").ToString) & "," & dtr1("Qty") * -1 & "," & dtr1("Price") & "," & _
        '     IIf(dtr1("AllowInvDiscount"), 1, 0) & "," & SafeSQL(dtr1("PriceGroup").ToString) & ",''," & dtr1("BaseQty") & "," & SafeSQL(dtr1("Uom").ToString) & "," & dDisc & "," & dDisPr & "," & SafeSQL("") & ",'')"
        '        arr.Add(sSQL)
        '        System.IO.File.AppendAllText("C:\Nav.txt", sSQL)
        '    End If
        'End While
        'dtr1.Close()
        'For iIndex = 0 To arr.Count - 1
        '    If iIndex Mod 10 = 0 Then
        '        Dim dt As Date = DateAdd(DateInterval.Second, 3, Date.Now)
        '        Do Until Date.Now > dt
        '            Debug.Print("Test")
        '        Loop
        '    End If
        '    ExecuteNavSQL(arr(iIndex).ToString)
        'Next
        'ExecuteSQL("Update Invoice Set Exported = 1 where Exported = 0")
    End Sub
    Private Sub ExportOrders()
        Dim dtr1 As OdbcDataReader
        Dim sCurCode As String = ""
        dtr1 = ReadNavRecord("Select * from ""General Ledger Setup"" ")
        If dtr1.Read Then
            sCurCode = dtr1("LCY Code")
        End If
        dtr1.Close()
        Dim dExRate As Double = 0
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select OrderHdr.CustId, OrdNo, PoNo, OrdDt, Customer.PaymentTerms as PayTerms, ShipmentMethod, MDT.Location, Dimension1, Dimension2, Customer.CustPostGroup, OrderHdr.CurCode, OrderHdr.CurExRate, PriceGroup, OrderHdr.TotalAmt, InvDisGroup, OrderHdr.AgentID, PONO, PaymentMethod, OrderHdr.Discount, System.GST, OrderHdr.Void, OrderHdr.PONO, OrderHdr.ShipName, OrderHdr.ShipAdd, OrderHdr.ShipAdd2, OrderHdr.ShipAdd3, OrderHdr.ShipAdd4, OrderHdr.ShipCity, OrderHdr.ShipPin, OrderHdr.Remarks  from OrderHdr, Customer, System,MDT where OrderHdr.AgentID=MDT.AgentID  and OrderHdr.CustId = Customer.CustNo and OrderHdr.Exported = 0")
        While dtr.Read
            Dim sExtDocNo As String = ""
            If dtr("Void") = True Then
                sExtDocNo = "VOIDED"
            Else
                sExtDocNo = dtr("PONO").ToString
            End If
            If dtr("CurCode") <> "" Then
                sCurCode = dtr("CurCode")
                dExRate = dtr("CurExRate")
            Else
                dExRate = 1
            End If
            '2-'Order', 1-'Released', 1-'%'

            'Document Type
            'Allow Line Disc
            'Status
            'Invoice Disc Calculation
            ExecuteNavSQL("Insert into ""POS Sales Header"" (""Document Type"",""Sell-to Customer No_"",""No_"",""Bill-to Customer No_"",""Ship-to Code"",""Order Date"",""Posting Date"",""Shipment Date"",""Posting Description"",""Payment Terms Code"",""Shipment Method Code"",""Location Code"",""Shortcut Dimension 1 Code"",""Shortcut Dimension 2 Code"",""Customer Posting Group"",""Currency Code"",""Currency Factor"",""Customer Price Group"",""Invoice Disc_ Code"",""Salesperson Code"",""Job No_"",""Shipping No_"",""Posting No_"",""Document Date"",""External Document No_"",""Payment Method Code"",""No_ Series"",""Posting No_ Series"",""Shipping No_ Series"",""Status"",""Invoice Discount Calculation"",""Invoice Discount Value"",""Allow Line Disc_"",""Ship-to Name"",""Ship-to Address"",""Ship-to Address 2"",""Ship-to Address 3"",""Ship-to Address 4"",""Ship-to Post Code"",""Ship-to City"") Values (" & _
            "1," & SafeSQL(dtr("CustId").ToString) & "," & SafeSQL(dtr("OrdNo").ToString) & "," & SafeSQL(dtr("CustID").ToString) & "," & SafeSQL(dtr("CustId").ToString) & ",'" & Format(dtr("OrdDt"), "yyyy-MM-dd") & "', " & _
            "'" & Format(dtr("OrdDt"), "yyyy-MM-dd") & "','" & Format(dtr("OrdDt"), "yyyy-MM-dd") & "','Order " & dtr("OrdNo").ToString & "'," & _
            SafeSQL(dtr("PayTerms").ToString) & "," & SafeSQL(dtr("ShipmentMethod").ToString) & "," & SafeSQL(dtr("Location").ToString) & "," & SafeSQL("") & "," & SafeSQL(dtr("Dimension2").ToString) & "," & _
            SafeSQL(dtr("CustPostGroup").ToString) & "," & SafeSQL(sCurCode) & "," & dExRate & "," & SafeSQL(dtr("PriceGroup").ToString) & "," & _
            SafeSQL(dtr("InvDisGroup").ToString) & "," & SafeSQL(dtr("AgentID").ToString) & ",''," & SafeSQL("") & "," & _
            SafeSQL("") & ",'" & Format(dtr("OrdDt"), "yyyy-MM-dd") & "'," & SafeSQL(sExtDocNo) & "," & SafeSQL("") & "," & _
            SafeSQL(dtr("OrdNo").ToString) & "," & SafeSQL(dtr("OrdNo").ToString) & "," & SafeSQL(dtr("OrdNo").ToString) & "," & _
            "1,2," & dtr("Discount").ToString & ",1," & SafeSQL(dtr("ShipName").ToString) & "," & SafeSQL(dtr("ShipAdd").ToString) & "," & SafeSQL(dtr("ShipAdd2").ToString) & "," & SafeSQL(dtr("ShipAdd3").ToString) & "," & SafeSQL(dtr("ShipAdd4").ToString) & "," & SafeSQL(dtr("ShipPin").ToString) & "," & SafeSQL(dtr("ShipCity").ToString) & ")")
        End While
        dtr.Close()
    End Sub

    Private Sub ExportOrdItem()
        Dim dtr1 As SqlDataReader
        Dim iIndex As Integer = 0
        Dim sOrdNo As String = ""
        Dim arr As New ArrayList
        dtr1 = ReadRecord("Select PriceGroup, OrderHdr.CustId, MDT.Location, OrderHdr.OrdDt, OrderHdr.Gst, OrderHdr.OrdNo, [LineNo], OrdItem.ItemNo, Item.Description, Item.BaseUOM, OrdItem.Qty, OrdItem.Price, DisPer, DisPr, AllowInvDiscount, UOM.BaseQty, OrdItem.Uom from OrdItem, Item, UOM, OrderHdr, Customer,MDT where OrderHdr.AgentID=MDT.AgentID and OrderHdr.CustId = Customer.CustNo and OrderHdr.OrdNo = OrdItem.OrdNo and UOM.ItemNo = OrdItem.ItemNo and Uom.Uom = OrdItem.Uom and Item.ItemNo = OrdItem.ItemNo and OrderHdr.Exported = 0 order by OrdNo, OrdItem.ItemNo")
        While dtr1.Read
            If sOrdNo <> dtr1("OrdNo") Then
                sOrdNo = dtr1("OrdNo")
                iIndex = 0
            End If
            iIndex = iIndex + 1
            '2-'Order', 1-'Item'
            Dim sSQL As String = "Insert into ""POS Sales Line"" (""Document Type"",""Sell-to Customer No_"",""Document No_"",""Line No_"",""Type"",""No_"",""Location Code"",""Shipment Date"",""Description"",""Unit of Measure"",""Quantity"",""Unit Price"",""Allow Invoice Disc_"",""Customer Price Group"",""Job No_"",""Qty_ per Unit of Measure"",""Unit of Measure Code"",""Shortcut Dimension 1 Code"",""Shortcut Dimension 2 Code"") Values (" & _
            "1," & SafeSQL(dtr1("CustId").ToString) & "," & SafeSQL(dtr1("OrdNo").ToString) & "," & CStr(iIndex * 10000) & ",2," & _
            SafeSQL(dtr1("ItemNo").ToString) & "," & SafeSQL(dtr1("Location").ToString) & ",'" & Format(dtr1("OrdDt"), "yyyy-MM-dd") & "'," & SafeSQL(dtr1("Description").ToString) & "," & SafeSQL(dtr1("BaseUOM").ToString) & "," & dtr1("Qty") & "," & dtr1("Price") & "," & _
             IIf(dtr1("AllowInvDiscount"), 1, 0) & "," & SafeSQL(dtr1("PriceGroup").ToString) & ",''," & dtr1("BaseQty") & "," & SafeSQL(dtr1("Uom").ToString) & "," & SafeSQL("") & ",'')"
            arr.Add(sSQL)
        End While
        dtr1.Close()
        For iIndex = 0 To arr.Count - 1
            If iIndex Mod 10 = 0 Then
                Dim dt As Date = DateAdd(DateInterval.Second, 3, Date.Now)
                Do Until Date.Now > dt
                    Debug.Print("Test")
                Loop
            End If
            ExecuteNavSQL(arr(iIndex).ToString)
        Next
        ExecuteSQL("Update OrderHdr Set Exported = 1 where Exported = 0")
    End Sub
    Private Function GetItemTransLastNo() As Integer
        Dim dtr As OdbcDataReader
        Dim i As Integer = 1
        dtr = ReadNavRecord("Select max(""Line No_"") as MaxLine from ""POS Item Journal Line""")
        If dtr.Read = True Then
            If Not IsDBNull(dtr("MaxLine")) Then
                i = dtr("MaxLine") / 10000 + 1
            End If
        End If
        dtr.Close()
        Return i
    End Function
    Private Function GetReceiptLastNo() As Integer
        Dim dtr As OdbcDataReader
        Dim i As Integer = 1
        dtr = ReadNavRecord("Select max(""Line No_"") as MaxLine from ""POS Gen_ Journal Line""")
        If dtr.Read = True Then
            If Not IsDBNull(dtr("MaxLine")) Then
                i = dtr("MaxLine") / 10000 + 1
            End If
        End If
        dtr.Close()
        Return i
    End Function
    Private Sub ExportReceipt(ByVal iIndex As Integer)
        Dim dtr2 As OdbcDataReader
        Dim sCurCode As String = ""
        dtr2 = ReadNavRecord("Select * from ""General Ledger Setup"" ")
        If dtr2.Read Then
            sCurCode = dtr2("LCY Code")
        End If
        dtr2.Close()
        Dim dtr As SqlDataReader
        Dim dExRate As Double = 0
        dtr = ReadRecord("Select CustNo, [Bill-toNo], BalAcType, BalAcNo, RcptDt, RcptNo, ChqNo, ChqDt, CurCode, Amount, CurExRate, Receipt.AgentID, BankAccount, Receipt.PayMethod, PayMethod.Description, MDT.Location from Receipt, PayMethod, Customer, System, MDT Where Receipt.PayMethod = PayMethod.Code and Receipt.CustId = Customer.CustNo and System.MDTNo=MDT.MDTNo and Receipt.Exported = 0 and Receipt.Void = 0")
        Dim SInvNo As String = ""
        While dtr.Read
            Dim dtr1 As SqlDataReader
            dtr1 = ReadRecordAnother("Select InvNo, AmtPaid from RcptItem Where RcptNo = " & SafeSQL(dtr("RcptNo")))
            While dtr1.Read
                If dtr("CurCode") <> "" Then
                    sCurCode = dtr("CurCode")
                    dExRate = dtr("CurExRate")
                Else
                    dExRate = 1
                End If
                ExecuteNavSQL("Insert into ""POS Gen_ Journal Line"" (""Journal Template Name"",""Line No_"",""Account Type"",""Account No_"",""Posting Date"",""Document Type"",""Document No_"",""Description"",""Currency Code"",""Amount"",""Debit Amount"",""Credit Amount"",""Currency Factor"",""Salespers_/Purch_ Code"",""Source Code"",""Applies-to Doc_ Type"",""Applies-to Doc_ No_"",""Journal Batch Name"") Values (" & SafeSQL(sGenJournalTemplate) & "," & CStr(iIndex * 10000) & ",1," & SafeSQL(IIf(dtr("Bill-toNo").ToString = "", dtr("CustNo").ToString, dtr("Bill-toNo").ToString)) & ",'" & Format(dtr("RcptDt"), "yyyy-MM-dd") & "',1," & SafeSQL(dtr("RcptNo").ToString) & "," & SafeSQL(dtr1("InvNo").ToString) & "," & SafeSQL(sCurCode) & "," & CStr(-1 * CDbl(dtr1("AmtPaid").ToString)) & ",0," & dtr1("AmtPaid").ToString & "," & dExRate & "," & SafeSQL(dtr("AgentID").ToString) & ",'CASHRECJNL',2," & SafeSQL(dtr1("InvNo").ToString) & "," & SafeSQL(sGenJournalBatch) & ")")
                SInvNo = dtr1("InvNo").ToString
                iIndex = iIndex + 1
            End While
            dtr1.Close()
            Dim iActype As Integer = 0
            ExecuteNavSQL("Insert into ""POS Gen_ Journal Line"" (""Journal Template Name"",""Line No_"",""Account Type"",""Account No_"",""Posting Date"",""Document Type"",""Document No_"",""Description"",""Currency Code"",""Amount"",""Debit Amount"",""Credit Amount"",""Currency Factor"",""Salespers_/Purch_ Code"",""Source Code"",""Applies-to Doc_ Type"",""Journal Batch Name"") Values (" & SafeSQL(sGenJournalTemplate) & "," & CStr(iIndex * 10000) & "," & iActype & "," & SafeSQL(dtr("BalAcNo").ToString) & "," & SafeSQL(Format(dtr("RcptDt"), "yyyy-MM-dd")) & ",1," & SafeSQL(dtr("RcptNo").ToString) & "," & SafeSQL(SInvNo) & "," & SafeSQL(sCurCode) & "," & dtr("Amount").ToString & "," & dtr("Amount").ToString & ",0," & dExRate & "," & SafeSQL(dtr("AgentID").ToString) & ",'CASHRECJNL',2," & SafeSQL(sGenJournalBatch) & ")")
            iIndex = iIndex + 1
        End While
        dtr.Close()
        ExecuteSQL("Update Receipt Set Exported = 1 where Exported = 0")
    End Sub

    Private Sub ExportItemTrans(ByVal iIndex As Integer)

        Dim dtr As SqlDataReader
        Try
            dtr = ReadRecord("Select DocNO, DocDt, ItemId, Qty, Uom, Location from ItemTrans Where DocType = 'GIN' and Exported = 0")
            While dtr.Read
                ExecuteNavSQL("Insert into ""POS Item Journal Line"" (""Journal Template Name"",""Line No_"",""Item No_"",""Posting Date"",""Document No_"",""Quantity"",""Unit of Measure Code"",""Location Code"",""Journal Batch Name"",""Entry Type"") Values (" & SafeSQL(sItemJnlTemplate) & "," & CStr(iIndex * 10000) & "," & SafeSQL(dtr("ItemId").ToString) & ",'" & Format(dtr("DocDt"), "yyyy-MM-dd") & "'," & SafeSQL(dtr("DocNo").ToString) & "," & dtr("Qty").ToString & "," & SafeSQL(dtr("UOM").ToString) & "," & SafeSQL(dtr("Location").ToString) & "," & SafeSQL(sItemJnlBatch) & ",2)")
                iIndex = iIndex + 1
            End While
            dtr.Close()
        Catch eodbc As OdbcException
            MsgBox(eodbc.Message, MsgBoxStyle.Critical, "Goods IN")
            ExecuteNavSQL("Delete from ""POS Item Journal Line""")
            dtr.Close()
            Return
        Catch ex As Exception
            MsgBox(ex.Message)
            ExecuteNavSQL("Delete from ""POS Item Journal Line""")
            dtr.Close()
            Return
        End Try
        Try
            dtr = ReadRecord("Select DocNO, DocDt, ItemId, Qty, Uom, Location from ItemTrans Where DocType = 'GOUT' and Exported = 0")
            While dtr.Read
                ExecuteNavSQL("Insert into ""POS Item Journal Line"" (""Journal Template Name"",""Line No_"",""Item No_"",""Posting Date"",""Document No_"",""Quantity"",""Unit of Measure Code"",""Location Code"",""Journal Batch Name"",""Entry Type"") Values (" & SafeSQL(sItemJnlTemplate) & "," & CStr(iIndex * 10000) & "," & SafeSQL(dtr("ItemId").ToString) & ",'" & Format(dtr("DocDt"), "yyyy-MM-dd") & "'," & SafeSQL(dtr("DocNo").ToString) & "," & dtr("Qty") & "," & SafeSQL(dtr("UOM").ToString) & "," & SafeSQL(dtr("Location").ToString) & "," & SafeSQL(sItemJnlBatch) & ",3)")
                iIndex = iIndex + 1
            End While
            dtr.Close()
        Catch eodbc As OdbcException
            MsgBox(eodbc.Message, MsgBoxStyle.Critical, "Goods OUT")
            ExecuteNavSQL("Delete from ""POS Item Journal Line""")
            dtr.Close()
            Return
        Catch ex As Exception
            MsgBox(ex.Message)
            ExecuteNavSQL("Delete from ""POS Item Journal Line""")
            dtr.Close()
            Return
        End Try
        Try
            dtr = ReadRecord("Select DocNO, DocDt, ItemId, Qty, Uom, Location from ItemTrans Where DocType = 'EX' and Exported = 0")
            While dtr.Read
                ExecuteNavSQL("Insert into ""POS Item Journal Line"" (""Journal Template Name"",""Line No_"",""Item No_"",""Posting Date"",""Document No_"",""Quantity"",""Unit of Measure Code"",""Location Code"",""Journal Batch Name"",""Entry Type"") Values (" & SafeSQL(sItemJnlTemplate) & "," & CStr(iIndex * 10000) & "," & SafeSQL(dtr("ItemId").ToString) & ",'" & Format(dtr("DocDt"), "yyyy-MM-dd") & "'," & SafeSQL(dtr("DocNo").ToString) & "," & dtr("Qty") & "," & SafeSQL(dtr("UOM").ToString) & "," & SafeSQL(dtr("Location").ToString) & "," & SafeSQL(sItemJnlBatch) & ",3)")
                iIndex = iIndex + 1
            End While
            dtr.Close()
        Catch eodbc As OdbcException
            MsgBox(eodbc.Message, MsgBoxStyle.Critical, "Exchange")
            ExecuteNavSQL("Delete from ""POS Item Journal Line""")
            dtr.Close()
            Return
        Catch ex As Exception
            MsgBox(ex.Message)
            ExecuteNavSQL("Delete from ""POS Item Journal Line""")
            dtr.Close()
            Return
        End Try
        Try
            dtr = ReadRecord("Select DocNO, DocDt, ItemId, Qty, Uom, Location from ItemTrans Where DocType = 'GVAR' and Exported = 0")
            While dtr.Read
                If dtr("Qty") > 0 Then
                    ExecuteNavSQL("Insert into ""POS Item Journal Line"" (""Journal Template Name"",""Line No_"",""Item No_"",""Posting Date"",""Document No_"",""Quantity"",""Unit of Measure Code"",""Location Code"",""Journal Batch Name"",""Entry Type"") Values (" & SafeSQL(sItemJnlTemplate) & "," & CStr(iIndex * 10000) & "," & SafeSQL(dtr("ItemId").ToString) & ",'" & Format(dtr("DocDt"), "yyyy-MM-dd") & "'," & SafeSQL(dtr("DocNo").ToString) & "," & dtr("Qty") & "," & SafeSQL(dtr("UOM").ToString) & "," & SafeSQL(dtr("Location").ToString) & "," & SafeSQL(sItemJnlBatch) & ",2)")
                Else
                    ExecuteNavSQL("Insert into ""POS Item Journal Line"" (""Journal Template Name"",""Line No_"",""Item No_"",""Posting Date"",""Document No_"",""Quantity"",""Unit of Measure Code"",""Location Code"",""Journal Batch Name"",""Entry Type"") Values (" & SafeSQL(sItemJnlTemplate) & "," & CStr(iIndex * 10000) & "," & SafeSQL(dtr("ItemId").ToString) & ",'" & Format(dtr("DocDt"), "yyyy-MM-dd") & "'," & SafeSQL(dtr("DocNo").ToString) & "," & Math.Abs(dtr("Qty")) & "," & SafeSQL(dtr("UOM").ToString) & "," & SafeSQL(dtr("Location").ToString) & "," & SafeSQL(sItemJnlBatch) & ",3)")
                End If
                iIndex = iIndex + 1
            End While
            dtr.Close()
        Catch eodbc As OdbcException
            MsgBox(eodbc.Message, MsgBoxStyle.Critical, "Goods Variance")
            ExecuteNavSQL("Delete from ""POS Item Journal Line""")
            dtr.Close()
            Return
        Catch ex As Exception
            MsgBox(ex.Message)
            ExecuteNavSQL("Delete from ""POS Item Journal Line""")
            dtr.Close()
            Return
        End Try
        ExecuteSQL("Update ItemTrans Set Exported = 1 where Exported = 0")
    End Sub

    Private Sub ExportCreditNote()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Customer.AcCustCode, CreditNote.CustNo, CreditNoteNo, [Bill-toNo], CreditDate, PaymentTerms, ShipmentMethod, Location, Dimension1, Dimension2, Customer.CustPostGroup, CurrencyCode, PriceGroup, TotalAmt, InvDisGroup, SalesPersonCode, PaymentMethod, Discount, System.GST, Void from CreditNote, Customer, System where CreditNote.CustNo = Customer.CustNo and CreditNote.Exported = 0")
        While dtr.Read
            Dim sExtDocNo As String = ""
            If dtr("Void") = True Then
                sExtDocNo = "VOIDED"
            End If
            ExecuteNavSQL("Insert into ""POS Sales Header"" (""Document Type"",""Sell-to Customer No_"",""No_"",""Bill-to Customer No_"",""Ship-to Code"",""Order Date"",""Posting Date"",""Shipment Date"",""Posting Description"",""Payment Terms Code"",""Shipment Method Code"",""Location Code"",""Shortcut Dimension 1 Code"",""Shortcut Dimension 2 Code"",""Customer Posting Group"",""Currency Code"",""Currency Factor"",""Customer Price Group"",""Invoice Disc_ Code"",""Salesperson Code"",""Job No_"",""Shipping No_"",""Posting No_"",""Document Date"",""External Document No_"",""Payment Method Code"",""No_ Series"",""Posting No_ Series"",""Shipping No_ Series"",""Status"",""Invoice Discount Calculation"",""Invoice Discount Value"",""Allow Line Disc_"",""Due Date"", ""Customer Disc_ Group"",""Applies-to Doc_ Type"",""Applies-to Doc_ No_"",""Reason Code"") Values (" & _
            "3," & SafeSQL(dtr("AcCustCode").ToString) & "," & SafeSQL(dtr("CreditNoteNo").ToString) & "," & SafeSQL(dtr("AcCustCode").ToString) & "," & SafeSQL(dtr("AcCustCode").ToString) & ",'" & Format(dtr("CreditDate"), "yyyy-MM-dd") & "', " & _
            "'" & Format(dtr("CreditDate"), "yyyy-MM-dd") & "','" & Format(dtr("CreditDate"), "yyyy-MM-dd") & "','Credit Memo " & dtr("CreditNoteNo").ToString & "'," & _
            SafeSQL(dtr("PaymentTerms").ToString) & "," & SafeSQL(dtr("ShipmentMethod").ToString) & "," & SafeSQL(dtr("Location").ToString) & "," & SafeSQL("") & "," & SafeSQL(dtr("Dimension2").ToString) & "," & _
            SafeSQL(dtr("CustPostGroup").ToString) & "," & SafeSQL(dtr("CurrencyCode").ToString) & ",1," & SafeSQL(dtr("PriceGroup").ToString) & "," & _
            SafeSQL(dtr("InvDisGroup").ToString) & "," & SafeSQL(dtr("SalesPersonCode").ToString) & ",'',''," & _
            SafeSQL(dtr("CreditNoteNo").ToString) & ",'" & Format(dtr("CreditDate"), "yyyy-MM-dd") & "',''," & SafeSQL("") & "," & _
            SafeSQL(dtr("CreditNoteNo").ToString) & "," & SafeSQL(dtr("CreditNoteNo").ToString) & "," & SafeSQL(dtr("CreditNoteNo").ToString) & "," & _
            "2,1," & dtr("Discount").ToString & ",1" & ",'" & Format(dtr("CreditDate"), "yyyy-MM-dd") & "','',0,'','')")
            '""Prices Including VAT"",
            '""Customer Disc_ Group"",
            Dim iIndex As Integer = 0
            Dim dtr1 As SqlDataReader
            dtr1 = ReadRecordAnother("Select CreditNoteDet.ItemNo, Description, CreditNoteDet.UOM, CreditNoteDet.Qty, CreditNoteDet.Price, AllowInvDiscount, UOM.BaseQty from CreditNoteDet, Item, UOM where UOM.UOM = CreditNoteDet.UOM and CreditNoteDet.ItemNo = UOM.ItemNo and Item.ItemNo = CreditNoteDet.ItemNo and CreditNoteNo = " & SafeSQL(dtr("CreditNoteNo")))
            While dtr1.Read
                iIndex = iIndex + 1
                '1-'Item'
                ExecuteNavSQL("Insert into ""POS Sales Line"" (""Document Type"",""Sell-to Customer No_"",""Document No_"",""Line No_"",""Type"",""No_"",""Location Code"",""Shipment Date"",""Description"",""Unit of Measure"",""Quantity"",""Unit Price"",""Line Discount %"",""Line Discount Amount"",""Allow Invoice Disc_"",""Customer Price Group"",""Job No_"",""Qty_ per Unit of Measure"",""Unit of Measure Code"",""GST %"",""Shortcut Dimension 1 Code"",""Shortcut Dimension 2 Code"") Values (" & _
                "3," & SafeSQL(dtr("AcCustCode")) & "," & SafeSQL(dtr("CreditNoteNo")) & "," & iIndex * 10000 & ",2," & _
                SafeSQL(dtr1("ItemNo")) & "," & SafeSQL(dtr("Location")) & ",'" & Format(dtr("CreditDate"), "yyyy-MM-dd") & "'," & SafeSQL(dtr1("Description")) & "," & SafeSQL(dtr1("UOM")) & "," & dtr1("Qty") & "," & dtr1("Price") & ",0,0," & _
                IIf(dtr1("AllowInvDiscount"), 1, 0) & "," & SafeSQL(dtr("PriceGroup")) & ",''," & dtr1("BaseQty") & "," & SafeSQL(dtr1("Uom")) & "," & dtr("GST") & "," & SafeSQL("") & ",'')")
                '""VAT %"",
                '" & dtr("GST") & ",
            End While
            dtr1.Close()
        End While
        dtr.Close()
        'MsgBox("1")
        ExecuteSQL("Update CreditNote Set Exported = 1 where Exported = 0")
        'MsgBox("OK")
    End Sub

    Private Sub ExportReturn()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select CustId, RtnNo, [Bill-toNo], RtnDt, Return.PayTerms, ShipmentMethod, Location, Dimension1, Dimension2, CustPostGroup, Return.CurCode, Return.CurExRate, PriceGroup, TotalAmt, InvDisGroup, Return.AgentID, PaymentMethod, Discount, GST from Return, Customer where Return.CustId = Customer.CustNo and Return.Exported = 0")
        While dtr.Read
            ExecuteNavSQL("Insert into ""POS Sales Header"" (""Document Type"",""Sell-to Customer No_"",""No_"",""Bill-to Customer No_"",""Ship-to Code"",""Order Date"",""Posting Date"",""Shipment Date"",""Posting Description"",""Payment Terms Code"",""Shipment Method Code"",""Location Code"",""Shortcut Dimension 1 Code"",""Shortcut Dimension 2 Code"",""Customer Posting Group"",""Currency Code"",""Currency Factor"",""Customer Price Group"",""Invoice Disc_ Code"",""Salesperson Code"",""Job No_"",""Shipping No_"",""Posting No_"",""Document Date"",""External Document No_"",""Payment Method Code"",""No_ Series"",""Posting No_ Series"",""Shipping No_ Series"",""Status"",""Invoice Discount Calculation"",""Invoice Discount Value"",""Allow Line Disc_"") Values (" & _
            "'Credit Memo'," & SafeSQL(dtr("CustId").ToString) & "," & SafeSQL(dtr("RtnNo").ToString) & "," & SafeSQL(dtr("Bill-toNo").ToString) & "," & SafeSQL(dtr("CustId").ToString) & ",'" & Format(dtr("RtnDt"), "yyyy-MM-dd HH:mm:ss") & "', " & _
            "'" & Format(dtr("RtnDt"), "yyyy-MM-dd HH:mm:ss") & "','" & Format(dtr("RtnDt"), "yyyy-MM-dd HH:mm:ss") & "','Credit Memo " & dtr("RtnNo").ToString & "'," & _
            SafeSQL(dtr("PayTerms").ToString) & "," & SafeSQL(dtr("ShipmentMethod").ToString) & "," & SafeSQL(dtr("Location").ToString) & "," & SafeSQL(dtr("Dimension1").ToString) & "," & SafeSQL(dtr("Dimension2").ToString) & "," & _
            SafeSQL(dtr("CustPostGroup").ToString) & "," & SafeSQL(dtr("CurCode").ToString) & "," & dtr("CurExRate").ToString & "," & SafeSQL(dtr("PriceGroup").ToString) & "," & _
            SafeSQL(dtr("InvDisGroup").ToString) & "," & SafeSQL(dtr("AgentID").ToString) & ",'',''," & _
            SafeSQL(dtr("RtnNo").ToString) & ",'" & Format(dtr("RtnDt"), "yyyy-MM-dd HH:mm:ss") & "',''," & SafeSQL(dtr("PaymentMethod").ToString) & "," & _
            SafeSQL(dtr("RtnNo").ToString) & "," & SafeSQL(dtr("RTNNo").ToString) & "," & SafeSQL(dtr("RtnNo").ToString) & "," & _
            "'Released','%'," & dtr("Discount").ToString & ",1)")

            '""Prices Including VAT"",
            '""Customer Disc_ Group"",
            Dim iIndex As Integer = 0
            Dim dtr1 As SqlDataReader
            dtr1 = ReadRecordAnother("Select RtnItem.ItemNo, Description, RtnItem.UOM, RtnItem.Qty, RtnItem.Price, DisPer, DisPr, AllowInvDiscount, RtnItem.BaseQty from RtnItem, Item where Item.ItemNo = RtnItem.ItemNo and RtnNo = " & SafeSQL(dtr("RtnNo")))
            While dtr1.Read
                iIndex = iIndex + 1
                ExecuteNavSQL("Insert into ""POS Sales Line"" (""Document Type"",""Sell-to Customer No_"",""Document No_"",""Line No_"",""Type"",""No_"",""Location Code"",""Shipment Date"",""Description"",""Unit of Measure"",""Quantity"",""Unit Price"",""Line Discount %"",""Line Discount Amount"",""Allow Invoice Disc_"",""Customer Price Group"",""Job No_"",""Qty_ per Unit of Measure"",""Unit of Measure Code"",""GST %"") Values (" & _
                "'Credit Memo'," & SafeSQL(dtr("CustId")) & "," & SafeSQL(dtr("RtnNo")) & "," & iIndex * 10000 & ",'Item'," & _
                SafeSQL(dtr1("ItemNo")) & "," & SafeSQL(dtr("Location")) & ",'" & Format(dtr("RtnDt"), "yyyy-MM-dd HH:mm:ss") & "'," & SafeSQL(dtr1("Description")) & "," & SafeSQL(dtr1("UOM")) & "," & dtr1("Qty") & "," & dtr1("Price") & "," & _
                dtr1("DisPer") & "," & dtr1("DisPr") & "," & IIf(dtr1("AllowInvDiscount"), 1, 0) & "," & SafeSQL(dtr("PriceGroup")) & ",''," & dtr1("BaseQty") & "," & SafeSQL(dtr1("Uom")) & "," & dtr("GST") & ")")
                '""VAT %"",
                '" & dtr("GST") & ",
            End While
            dtr1.Close()
        End While
        dtr.Close()
        ExecuteSQL("Update return Set Exported = 1 where Exported = 0")
    End Sub

    Private Sub ExportData()
        Try
            '   Dim dtr1 As OdbcDataReader
            'Dim bProcess As Boolean = False
            'dtr1 = ReadNavRecord("Select ""POS Posting Status"" from ""Sales & Receivables Setup"" ")
            'If dtr1.Read Then
            '    If CInt(dtr1("POS Posting Status").ToString) <> 0 Then
            '        dtr1.Close()
            '        DisconnectNavDB()
            '        DisconnectAnotherDB()
            '        DisconnectDB()
            '        Me.Close()
            '    Else
            '        bProcess = True
            '    End If
            'End If
            'dtr1.Close()
            'MsgBox(bProcess)
            'If bProcess = True Then
            '    ExecuteNavSQL("Update ""Sales & Receivables Setup"" Set [POS Posting Status] = 2")
            '    ExecuteNavSQL("Update ""Sales & Receivables Setup"" Set [POS Transfer Start Time] ='" & Format(Date.Now, "yyyyMMdd HH:mm:ss") & "'")
            'End If
            'MsgBox("1")
            UpdateSystem()
            Dim dtr As SqlDataReader
            dtr = ReadRecord("Select * from system")
            If dtr.Read Then
                sCustPostGroup = dtr("CustPostGroup").ToString
                sGenPostGroup = dtr("GenPostGroup").ToString
                sGSTPostGroup = dtr("GSTPostGroup").ToString
                sGenJournalTemplate = dtr("GenJournalTemplate").ToString
                sGenJournalBatch = dtr("GenJournalBatch").ToString
                sWorkSheetTemplate = dtr("WorkSheetTemplate").ToString
                sJournalBatch = dtr("JournalBatch").ToString
                sItemJnlTemplate = dtr("ItemJnlTemplate").ToString
                sItemJnlBatch = dtr("ItemJnlBatch").ToString
            End If
            dtr.Close()
            'MsgBox("2")
            'dtr = ReadRecord("Select * from system")
            'If dtr.Read Then
            '    sCustPostGroup = dtr("CustPostGroup").ToString
            '    sGenPostGroup = dtr("GenPostGroup").ToString
            '    sGSTPostGroup = dtr("GSTPostGroup").ToString
            '    sGenJournalTemplate = dtr("GenJournalTemplate").ToString
            '    sGenJournalBatch = dtr("GenJournalBatch").ToString
            '    sWorkSheetTemplate = dtr("WorkSheetTemplate").ToString
            '    sJournalBatch = dtr("JournalBatch").ToString
            '    sItemJnlTemplate = dtr("ItemJnlTemplate").ToString
            '    sItemJnlBatch = dtr("ItemJnlBatch").ToString
            'End If
            'dtr.Close()
            '      MsgBox("1")
            'ExportCustomer()
            '     MsgBox("2")
            ExportInvoices()
            'MsgBox("3")
            ExportInvItem()
            'MsgBox("4")
            '  ExportVoidInvoices()
            ' ExportVoidInvItem()
            'ExportOrders()
            'ExportOrdItem()
            ExportCreditNote()
            Dim iIndex As Integer = 0
            ' iIndex = GetItemTransLastNo()
            'ExportItemTrans(iIndex)
            'iIndex = GetReceiptLastNo()
            'ExportReceipt(iIndex)
            'ExportVoidReceipt(iIndex)
            'LoadExportCombo()
            '     ExecuteNavSQL("Update ""Sales & Receivables Setup"" Set [POS Posting Status] = 3")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        MsgBox("Export Completed")
        'MsgBox(rMgr.GetString("Msg_ExpComp"))
    End Sub

    Private Sub btnEx_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEx.Click
        btnEx.Enabled = False
        btnIm.Enabled = False
        ConnectNavDB()
        ConnectAnotherDB()
        Dim ArrList As New ArrayList
        cnt = 0
        ArrList.Add("ItemTrans")
        ArrList.Add("Receipt")
        'ArrList.Add("Orders")
        ArrList.Add("Invoices")
        ArrList.Add("Credit Note")
        dgvStatus.Rows.Clear()
        Dim fname As String = Application.StartupPath & "\" & "Cross.bmp"
        Dim imgTick As New System.Drawing.Bitmap(fname)
        For i = 0 To ArrList.Count - 1
            dgvStatus.Rows.Add(ArrList.Item(i).ToString, imgTick)
        Next
        ExportData()
        DisconnectNavDB()
        DisconnectAnotherDB()
        btnEx.Enabled = True
        btnIm.Enabled = True
    End Sub

    Public Sub ExportDatatoNavisionAuto()
        If bStatus = True Then Exit Sub
        ConnectNavDB()
        ConnectAnotherDB()
        'MsgBox("1")
        ExportData()
        'MsgBox("2")
        DisconnectNavDB()
        DisconnectAnotherDB()
        'MsgBox("3")
        Me.Close()
        'MsgBox("4")
    End Sub
    Private Sub dgvStatus_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvStatus.CellContentClick

    End Sub

    Private Sub Imex_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub

    Private Sub Imex_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '  ExportDatatoNavisionAuto()
        '
        ConnectDB()
        LoadExportCombo()
        'ImportData()
        'ImportDataAuto()
        'Me.Close()
    End Sub

    Private Sub LoadExportCombo()
        Dim dtr As SqlDataReader
        Dim arr As New ArrayList
        Dim bFound As Boolean = False
        dtr = ReadRecord("Select InvDt from Invoice where Exported = 0 and void = 0 union Select RcptDt from Receipt where exported = 0 union Select Docdt from ItemTrans where (DocType = 'GIN' or DocType = 'GOUT' or DocType = 'GVAR' or DocType = 'EX') and exported = 0 union Select CreditDate from CreditNote where exported = 0")
        While dtr.Read
            bFound = False
            For iIndex As Integer = 0 To arr.Count - 1
                If arr(iIndex).ToString = Format(dtr(0), "dd-MMM-yyyy") Then
                    bFound = True
                    Exit For
                End If
            Next
            If bFound = False Then
                arr.Add(Format(dtr(0), "dd-MMM-yyyy"))
            End If
        End While
        dtr.Close()
        cmbExport.Items.Clear()
        For iIndex As Integer = 0 To arr.Count - 1
            cmbExport.Items.Add(arr(iIndex).ToString)
        Next
    End Sub

    Private Function IsRecInvExists() As Boolean
        Dim IsExist As Boolean = False
        Dim rs As SqlDataReader
        rs = ReadRecord("Select * from Invoice where Exported = 0")
        If rs.Read = True Then
            IsExist = True
        End If
        rs.Close()
        If IsExist = False Then Return True
        Dim dtr As OdbcDataReader
        dtr = ReadNavRecord("Select * from ""POS Sales Header""")
        If dtr.Read = True Then
            IsExist = True
        Else
            IsExist = False
        End If
        dtr.Close()
        If IsExist = False Then Return True
        If MsgBox(rMgr.GetString("Msg_DeleteExistingData"), MessageBoxButtons.YesNo, rMgr.GetString("Msg_DeleteExistingData_Cap")) = MsgBoxResult.Yes Then
            ExecuteNavSQL("Delete from ""POS Sales Header""")
            ExecuteNavSQL("Delete from ""POS Sales Line""")
            Return True
        Else
            Return False
        End If

    End Function

    Private Function IsRecReceiptExists() As Boolean
        Dim IsExist As Boolean = False
        Dim rs As SqlDataReader
        rs = ReadRecord("Select * from Receipt where Exported = 0")
        If rs.Read = True Then
            IsExist = True
        End If
        rs.Close()

        rs = ReadRecord("Select * from AdvReceipt where Exported = 0")
        If rs.Read = True Then
            IsExist = True
        End If
        rs.Close()
        If IsExist = False Then Return True
        If IsExist = True Then
            Dim dtr As OdbcDataReader
            dtr = ReadNavRecord("Select * from ""POS Gen_ Journal Line""")
            If dtr.Read = True Then
                IsExist = True
            Else
                IsExist = False
            End If
            dtr.Close()
        End If
        If IsExist = False Then Return True
        If MsgBox(rMgr.GetString("Msg_DeleteExistingData"), MessageBoxButtons.YesNo, rMgr.GetString("Msg_DeleteExistingData_Cap")) = MsgBoxResult.Yes Then
            ExecuteNavSQL("Delete from ""POS Gen_ Journal Line""")
            Return True
        Else
            Return False
        End If
    End Function

    Private Function IsRecItTransExists() As Boolean
        Dim IsExist As Boolean = False
        Dim rs As SqlDataReader
        rs = ReadRecord("Select * from ItemTrans where Exported = 0")
        If rs.Read = True Then
            IsExist = True
        End If
        rs.Close()
        If IsExist = True Then
            Dim dtr As OdbcDataReader
            dtr = ReadNavRecord("Select * from ""POS Item Journal Line""")
            If dtr.Read = True Then
                IsExist = True
            Else
                IsExist = False
            End If
            dtr.Close()
        End If
        If IsExist = False Then Return True
        If MsgBox(rMgr.GetString("Msg_DeleteExistingData"), MessageBoxButtons.YesNo, rMgr.GetString("Msg_DeleteExistingData_Cap")) = MsgBoxResult.Yes Then
            ExecuteNavSQL("Delete from ""POS Item Journal Line""")
            Return True
        Else
            Return False
        End If
    End Function

    Private Function IsRecCrNoteExists() As Boolean
        Dim IsExist As Boolean = False
        Dim rs As SqlDataReader
        rs = ReadRecord("Select * from CreditNote where Exported = 0")
        If rs.Read = True Then
            IsExist = True
        End If
        rs.Close()
        If IsExist = True Then
            Dim dtr As OdbcDataReader
            dtr = ReadNavRecord("Select * from ""POS Sales Header"" where ""Document Type""='Credit Memo'")
            If dtr.Read = True Then
                IsExist = True
            Else
                IsExist = False
            End If
            dtr.Close()
        End If
        If IsExist = False Then Return True
        If MsgBox(rMgr.GetString("Msg_DeleteExistingData"), MessageBoxButtons.YesNo, rMgr.GetString("Msg_DeleteExistingData_Cap")) = MsgBoxResult.Yes Then
            ExecuteNavSQL("Delete from ""POS Sales Header"" where ""Document Type""='Credit Memo'")
            ExecuteNavSQL("Delete from ""POS Sales Line"" where ""Document Type""='Credit Memo'")
            Return True
        Else
            Return False
        End If
    End Function

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
    'Localization
    Private Sub Localization()
        Try
            cInfo = New CultureInfo(sLang)
            Thread.CurrentThread.CurrentCulture = cInfo
            Thread.CurrentThread.CurrentUICulture = cInfo
            rMgr = New ResourceManager("ImExNavision.ImExNavision", [Assembly].GetExecutingAssembly())

            Me.Text = rMgr.GetString("ImEx")
            btnEx.Text = rMgr.GetString("btnEx")
            btnIm.Text = rMgr.GetString("btnIm")
            lblExportDate.Text = rMgr.GetString("lblExportDate")
            dgvStatus.Columns("Tables").HeaderText = rMgr.GetString("Col_Action")
            dgvStatus.Columns("Status").HeaderText = rMgr.GetString("Col_Status")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub SetCulture(ByVal CultureName As String) Implements SalesInterface.MobileSales.ISalesBase.SetCulture
        sLang = CultureName
        Localization()
    End Sub

    Private Sub btnImpInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImpInvoice.Click
        bStatus = True
        btnIm.Enabled = False
        btnEx.Enabled = False
        btnImpInvoice.Enabled = False
        ConnectNavDB()
        ConnectAnotherDB()
        ImportCreditNote()
        ImportCustomer()
        ImportItemPrice()
        'ImportItemPricePromotion()
        'ImportCustDiscGroup()
        'ImportCustDiscPromotion()
        ImportCustInvDiscount()
        ImportInvoice()
        ImportCustProd()
        bStatus = False
        DisconnectNavDB()
        DisconnectAnotherDB()
        ExecuteSQL("Update System set LastImExDate = " & SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")))
        btnIm.Enabled = True
        btnEx.Enabled = True
        btnImpInvoice.Enabled = True
        MsgBox("Import Completed")
    End Sub


    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If MsgBox("Do you want to Delete the Temporary Tables in Navision?", MsgBoxStyle.YesNo, "Delete?") = MsgBoxResult.Yes Then
            Try
                ConnectNavDB()
                ExecuteNavSQL("Delete from ""POS Sales Header""")
                ExecuteNavSQL("Delete from ""POS Sales Line""")
                DisconnectNavDB()
                MsgBox("Deleted! Click Export to Navision again", MsgBoxStyle.OkOnly, "Records Deleted")
            Catch ex As Exception
                MsgBox(ex.Message)
                DisconnectNavDB()
            End Try
        End If
    End Sub
End Class
