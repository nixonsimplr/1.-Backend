Imports System
Imports System.Resources
Imports System.Threading
Imports System.Reflection
Imports System.Globalization
Imports SalesInterface.MobileSales
Imports System.Data.SqlClient
Imports DataInterface.IbizDO

Public Class PO
    Implements ISalesBase


#Region "Localization Variants"
    Private rMgr As ResourceManager
    Private cInfo As CultureInfo
    Private sLang As String
#End Region

    Dim delDate As String
    Dim dTotalAmt As Double
    Private objDO As New DataInterface.IbizDO
    Private myAdapter As SqlDataAdapter
    Private myDataSet As DataSet
    Private myDataView As DataView
    Private RowNo As Integer
    Private IsNewRecord As Boolean = False
    Private IsModify As Boolean = True
    Private dGSTPercent As Double
    Public sItemId, spItemId As String
    Public sUOM As String
    Public sItemPr As String
    Dim dDis As Double
    Public sItemName As String
    Public sCustNo As String
    Public sCustName As String
    Public sPayTerms As String
    Public sAgent As String
    Public sZone, sPrGroup, sLoc As String
    Public dFOCQty, dDisPrice, dDiscount As Double
    Private sMDT As String = ""
    Dim itCnt As Integer = 0
    Dim PrCnt As Integer = 0
    Dim IsAdded As Boolean = False
    Public dQty As Double
    Private rInd As Integer
    Private ArrPromoCount As ArrayList
    Private bIndex0 As Boolean = False
    Private bIndex6 As Boolean = False
    Private bIndex5 As Boolean = False
    Private bChange As Boolean = False
    Private bView As Boolean = False
    Private IsFormLoading As Boolean = False

    Public Structure InvPromoDet
        Dim PromoId As String
        Dim CatBased As Boolean
        Dim ItemCondition As Boolean
        Dim MinAmt As Double
        Dim MaxAmt As Double
        Dim DisPer As Double
        Dim DisAmt As Double
        Dim Multiply As String
        Dim Entitle As Integer
        Dim EntitleType As String
    End Structure

    Public Structure ItemPromoDet
        Dim PromoId As String
        Dim MinQty As Double
        Dim MaxQty As Double
        Dim Priority As Integer
        Dim ItemId As String
        Dim UOM As String
        Dim Multiply As String
    End Structure

    Public Structure PromoDet
        Dim PromoId As String
        Dim PromoName As String
        Dim PromoType As String
        Dim CatBased As Boolean
        Dim ItemCondition As Boolean
        Dim MinAmt As Double
        Dim MaxAmt As Double
        Dim DisPer As Double
        Dim DisAmt As Double
        Dim Multiply As String
    End Structure

    Public Structure PromoCount
        Dim PromoId As String
        Dim Count As Integer
    End Structure

    Public Structure PromoEntitle
        Dim PromoId As String
        Dim Entitle As Integer
        Dim EntitleType As String
    End Structure

    Private Sub SalesOrder_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        objDO.DisconnectDB()
    End Sub

    Private Sub LoadRecord()
        Try
            If RowNo < 0 Or RowNo >= myDataView.Count Then Exit Sub
            InvNoTextBox.Text = myDataView(RowNo).Item("PONo").ToString
            InvDtDateTimePicker.Value = myDataView(RowNo).Item("PODt")
            CustIdTextBox.Text = myDataView(RowNo).Item("CustID")
            'txtCustName.Text = myDataView(RowNo).Item("OrderNo")
            'PONoTextBox.Text = myDataView(RowNo).Item("Pono").ToString
            PayTermsTextBox.Text = myDataView(RowNo).Item("PayTerms").ToString
            AgentIDTextBox.Text = myDataView(RowNo).Item("AgentID").ToString
            CurCodeTextBox.Text = myDataView(RowNo).Item("CurCode").ToString
            CurExRateTextBox.Text = myDataView(RowNo).Item("CurExRate")
            DiscountLabel.Text = Format(myDataView(RowNo).Item("Discount"), "0.00")
            SubTotalLabel.Text = Format(myDataView(RowNo).Item("SubTotal"), "0.00")
            GSTAmtLabel.Text = Format(myDataView(RowNo).Item("GSTAmt"), "0.00")
            TotalAmtLabel.Text = Format(myDataView(RowNo).Item("TotalAmt"), "0.00")
            IsFormLoading = True
            '            OrdNoTextBox.Text = myDataView(RowNo).Item("OrdNo").ToString
            IsFormLoading = False
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub SalesOrder_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objDO.ConnectionString = My.Settings.ConnectionString
        objDO.ConnectDB()
        GetOrders("select * from PO order by PoNo")
        GetTextBox()
    End Sub

    Private Sub GetOrders(ByVal sSQL As String)
        myAdapter = New SqlDataAdapter(sSQL, My.Settings.ConnectionString)
        myDataSet = New DataSet
        myAdapter.Fill(myDataSet, "PO")
        myDataView = New DataView(myDataSet.Tables("PO"))
        If myDataView.Count = 0 Then
            RowNo = -1
        Else
            RowNo = 0
        End If
        LoadRecord()
        ' Sort the view based on the first column name.
        'myDataView.Sort = "OrdNo"
    End Sub

    Public Sub ListViewClick() Implements SalesInterface.MobileSales.ISalesBase.ListViewClick

    End Sub

    Public Sub MoveFirstClick() Implements SalesInterface.MobileSales.ISalesBase.MoveFirstClick
        If IsNewRecord = True Then
            '    If MsgBox(rMgr.GetString("Msg_RecordNotSaved"), MsgBoxStyle.YesNo, rMgr.GetString("Msg_RecordNotSaved_Cap")) = MsgBoxResult.No Then
            Exit Sub
            '   End If
        End If
        RowNo = 0
        LoadRecord()
    End Sub

    Public Sub MoveLastClick() Implements SalesInterface.MobileSales.ISalesBase.MoveLastClick
        If IsNewRecord = True Then
            '      If MsgBox(rMgr.GetString("Msg_RecordNotSaved"), MsgBoxStyle.YesNo, rMgr.GetString("Msg_RecordNotSaved_Cap")) = MsgBoxResult.No Then
            Exit Sub
            'End If
        End If
        RowNo = myDataView.Count - 1
        LoadRecord()
    End Sub

    Public Sub MoveNextClick() Implements SalesInterface.MobileSales.ISalesBase.MoveNextClick
        If IsNewRecord = True Then
            '   If MsgBox(rMgr.GetString("Msg_RecordNotSaved"), MsgBoxStyle.YesNo, rMgr.GetString("Msg_RecordNotSaved_Cap")) = MsgBoxResult.No Then
            Exit Sub
            'End If
        End If
        If RowNo >= myDataView.Count - 1 Then Exit Sub
        RowNo = RowNo + 1
        LoadRecord()
    End Sub

    Public Sub MovePositionClick(ByVal Position As Long) Implements SalesInterface.MobileSales.ISalesBase.MovePositionClick

    End Sub

    Public Sub MovePreviousClick() Implements SalesInterface.MobileSales.ISalesBase.MovePreviousClick
        If IsNewRecord = True Then
            '   If MsgBox(rMgr.GetString("Msg_RecordNotSaved"), MsgBoxStyle.YesNo, rMgr.GetString("Msg_RecordNotSaved_Cap")) = MsgBoxResult.No Then
            Exit Sub
            'End If
        End If
        If RowNo <= 0 Then Exit Sub
        RowNo = RowNo - 1
        LoadRecord()
    End Sub

    Public Sub Print() Implements SalesInterface.MobileSales.ISalesBase.Print

    End Sub

    Public Sub Print(ByVal PageNo As Integer) Implements SalesInterface.MobileSales.ISalesBase.Print

    End Sub

    Private Sub ClearData()
        InvDtDateTimePicker.Value = Date.Now
        '   OrdNoTextBox.Text = ""
        CustIdTextBox.Text = String.Empty
        CustNameTextBox.Text = String.Empty
        ' PONoTextBox.Text = String.Empty
        PayTermsTextBox.Text = String.Empty
        AgentIDTextBox.Text = String.Empty
        CurCodeTextBox.Text = String.Empty
        CurExRateTextBox.Text = "0"
        '  OrdNoTextBox.Text = String.Empty
        dgvOrdItem.Rows.Clear()
        dgvOrdItem.ReadOnly = False
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            'InvoiceBindingSource.AddNew()
            ClearData()
            Dim sInvNo As String = ""

            Dim rs As SqlDataReader
            rs = objDO.ReadRecord("Select MDTNo, GST from System")
            If rs.Read = True Then
                sMDT = rs("MDTNo")
                dGSTPercent = rs("GST")
            End If
            rs.Close()
            rs = objDO.ReadRecord("Select * from MDT where MdtNo = " & objDO.SafeSQL(sMDT) & "")
            If rs.Read = True Then
                Dim sPre As String = rs("PrePONo")
                Dim iLen As Int32 = rs("LenPONo")
                Dim iOrdNo As Int32 = rs("LastPONo") + 1
                sInvNo = sPre & StrDup(iLen - Len(sPre) - Len(CStr(iOrdNo)), "0") & CStr(iOrdNo)
                sLoc = rs("Location").ToString
            End If
            rs.Close()
            InvNoTextBox.Text = sInvNo
            DiscountLabel.Text = "0.00"
            SubTotalLabel.Text = "0.00"
            GSTAmtLabel.Text = "0.00"
            TotalAmtLabel.Text = "0.00"
            IsNewRecord = True
            dgvOrdItem.Rows.Clear()
            ArrPromoCount = New ArrayList
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        '   btnNew.Enabled = False
        '   btnSave.Enabled = True
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        CheckInvoicePromotion(TotalAmtLabel.Text)
        Save()
    End Sub

    Public Sub ShipInvoice()
        Try
            Dim rs As SqlDataReader
            Dim sInvNo As String = ""
            rs = objDO.ReadRecord("Select MDTNo from System")
            If rs.Read = True Then
                sMDT = rs("MDTNo")
            End If
            rs.Close()
            rs = objDO.ReadRecord("Select * from MDT where MdtNo = " & objDO.SafeSQL(sMDT) & "")
            If rs.Read = True Then
                Dim sPre As String = rs("PrePoNo")
                Dim iLen As Int32 = rs("LenPoNo")
                Dim iOrdNo As Int32 = rs("LastPoNo") + 1
                sInvNo = sPre & StrDup(iLen - Len(sPre) - Len(CStr(iOrdNo)), "0") & CStr(iOrdNo)
            End If
            rs.Close()
            ' GetOrderData()
            rs = objDO.ReadRecord("Select CustName from Customer where Custno = '" & CustIdTextBox.Text & "'")
            If rs.Read = True Then
                CustNameTextBox.Text = rs("CustName").ToString
            End If
            rs.Close()
            Dim dAdvAmt As Double = 0
            'If OrdNoTextBox.Text <> "" Then
            '    rs = objDO.ReadRecord("Select Sum(Amount) as SumAmt from AdvReceipt where OrdNo = " & objDO.SafeSQL(OrdNoTextBox.Text) & "")
            '    If rs.Read = True Then
            '        If rs("SumAmt").ToString <> "" Then dAdvAmt = rs("SumAmt")
            '    End If
            '    rs.Close()
            'End If
            objDO.ExecuteSQL("Insert into PO (PoNo, PoDt, CustId, AgentId, Discount, SubTotal, GSTAmt, TotalAmt, PayTerms, CurCode, CurExRate,Void,PaidAmt, GST, Exported,DTG) Values (" & objDO.SafeSQL(sInvNo) & ",'" & Format(InvDtDateTimePicker.Value, "dd-MMM-yyyy") & "'," & objDO.SafeSQL(CustIdTextBox.Text) & ", " & objDO.SafeSQL(AgentIDTextBox.Text) & ", " & DiscountLabel.Text & ", " & SubTotalLabel.Text & ", " & GSTAmtLabel.Text & ", " & TotalAmtLabel.Text & ", " & objDO.SafeSQL(PayTermsTextBox.Text) & ", " & objDO.SafeSQL(CurCodeTextBox.Text) & ", " & CurExRateTextBox.Text & ",0," & CStr(dAdvAmt) & "," & dGSTPercent & ",0" & "," & objDO.SafeSQL(Format(Date.Now, "yyyyMMdd 00:00:00")).ToString & ")")
            Dim iIndex As Int16
            For iIndex = 1 To dgvOrdItem.Rows.Count
                If Not dgvOrdItem.Item(0, iIndex - 1).Value Is Nothing Then
                    '"Insert into OrdItem (OrdNo, ItemNo, [LineNo], VariantCode, Description, Uom, Qty, Foc, Price, DisPer, DisPr, Discount, SubAmt, GSTAmt, DeliQty, Location, DeliveryDate, PromoId, PromoOffer, Priority) Values (" & objDO.SafeSQL(OrdNoTextBox.Text) & ", " & objDO.SafeSQL(dgvOrdItem.Item(0, iIndex - 1).Value.ToString) & "," & iIndex & "," & objDO.SafeSQL(dgvOrdItem.Item(1, iIndex - 1).Value.ToString) & "," & objDO.SafeSQL(dgvOrdItem.Item(2, iIndex - 1).Value.ToString) & ", " & objDO.SafeSQL(dgvOrdItem.Item(4, iIndex - 1).Value.ToString) & ", " & dgvOrdItem.Item(5, iIndex - 1).Value.ToString & ", " & dgvOrdItem.Item(10, iIndex - 1).Value.ToString & ", " & dgvOrdItem.Item(6, iIndex - 1).Value.ToString & ", 0,0,0," & dgvOrdItem.Item(7, iIndex - 1).Value.ToString & ", " & CDbl(dgvOrdItem.Item(7, iIndex - 1).Value.ToString) * dGSTPercent / 100 & ",0," & objDO.SafeSQL(dgvOrdItem.Item(3, iIndex - 1).Value.ToString) & "," & objDO.SafeSQL(dgvOrdItem.Item(11, iIndex - 1).Value.ToString) & "," & objDO.SafeSQL(dgvOrdItem.Item(12, iIndex - 1).Value.ToString) & "," & objDO.SafeSQL(dgvOrdItem.Item(13, iIndex - 1).Value.ToString) & "," & dgvOrdItem.Item(14, iIndex - 1).Value.ToString & ")"
                    objDO.ExecuteSQL("Insert into POItem (PoNo, ItemNo, [LineNo], VariantCode, Description, Uom, Qty, Foc, Price, DisPer, DisPr, Discount, SubAmt, GSTAmt, DeliQty, Location, DeliveryDate, PromoId, PromoOffer, Priority) Values (" & objDO.SafeSQL(sInvNo) & ", " & objDO.SafeSQL(dgvOrdItem.Item(0, iIndex - 1).Value.ToString) & "," & iIndex & "," & objDO.SafeSQL(dgvOrdItem.Item(1, iIndex - 1).Value.ToString) & "," & objDO.SafeSQL(dgvOrdItem.Item(2, iIndex - 1).Value.ToString) & ", " & objDO.SafeSQL(dgvOrdItem.Item(4, iIndex - 1).Value.ToString) & ", " & dgvOrdItem.Item(5, iIndex - 1).Value.ToString & ",0, " & dgvOrdItem.Item(6, iIndex - 1).Value.ToString & ", " & dgvOrdItem.Item(8, iIndex - 1).Value.ToString & ", " & dgvOrdItem.Item(7, iIndex - 1).Value.ToString & ", " & dgvOrdItem.Item(19, iIndex - 1).Value.ToString & ", " & dgvOrdItem.Item(9, iIndex - 1).Value.ToString & ", " & CDbl(dgvOrdItem.Item(9, iIndex - 1).Value.ToString) * dGSTPercent / 100 & ", " & dgvOrdItem.Item(5, iIndex - 1).Value.ToString & ", " & objDO.SafeSQL(dgvOrdItem.Item(3, iIndex - 1).Value.ToString) & ", " & objDO.SafeSQL(dgvOrdItem.Item(13, iIndex - 1).Value.ToString) & ", " & objDO.SafeSQL(dgvOrdItem.Item(14, iIndex - 1).Value.ToString) & ", " & objDO.SafeSQL(dgvOrdItem.Item(15, iIndex - 1).Value.ToString) & ", " & dgvOrdItem.Item(16, iIndex - 1).Value.ToString & ")")
                End If
            Next
            MsgBox("PO Generated Successfully")
            'MessageBox.Show(rMgr.GetString("Msg_InvAdded"), rMgr.GetString("Msg_InvAdded_Cap"), MessageBoxButtons.OK)
            objDO.ExecuteSQL("Update MDT set LastPoNo = LastPoNo + 1 where MDTNo = " & objDO.SafeSQL(sMDT))
            'objDO.ExecuteSQL("Update MDT set LastOrdNo = LastOrdNo + 1 where MDTNo = " & objDO.SafeSQL(sMDT))
            'If OrdNoTextBox.Text <> "" Then
            '    objDO.ExecuteSQL("Update OrderHdr set Delivered = 1 where OrdNo = " & objDO.SafeSQL(OrdNoTextBox.Text))
            '    objDO.ExecuteSQL("Update OrdItem set DeliQty = Qty where OrdNo = " & objDO.SafeSQL(OrdNoTextBox.Text))
            'End If
            IsNewRecord = False
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub Save()
        If IsNewRecord = True Then
            ShipInvoice()
            Dim stP As PromoCount
            For iCnt As Integer = ArrPromoCount.Count - 1 To 0 Step -1
                stP = ArrPromoCount(iCnt)
                SavePromoEntitlement(InvNoTextBox.Text, Date.Now, CustIdTextBox.Text, stP.PromoId, stP.Count)
            Next
            'LoadReport(InvNoTextBox.Text)
            GetOrders("select * from PO order by PoNo")
            IsNewRecord = False
            RowNo = 0
            'LoadRow()
            'Else
            '    ModifyOrder()
        End If
        ' btnNew.Enabled = True
        'IsModify = False

        '    btnSave.Enabled = False
        'btnModify.Enabled = True
    End Sub



    Private Sub CustIdTextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustIdTextBox.TextChanged
        If CustIdTextBox.Text = "" Then Exit Sub
        If IsFormLoading = True Then Exit Sub
        'If IsNewRecord = False Then Exit Sub
        Try
            Dim rs As SqlDataReader
            'rs = objDO.ReadRecord("Select ExpiryDate, CustType from Customer where CustNo=" & objDO.SafeSQL(CustIdTextBox.Text))
            'If rs.Read = True Then
            '    If rs("CustType") = "Member" Then
            '        If rs("ExpiryDate") < Date.Now Then
            '            'MsgBox(rMgr.GetString("Msg_MemExpired"), MsgBoxStyle.OkOnly, "Information")
            '            MsgBox("Membership Expired", MsgBoxStyle.OkOnly, "Information")
            '            CustIdTextBox.Text = ""
            '        End If
            '    End If
            'End If
            'rs.Close()
            rs = objDO.ReadRecord("Select CustName, SalesAgent, PaymentTerms, CurrencyCode, PriceGroup,ZoneCode from Customer where Custno = '" & CustIdTextBox.Text & "'")
            If rs.Read = True Then
                CustNameTextBox.Text = rs("CustName").ToString
                sPrGroup = rs("PriceGroup").ToString.ToString
                sZone = rs("ZoneCode").ToString
                If AgentIDTextBox.Text.Trim = "" Then AgentIDTextBox.Text = rs("SalesAgent").ToString
                If IsNewRecord = True Then
                    AgentIDTextBox.Text = rs("SalesAgent").ToString
                    PayTermsTextBox.Text = rs("PaymentTerms").ToString
                    CurCodeTextBox.Text = rs("CurrencyCode").ToString
                End If
            End If
            rs.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        IsNewRecord = False
        ' IsModify = False
        'ClearData()
        RowNo = 0
        LoadRecord()
        '  btnSave.Enabled = False
        '  btnNew.Enabled = True
        btnModify.Enabled = True
    End Sub

    Private Sub dgvOrdItem_CellValueChanged_1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Exit Sub
        Try
            If e.ColumnIndex = 2 Then
                'Dim dtr As SqlDataReader
                'Dim str As String = "Select BaseQty from Uom where ItemNo = " & objDO.SafeSQL(dgvOrdItem.Item(0, e.RowIndex).Value.ToString) & " and Uom = " & objDO.SafeSQL(dgvOrdItem.Item(2, e.RowIndex).Value.ToString)
                ' ConnectAnotherDB()
                'dtr = objDO.ReadRecordAnother(str)
                'If dtr.Read = True Then
                '    dgvOrdItem.Item(3, e.RowIndex).Value = CStr(CDbl(dgvOrdItem.Item(8, e.RowIndex).Value.ToString) * dtr("BaseQty"))
                'End If
                'dtr.Close()
                '   DisconnectAnotherDB()
            ElseIf e.ColumnIndex = 6 Then
                If dgvOrdItem.Item(6, e.RowIndex).Value.ToString = "" Or dgvOrdItem.Item(5, e.RowIndex).Value = "" Then Exit Sub
                dgvOrdItem.Item(7, e.RowIndex).Value = Format(CDbl(dgvOrdItem.Item(6, e.RowIndex).Value.ToString) * CDbl(dgvOrdItem.Item(5, e.RowIndex).Value.ToString), "0.00")
            ElseIf e.ColumnIndex = 5 Then
                If dgvOrdItem.Item(6, e.RowIndex).Value.ToString = "" Or dgvOrdItem.Item(5, e.RowIndex).Value.ToString = "" Then Exit Sub
                dgvOrdItem.Item(7, e.RowIndex).Value = Format(CDbl(dgvOrdItem.Item(6, e.RowIndex).Value.ToString) * CDbl(dgvOrdItem.Item(5, e.RowIndex).Value.ToString), "0.00")
            ElseIf e.ColumnIndex = 6 Then
                Calculate()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub CalculateDiscount(ByVal DisPer As Double, ByVal DisAmt As Double)
        Dim dNet As Double
        Dim dGst As Double
        Dim iIndex As Int16
        dNet = 0
        dGst = 0

        For iIndex = 1 To dgvOrdItem.Rows.Count
            dNet = dNet + CDbl(dgvOrdItem.Item(7, iIndex - 1).Value)
        Next
        dDiscount = DisAmt
        dDiscount += (dNet - DisAmt) * DisPer / 100
        DiscountLabel.Text = "$" & Format(dDiscount, "0.00")
        dGst = (dNet - dDiscount) * dGSTPercent / 100
        SubTotalLabel.Text = "" & Format(dNet, "0.00")
        GSTAmtLabel.Text = "$" & Format(dGst, "0.00")
        TotalAmtLabel.Text = "$" & Format((dNet - dDiscount) + dGst, "0.00")
    End Sub

    Private Sub Calculate()
        Dim dNet As Double
        Dim dGst As Double
        Dim iIndex As Int16
        dNet = 0
        dGst = 0

        For iIndex = 1 To dgvOrdItem.Rows.Count
            dNet = dNet + CDbl(dgvOrdItem.Item(9, iIndex - 1).Value)
            dDis = dDis + CDbl(dgvOrdItem.Item(19, iIndex - 1).Value)
        Next
        dGst = dNet * dGSTPercent / 100
        DiscountLabel.Text = "" & Format(dDis, "0.00")
        SubTotalLabel.Text = "" & Format(dNet, "0.00")
        GSTAmtLabel.Text = "" & Format(dGst, "0.00")
        TotalAmtLabel.Text = "" & Format(dNet + dGst, "0.00")
        '  InvPromotion(cnt)
    End Sub

    'Private Sub SaveOrder()
    '    objDO.ExecuteSQL("Insert into Invoice (OrdNo, OrdDt, CustId, PONo, AgentId, Discount, SubTotal, GSTAmt, TotalAmt, PayTerms, CurCode, CurExRate,Delivered,void,exported) Values (" & objDO.SafeSQL(OrdNoTextBox.Text) & ", '" & Format(InvDtDateTimePicker.Value, "dd-MMM-yyyy") & "', " & objDO.SafeSQL(CustIdTextBox.Text) & ", " & objDO.SafeSQL(PONoTextBox.Text) & ", " & objDO.SafeSQL(AgentIDTextBox.Text) & ", " & DiscountLabel.Text & ", " & SubTotalLabel.Text & ", " & GSTAmtLabel.Text & ", " & TotalAmtLabel.Text & ", " & objDO.SafeSQL(PayTermsTextBox.Text) & ", " & objDO.SafeSQL(CurCodeTextBox.Text) & ", " & CurExRateTextBox.Text & ",0,0,0)")
    '    objDO.ExecuteSQL("Update MDT set LastOrdNo = LastordNo + 1 where MDTNo = " & objDO.SafeSQL(sMDT))
    'End Sub

    Private Sub UpdateOrder()
        'objDO.ExecuteSQL("Update Invoice Set OrdDt = '" & Format(OrdDtDateTimePicker.Value, "dd-MM-yyyy") & "', CustId = " & objDO.SafeSQL(CustIdTextBox.Text) & ", PONo = " & objDO.SafeSQL(PONoTextBox.Text) & ", AgentId = " & objDO.SafeSQL(AgentIDTextBox.Text) & ", Discount = " & DiscountLabel.Text & ", SubTotal = " & SubTotalLabel.Text & ", GSTAmt = " & GSTAmtLabel.Text & ", TotalAmt = " & TotalAmtLabel.Text & ", PayTerms = " & objDO.SafeSQL(PayTermsTextBox.Text) & ", CurCode = " & objDO.SafeSQL(CurCodeTextBox.Text) & ", CurExRate = " & CurExRateTextBox.Text & " where OrdNo = " & objDO.SafeSQL(OrdNoTextBox.Text))
        'objDO.ExecuteSQL("Delete from OrdItem where OrdNo = " & objDO.SafeSQL(OrdNoTextBox.Text))
    End Sub

    'Private Sub SaveOrderItem()
    '    Dim dDate As Date
    '    Dim iIndex As Int16
    '    Dim str As String
    '    For iIndex = 1 To dgvOrdItem.Rows.Count
    '        'If PromotionalPrice(objDO.SafeSQL(dgvOrdItem.Item(0, iIndex - 1).Value.ToString), dgvOrdItem.Item(4, iIndex - 1).Value) Then
    '        'End If
    '        If Not dgvOrdItem.Item(0, iIndex - 1).Value Is Nothing Or Not dgvOrdItem.Item(2, iIndex - 1).Value Is Nothing Then
    '            str = "Insert into OrdItem (OrdNo, ItemNo, [LineNo], VariantCode, Description, Uom, Qty, Foc, Price, DisPer, DisPr, Discount, SubAmt, GSTAmt, DeliQty, Location, DeliveryDate, PromoId, PromoOffer, Priority) Values (" & objDO.SafeSQL(OrdNoTextBox.Text) & ", " & objDO.SafeSQL(dgvOrdItem.Item(0, iIndex - 1).Value.ToString) & "," & iIndex & "," & objDO.SafeSQL(dgvOrdItem.Item(1, iIndex - 1).Value.ToString) & "," & objDO.SafeSQL(dgvOrdItem.Item(2, iIndex - 1).Value.ToString) & ", " & objDO.SafeSQL(dgvOrdItem.Item(4, iIndex - 1).Value.ToString) & ", " & dgvOrdItem.Item(5, iIndex - 1).Value.ToString & ", " & dgvOrdItem.Item(10, iIndex - 1).Value.ToString & ", " & dgvOrdItem.Item(6, iIndex - 1).Value.ToString & ", 0,0,0," & dgvOrdItem.Item(7, iIndex - 1).Value.ToString & ", " & CDbl(dgvOrdItem.Item(7, iIndex - 1).Value.ToString) * dGSTPercent / 100 & ",0," & objDO.SafeSQL(dgvOrdItem.Item(3, iIndex - 1).Value.ToString) & "," & objDO.SafeSQL(dgvOrdItem.Item(11, iIndex - 1).Value.ToString) & "," & objDO.SafeSQL(dgvOrdItem.Item(12, iIndex - 1).Value.ToString) & "," & objDO.SafeSQL(dgvOrdItem.Item(13, iIndex - 1).Value.ToString) & "," & dgvOrdItem.Item(14, iIndex - 1).Value.ToString & ")"
    '            objDO.ExecuteSQL(str)
    '        End If
    '    Next
    '    'MessageBox.Show("New Order Created", "New Record", MessageBoxButtons.OK)
    'End Sub

    Private Sub btnCustNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCustNo.Click
        RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\DataList.dll", "DataList.Customer", "SalesPromo.PO", "txtCustNo", 0, 0)
    End Sub

    Private Sub btnPay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPay.Click
        RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\DataList.dll", "DataList.PaymentTerms", "SalesPromo.PO", "txtPayTerms", 0, 0)
    End Sub

    Private Sub btnAgent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAgent.Click
        RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\DataList.dll", "DataList.Agent", "SalesPromo.PO", "txtSalesPersonCode", 0, 0)
    End Sub

    Public Function GetListViewForm() As String Implements SalesInterface.MobileSales.ISalesBase.GetListViewForm
        RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\SalesPromo.dll", "SalesPromo.POList", "SalesPromo.POList", "InvNoTextBox", 0, 0)
        Return Windows.Forms.Application.StartupPath & "\SalesPromo.dll,SalesPromo.POList"
    End Function

    Public Event LoadForm(ByVal Location As String, ByVal LoadType As String) Implements SalesInterface.MobileSales.ISalesBase.LoadForm

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnVoid.Click
        Dim StrSql As String
        Try
            StrSql = "Delete from PO where PONo=" & objDO.SafeSQL(InvNoTextBox.Text)
            objDO.ExecuteSQL(StrSql)
            MsgBox("Deleted Successfully")
            '   MessageBox.Show(rMgr.GetString("Msg_RecVoided"), rMgr.GetString("Msg_InvAdded_Cap"), MessageBoxButtons.OK)
            ClearData()
            GetOrders("select * from PO order by PONo")
            RowNo = 0
            LoadRecord()
            '   objDO.ExecuteSQL("Update MDT set LastExNo=" & iExNo & " where MdtNo = " & objDO.SafeSQL(sMDT))
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub


    Private Sub dgvOrdItem_CellMouseDown1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvOrdItem.CellMouseDown
        If e.RowIndex < 0 Then Exit Sub
        If e.ColumnIndex < 0 Then Exit Sub
        If e.ColumnIndex = 0 Then
            rInd = e.RowIndex
            If CustIdTextBox.Text = String.Empty Then
                MsgBox("Select Customer No")
                Me.CustIdTextBox.Select()
                Return
            End If
            If IsNewRecord = True Then
                If e.ColumnIndex = 0 Then
                    RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\DataList.dll", "DataList.ItemList", "SalesPromo.PO", "dgvOrdItem", 0, 0)
                End If
            End If
        Else
            dgvOrdItem.CurrentCell = dgvOrdItem.Item(e.ColumnIndex, e.RowIndex)
            dgvOrdItem.BeginEdit(True)
        End If
    End Sub

    Private Sub dgvOrdItem_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvOrdItem.CellValueChanged
        If bView = True Then Exit Sub
        If e.RowIndex < 0 Then Exit Sub
        If dgvOrdItem.Item(0, e.RowIndex).Value <> "" Then
            If e.ColumnIndex = 0 Then
                bIndex0 = True
                bChange = True
                Dim bExplode As Boolean = False
                Dim dtr As SqlDataReader
                Dim cb1 As New DataGridViewComboBoxCell
                cb1 = dgvOrdItem.Item(1, e.RowIndex)
                dgvOrdItem.Item(1, e.RowIndex).Value = ""
                cb1.Items.Clear()
                Dim bCur1 As Boolean = False
                'MsgBox(dgvOrdItem.Item(0, e.RowIndex).Value)
                dtr = objDO.ReadRecord("Select Code from ItemVariant where ItemNo=" & objDO.SafeSQL(dgvOrdItem.Item(0, e.RowIndex).Value))
                While dtr.Read
                    cb1.Items.Add(dtr("Code"))
                    bCur1 = True
                End While
                dtr.Close()
                Dim sBaseUOM As String = ""
                dtr = objDO.ReadRecord("Select * from Item where ItemNo=" & objDO.SafeSQL(dgvOrdItem.Item(0, e.RowIndex).Value))
                If dtr.Read = True Then
                    '   dgvOrdItem.Item(1, e.RowIndex).Value = ""
                    dgvOrdItem.Item(2, e.RowIndex).Value = dtr("Description").ToString
                    '                dgvOrdItem.Item(2, e.RowIndex).Value = dtr("BaseUOM").ToString
                    dgvOrdItem.Item(6, e.RowIndex).Value = dtr("UnitPrice").ToString
                    bExplode = dtr("AssemblyBOM")
                    sBaseUOM = dtr("BaseUOM").ToString
                End If
                dtr.Close()
                'If bExplode = True Then
                '    dgvOrdItem.Item(6, e.RowIndex).Value = "1"
                '    dtr = objDO.ReadRecord("Select * from AssemblyBOM where ParentItemNo=" & objDO.SafeSQL(dgvOrdItem.Item(0, e.RowIndex).Value))
                '    While dtr.Read
                '        Dim row As String() = New String() {"", "", "  " & dtr("Qty").ToString & " x " & dtr("Description").ToString, "", "", _
                '        "0", "0", "0", "", "", "0", "", "", "", "0", "0", "0"}
                '        dgvOrdItem.Rows.Add(row)
                '    End While
                '    dtr.Close()
                'End If
                Dim cb As New DataGridViewComboBoxCell
                cb = dgvOrdItem.Item(4, e.RowIndex)
                dgvOrdItem.Item(4, e.RowIndex).Value = ""
                cb.Items.Clear()
                Dim bCur As Boolean = False
                'MsgBox(dgvOrdItem.Item(0, e.RowIndex).Value)
                dtr = objDO.ReadRecord("Select Uom from UOM where ItemNo=" & objDO.SafeSQL(dgvOrdItem.Item(0, e.RowIndex).Value))
                While dtr.Read
                    cb.Items.Add(dtr("Uom"))
                    bCur = True
                End While
                dtr.Close()
                Dim cntno As Integer = 0
                Dim scUOM As String = ""
                dgvOrdItem.Item(7, e.RowIndex).Value = "0"
                dgvOrdItem.Item(8, e.RowIndex).Value = "0"
                dgvOrdItem.Item(9, e.RowIndex).Value = "0"
                dgvOrdItem.Item(11, e.RowIndex).Value = sItemPr
                dgvOrdItem.Item(12, e.RowIndex).Value = "0"
                dgvOrdItem.Item(13, e.RowIndex).Value = ""
                dgvOrdItem.Item(14, e.RowIndex).Value = ""
                dgvOrdItem.Item(15, e.RowIndex).Value = "0"
                dgvOrdItem.Item(16, e.RowIndex).Value = "0"
                dgvOrdItem.Item(19, e.RowIndex).Value = "0"
                'Dim cb2 As New DataGridViewComboBoxCell
                'cb2 = dgvOrdItem.Item(3, e.RowIndex)
                'dgvOrdItem.Item(3, e.RowIndex).Value = ""
                'cb2.Items.Clear()
                'dtr = objDO.ReadRecord("Select code from Location") ' where ItemNo=" & objDO.SafeSQL(sItemId))
                'While dtr.Read
                '    'MsgBox(dtr("code"))
                '    cb2.Items.Add(dtr("code"))
                '    '  bCur = True
                'End While
                'dtr.Close()
                dgvOrdItem.Item(13, e.RowIndex).Value = Format(Date.Now, "dd/MMM/yyyy")
                bIndex0 = False
                dgvOrdItem.Item(3, e.RowIndex).Value = sLoc
                dgvOrdItem.Item(4, e.RowIndex).Value = sBaseUOM
                '   DisconnectAnotherDB()
            ElseIf e.ColumnIndex = 1 Then
                If dgvOrdItem.Item(0, e.RowIndex).Value = "" Then Exit Sub
                If dgvOrdItem.Item(1, e.RowIndex).Value = "" Then Exit Sub
                Dim dtr As SqlDataReader
                dtr = objDO.ReadRecordAnother("Select description from ItemVariant where ItemNo=" & objDO.SafeSQL(dgvOrdItem.Item(0, e.RowIndex).Value.ToString) & " and Code=" & objDO.SafeSQL(dgvOrdItem.Item(1, e.RowIndex).Value.ToString))
                If dtr.Read = True Then
                    dgvOrdItem.Item(2, e.RowIndex).Value = dtr("Description").ToString
                End If
                dtr.Close()
            ElseIf e.ColumnIndex = 3 Then
                GoTo CheckWH
            ElseIf e.ColumnIndex = 6 Then
                bIndex6 = True
                bChange = True
                If CDbl(dgvOrdItem.Item(6, e.RowIndex).Value).ToString = "" Or CDbl(dgvOrdItem.Item(5, e.RowIndex).Value = "").ToString Then Exit Sub
                dTotalAmt = CDbl(dgvOrdItem.Item(6, e.RowIndex).Value.ToString) * CDbl(dgvOrdItem.Item(5, e.RowIndex).Value.ToString)
                dgvOrdItem.Item(19, e.RowIndex).Value = Format(dTotalAmt * (CDbl(dgvOrdItem.Item(8, e.RowIndex).Value.ToString) / 100) + CDbl(dgvOrdItem.Item(7, e.RowIndex).Value.ToString), "0.00")
                dgvOrdItem.Item(9, e.RowIndex).Value = Format(dTotalAmt - (dTotalAmt * (CDbl(dgvOrdItem.Item(8, e.RowIndex).Value.ToString) / 100)) - CDbl(dgvOrdItem.Item(7, e.RowIndex).Value.ToString), "0.00")
                bIndex6 = False
            ElseIf e.ColumnIndex = 5 Then
CheckWH:
                bIndex5 = True
                bChange = True
                Dim str As String = ""
                If CDbl(dgvOrdItem.Item(5, e.RowIndex).Value) > 0 Then
                    If dgvOrdItem.Item(4, e.RowIndex).Value = "" Then
                        'MessageBox.Show(rMgr.GetString("Msg_SelUOM"), rMgr.GetString("Msg_SelUOM_Cap"), MessageBoxButtons.OK)
                        MsgBox("Select UOM")
                        '  dgvOrdItem.Item(4, e.RowIndex).Value = ""
                        Exit Sub
                    End If

                    'If CDbl(dgvOrdItem.Item(5, e.RowIndex).Value) > dQty Then
                    'Dim dtrQty As SqlDataReader
                    'Dim sLeadTime As String = "0D"
                    'Dim sSafetyLeadTime As String = "0D"
                    'Dim bNonStock As Boolean = False
                    'dtrQty = objDO.ReadRecord("Select NonStockItem, LeadTime, SafetyLeadTime from Item Where ItemNo = " & objDO.SafeSQL(dgvOrdItem.Item(0, e.RowIndex).Value))
                    'If dtrQty.Read = True Then
                    '    bNonStock = dtrQty("NonStockItem")
                    '    If dtrQty("LeadTime").ToString <> "" Then sLeadTime = dtrQty("LeadTime")
                    '    If dtrQty("SafetyLeadTime").ToString <> "" Then sSafetyLeadTime = dtrQty("SafetyLeadTime")
                    'End If
                    'dtrQty.Close()
                    'If bNonStock = True Then
                    '    Dim dActPr As Double
                    '    dActPr = GetPrice(dgvOrdItem.Item(0, e.RowIndex).Value, dgvOrdItem.Item(1, e.RowIndex).Value, dgvOrdItem.Item(5, e.RowIndex).Value)
                    '    dgvOrdItem.Item(6, e.RowIndex).Value = dActPr
                    '    'dgvOrdItem.Item(7, e.RowIndex).Value = Format(CDbl(dgvOrdItem.Item(6, e.RowIndex).Value.ToString) * CDbl(dgvOrdItem.Item(5, e.RowIndex).Value.ToString), "0.00")
                    'Else
                    '    '  MessageBox.Show(rMgr.GetString("Msg_QuntyOnInvent1") & vbCrLf & rMgr.GetString("Msg_QuntyOnInvent2"), rMgr.GetString("Msg_QuntyOnInvent3"), MessageBoxButtons.OK)
                    '    dgvOrdItem.Item(5, e.RowIndex).Value = "0"
                    '    Exit Sub
                    'End If
                    'Dim iDays As Integer = 0
                    'If Mid(sLeadTime, sLeadTime.Length, 1) = "D" Then
                    '    iDays = CInt(Mid(sLeadTime, 1, sLeadTime.Length - 1))
                    'ElseIf Mid(sLeadTime, sLeadTime.Length, 1) = "W" Then
                    '    iDays = CInt(Mid(sLeadTime, 1, sLeadTime.Length - 1)) * 7
                    'End If
                    'If Mid(sSafetyLeadTime, sLeadTime.Length, 1) = "D" Then
                    '    iDays += CInt(Mid(sLeadTime, 1, sLeadTime.Length - 1))
                    'ElseIf Mid(sLeadTime, sLeadTime.Length, 1) = "W" Then
                    '    iDays += CInt(Mid(sLeadTime, 1, sLeadTime.Length - 1)) * 7
                    'End If
                    dgvOrdItem.Item(13, e.RowIndex).Value = Format(Date.Now, "dd/MMM/yyyy")
                    'Else
                    Dim dActPr As Double
                    dActPr = GetPrice(dgvOrdItem.Item(0, e.RowIndex).Value, dgvOrdItem.Item(1, e.RowIndex).Value, dgvOrdItem.Item(5, e.RowIndex).Value)
                    dgvOrdItem.Item(6, e.RowIndex).Value = dActPr
                    'End If
                    'PromotionalPrice(e.RowIndex)
                    'If dgvOrdItem.Rows.Count > e.RowIndex Then
                    '    If dgvOrdItem.Item(6, e.RowIndex + 1).Value = "0.00" Then
                    '        dgvOrdItem.Rows.RemoveAt(e.RowIndex + 1)
                    '    End If
                    'End If
                    If CDbl(dgvOrdItem.Item(6, e.RowIndex).Value).ToString = "" Or CDbl(dgvOrdItem.Item(5, e.RowIndex).Value = "").ToString Then Exit Sub
                    dTotalAmt = CDbl(dgvOrdItem.Item(6, e.RowIndex).Value.ToString) * CDbl(dgvOrdItem.Item(5, e.RowIndex).Value.ToString)
                    dgvOrdItem.Item(19, e.RowIndex).Value = Format(dTotalAmt * (CDbl(dgvOrdItem.Item(8, e.RowIndex).Value.ToString) / 100) + CDbl(dgvOrdItem.Item(7, e.RowIndex).Value.ToString), "0.00")
                    dgvOrdItem.Item(9, e.RowIndex).Value = Format(dTotalAmt - (dTotalAmt * (CDbl(dgvOrdItem.Item(8, e.RowIndex).Value.ToString) / 100)) - CDbl(dgvOrdItem.Item(7, e.RowIndex).Value.ToString), "0.00")
                End If

                bIndex5 = False
            ElseIf e.ColumnIndex = 6 Then
                If Convert.ToString(dgvOrdItem.Item(6, e.RowIndex).Value) = "" Or Convert.ToString(dgvOrdItem.Item(5, e.RowIndex).Value) = "" Then Exit Sub
                dTotalAmt = CDbl(dgvOrdItem.Item(6, e.RowIndex).Value.ToString) * CDbl(dgvOrdItem.Item(5, e.RowIndex).Value.ToString)
                dgvOrdItem.Item(19, e.RowIndex).Value = Format(dTotalAmt * (CDbl(dgvOrdItem.Item(8, e.RowIndex).Value.ToString) / 100) + CDbl(dgvOrdItem.Item(7, e.RowIndex).Value.ToString), "0.00")
                dgvOrdItem.Item(9, e.RowIndex).Value = Format(dTotalAmt - (dTotalAmt * (CDbl(dgvOrdItem.Item(8, e.RowIndex).Value.ToString) / 100)) - CDbl(dgvOrdItem.Item(7, e.RowIndex).Value.ToString), "0.00")

            ElseIf e.ColumnIndex = 7 Then
                If Convert.ToString(dgvOrdItem.Item(6, e.RowIndex).Value) = "" Or Convert.ToString(dgvOrdItem.Item(5, e.RowIndex).Value) = "" Then Exit Sub
                dTotalAmt = CDbl(dgvOrdItem.Item(6, e.RowIndex).Value.ToString) * CDbl(dgvOrdItem.Item(5, e.RowIndex).Value.ToString)
                dgvOrdItem.Item(19, e.RowIndex).Value = Format(dTotalAmt * (CDbl(dgvOrdItem.Item(8, e.RowIndex).Value.ToString) / 100) + CDbl(dgvOrdItem.Item(7, e.RowIndex).Value.ToString), "0.00")
                dgvOrdItem.Item(9, e.RowIndex).Value = Format(dTotalAmt - (dTotalAmt * (CDbl(dgvOrdItem.Item(8, e.RowIndex).Value.ToString) / 100)) - CDbl(dgvOrdItem.Item(7, e.RowIndex).Value.ToString), "0.00")

            ElseIf e.ColumnIndex = 8 Then
                If Convert.ToString(dgvOrdItem.Item(6, e.RowIndex).Value) = "" Or Convert.ToString(dgvOrdItem.Item(5, e.RowIndex).Value) = "" Then Exit Sub
                dTotalAmt = CDbl(dgvOrdItem.Item(6, e.RowIndex).Value.ToString) * CDbl(dgvOrdItem.Item(5, e.RowIndex).Value.ToString)
                dgvOrdItem.Item(19, e.RowIndex).Value = Format(dTotalAmt * (CDbl(dgvOrdItem.Item(8, e.RowIndex).Value.ToString) / 100) + CDbl(dgvOrdItem.Item(7, e.RowIndex).Value.ToString), "0.00")
                dgvOrdItem.Item(9, e.RowIndex).Value = Format(dTotalAmt - (dTotalAmt * (CDbl(dgvOrdItem.Item(8, e.RowIndex).Value.ToString) / 100)) - CDbl(dgvOrdItem.Item(7, e.RowIndex).Value.ToString), "0.00")

            ElseIf e.ColumnIndex = 9 Then
                Calculate()
            End If
        End If
        If e.ColumnIndex = 2 Then
            Try
                If dgvOrdItem.Item(0, e.RowIndex).Value.ToString = "" Then
                    dgvOrdItem.Item(0, e.RowIndex).Value = ""
                    dgvOrdItem.Item(1, e.RowIndex).Value = ""
                    dgvOrdItem.Item(3, e.RowIndex).Value = ""
                    dgvOrdItem.Item(4, e.RowIndex).Value = ""
                    dgvOrdItem.Item(5, e.RowIndex).Value = 0
                    dgvOrdItem.Item(6, e.RowIndex).Value = "0"
                    dgvOrdItem.Item(7, e.RowIndex).Value = "0"
                    dgvOrdItem.Item(8, e.RowIndex).Value = "0"
                    dgvOrdItem.Item(9, e.RowIndex).Value = "0"
                    dgvOrdItem.Item(10, e.RowIndex).Value = ""
                    dgvOrdItem.Item(11, e.RowIndex).Value = "0"
                    dgvOrdItem.Item(12, e.RowIndex).Value = "0"
                    dgvOrdItem.Item(13, e.RowIndex).Value = Date.Now
                    dgvOrdItem.Item(14, e.RowIndex).Value = ""
                    dgvOrdItem.Item(15, e.RowIndex).Value = ""
                    dgvOrdItem.Item(16, e.RowIndex).Value = "0"
                    dgvOrdItem.Item(17, e.RowIndex).Value = "0"
                    dgvOrdItem.Item(19, e.RowIndex).Value = "0"
                End If
            Catch ex As Exception
                dgvOrdItem.Item(0, e.RowIndex).Value = ""
                dgvOrdItem.Item(1, e.RowIndex).Value = ""
                dgvOrdItem.Item(3, e.RowIndex).Value = ""
                dgvOrdItem.Item(4, e.RowIndex).Value = ""
                dgvOrdItem.Item(5, e.RowIndex).Value = 0
                dgvOrdItem.Item(6, e.RowIndex).Value = "0"
                dgvOrdItem.Item(7, e.RowIndex).Value = "0"
                dgvOrdItem.Item(8, e.RowIndex).Value = "0"
                dgvOrdItem.Item(9, e.RowIndex).Value = "0"
                dgvOrdItem.Item(10, e.RowIndex).Value = ""
                dgvOrdItem.Item(11, e.RowIndex).Value = "0"
                dgvOrdItem.Item(12, e.RowIndex).Value = "0"
                dgvOrdItem.Item(13, e.RowIndex).Value = Date.Now
                dgvOrdItem.Item(14, e.RowIndex).Value = ""
                dgvOrdItem.Item(15, e.RowIndex).Value = ""
                dgvOrdItem.Item(16, e.RowIndex).Value = "0"
                dgvOrdItem.Item(17, e.RowIndex).Value = "0"
                dgvOrdItem.Item(19, e.RowIndex).Value = "0"
            End Try
        End If
        If e.ColumnIndex = 0 Or e.ColumnIndex = 5 Or e.ColumnIndex = 6 Then
            If bIndex0 = False And bIndex5 = False And bIndex6 = False And bChange = True Then
                If CDbl(dgvOrdItem.Item(6, e.RowIndex).Value.ToString) <> 0 Then
                    '        CheckPromotion(dgvOrdItem.Item(0, e.RowIndex).Value, dgvOrdItem.Item(4, e.RowIndex).Value, dgvOrdItem.Item(5, e.RowIndex).Value, GetCategory(dgvOrdItem.Item(0, e.RowIndex).Value))
                End If
                bChange = False
            End If
        End If
        ' Catch ex As Exception
        ' End Try
    End Sub

    Public Event LoadDataForm(ByVal Location As String, ByVal LoadType As String, ByVal ParentLoadType As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.LoadDataForm

    Public Event ResultData(ByVal ChildLoadType As String, ByVal Value As String) Implements SalesInterface.MobileSales.ISalesBase.ResultData

    Public Sub ReturnedData(ByVal Value As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.ReturnedData
        Try
            If ResultTo = "InvNoTextBox" Then
                GetOrders("select * from PO where PONo = '" & Value & "'")
                GetTextBox()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        If ResultTo = "txtSalesPersonCode" Then
            AgentIDTextBox.Text = Value
        End If
        If ResultTo = "txtCustNo" Then
            CustIdTextBox.Text = Value
        End If
        If ResultTo = "txtPayTerms" Then
            PayTermsTextBox.Text = Value
        End If
        'If ResultTo = "txtOrdNo" Then
        '    OrdNoTextBox.Text = Value
        'End If
        If ResultTo = "dgvOrdItem" Then
            dgvOrdItem.Item(0, rInd).Value = Value
        End If
        'If ResultTo = "OrdNoTextBox" Then
        '    OrdNoTextBox.Text = Value
        'End If
    End Sub

    Public Function GetSpecialPrice(ByVal sItemId As String) As Double
        Dim dtr As SqlDataReader
        Dim dPr As Double = 0
        dtr = objDO.ReadRecord("SELECT PromoOffer.DisPrice FROM PromoApply, Promotion, PromoCondition, PromoOffer WHERE Promotion.PromoName = 'Special Price' and PromoApply.PromoID = Promotion.PromoID AND PromoApply.PromoID = PromoCondition.PromoID AND PromoCondition.PromoID = PromoOffer.PromoID AND (PromoCondition.ItemID = " & objDO.SafeSQL(sItemId) & ") AND (PromoOffer.ItemID = " & objDO.SafeSQL(sItemId) & ") AND (PromoApply.ID = " & objDO.SafeSQL(sPrGroup) & ") AND (Promotion.FromDate <= '" & Format(Date.Now, "yyyyMMdd") & "') AND (Promotion.ToDate >= '" & Format(Date.Now, "yyyyMMdd") & "') ORDER BY PromoOffer.DisPrice ASC")
        If dtr.Read Then
            dPr = dtr("DisPrice")
        End If
        dtr.Close()
        Return dPr
    End Function


    Public Function GetPrice(ByVal sItemNo As String, ByVal sVariant As String, ByVal Qty As Double) As Double
        Dim dtr As SqlDataReader
        Dim dPr1 As Double = 0
        Dim dPr2 As Double = 0
        Dim dPr3 As Double = 0
        Dim dPr As Double = 0

        dPr = GetSpecialPrice(sItemNo)
        If dPr > 0 Then
            Return dPr
        End If
        dtr = objDO.ReadRecord("Select UnitPrice from ItemPr where ItemNo = '" & sItemNo & "' and VariantCode = '" & sVariant & "' and PriceGroup = '" & sPrGroup & "' and SalesType = 'Customer Price Group' and MinQty <= " & Qty & " order by MinQty Desc")
        If dtr.Read Then
            dPr1 = dtr("UnitPrice")
            dPr = dPr1
        End If
        dtr.Close()
        dtr = objDO.ReadRecord("Select UnitPrice from ItemPr where ItemNo = '" & sItemNo & "' and VariantCode = '" & sVariant & "' and PriceGroup = '" & CustIdTextBox.Text & "' and SalesType = 'Customer' and MinQty <= " & Qty & " order by MinQty Desc")
        If dtr.Read Then
            dPr2 = dtr("UnitPrice")
            If dPr = 0 Then
                dPr = dPr2
            ElseIf dPr2 < dPr And dPr2 > 0 Then
                dPr = dPr2
            End If
        End If
        dtr.Close()
        dtr = objDO.ReadRecord("Select UnitPrice from ItemPr where ItemNo = '" & sItemNo & "' and VariantCode = '" & sVariant & "' and SalesType = 'All Customers' and MinQty <= " & Qty & " order by MinQty Desc")
        If dtr.Read Then
            dPr3 = dtr("UnitPrice")
            If dPr = 0 Then
                dPr = dPr3
            ElseIf dPr3 < dPr And dPr3 > 0 Then
                dPr = dPr3
            End If
        End If
        dtr.Close()
        If dPr = 0 Then
            dtr = objDO.ReadRecord("Select UnitPrice from Item where ItemNo = '" & sItemNo & "'")
            If dtr.Read Then
                dPr = dtr("UnitPrice")
            End If
            dtr.Close()
        End If
        Return dPr
    End Function

    '    Private Sub CheckPromotion(ByVal sItemId As String, ByVal sUOM As String, ByVal dQty As Double, ByVal sCat As String)
    '        Dim dtr As SqlDataReader
    '        Dim dtrOffer As SqlDataReader
    '        Dim Arr As New ArrayList
    '        dtr = objDO.ReadRecord("Select Promotion.PromoId, Entitle, EntitleType from Promotion, PromoCondition, PromoApply Where PromoName <> 'Special Price' and PromoType = 'Item Promotion'  and (Promotion.PromoId = PromoCondition.PromoId And Promotion.PromoId = PromoApply.PromoId) and FromDate <= '" & Format(Date.Now, "yyyyMMdd") & "' and ToDate >= '" & Format(Date.Now, "yyyyMMdd") & "' and (ItemId = '" & sItemId & "' or ItemId = '" & sCat & "') and Uom = '" & sUOM & "' and ((Promotion.ApType = 'All') or (PromoApply.Id = '" & CustIdTextBox.Text & "' and Promotion.ApType = 'Customer') or (PromoApply.Id = '" & sZone & "' and Promotion.ApType = 'Zone') or (PromoApply.Id = '" & AgentIDTextBox.Text & "' and Promotion.ApType = 'Agent') or (PromoApply.Id = '" & sPrGroup & "' and Promotion.ApType = 'Price Group')) order by Priority")
    '        While dtr.Read
    '            Dim stPromo As PromoEntitle
    '            stPromo.PromoId = dtr("PromoId")
    '            stPromo.Entitle = dtr("Entitle")
    '            stPromo.EntitleType = dtr("EntitleType")
    '            Arr.Add(stPromo)
    '        End While
    '        dtr.Close()
    '        If Arr.Count = 0 Then Exit Sub
    '        For iIndex As Integer = 0 To Arr.Count - 1
    '            Dim stPromo As PromoEntitle
    '            stPromo = Arr(iIndex)
    '            Dim arrCnt As New ArrayList
    '            Dim bCondition As Boolean = False
    '            Dim bMultiply As Boolean = False
    '            Dim iMultiCnt As Integer = 1
    '            Dim iMulti As Integer = 1
    '            Dim iPriority As Integer
    '            Dim bItemPromo As Boolean = False
    '            Dim bCatPromo As Boolean = False
    '            Dim iPCount As Integer = 0
    '            'Checking Items
    '            If stPromo.Entitle > 0 Then
    '                If stPromo.EntitleType = "Per Day" Then
    '                    dtr = objDO.ReadRecord("Select Sum(PromoCount) as TotalPromoApplied from PromoEntitlement where PromoId = " & objDO.SafeSQL(stPromo.PromoId) & " and OrderDate = '" & Format(Date.Now, "yyyyMMdd") & "'")
    '                    If dtr.Read Then
    '                        If dtr("TotalPromoApplied").ToString <> "" Then
    '                            iPCount = dtr("TotalPromoApplied")
    '                        End If
    '                    End If
    '                    dtr.Close()
    '                ElseIf stPromo.EntitleType = "Per Week" Then
    '                    dtr = objDO.ReadRecord("Select Sum(PromoCount) as TotalPromoApplied from PromoEntitlement where PromoId = " & objDO.SafeSQL(stPromo.PromoId) & " and OrderDate >= '" & Format(DateAdd(DateInterval.Day, Date.Now.DayOfWeek * -1, Date.Now), "yyyyMMdd") & "'")
    '                    If dtr.Read Then
    '                        If dtr("TotalPromoApplied").ToString <> "" Then
    '                            iPCount = dtr("TotalPromoApplied")
    '                        End If
    '                    End If
    '                    dtr.Close()
    '                ElseIf stPromo.EntitleType = "Per Month" Then
    '                    dtr = objDO.ReadRecord("Select Sum(PromoCount) as TotalPromoApplied from PromoEntitlement where PromoId = " & objDO.SafeSQL(stPromo.PromoId) & " and OrderDate >= '" & Format(Date.Now, "yyyy") & Format(Date.Now, "MM") & "01" & "'")
    '                    If dtr.Read Then
    '                        If dtr("TotalPromoApplied").ToString <> "" Then
    '                            iPCount = dtr("TotalPromoApplied")
    '                        End If
    '                    End If
    '                    dtr.Close()
    '                ElseIf stPromo.EntitleType = "Per Promotion" Then
    '                    dtr = objDO.ReadRecord("Select Sum(PromoCount) as TotalPromoApplied from PromoEntitlement where PromoId = " & objDO.SafeSQL(stPromo.PromoId))
    '                    If dtr.Read Then
    '                        If dtr("TotalPromoApplied").ToString <> "" Then
    '                            iPCount = dtr("TotalPromoApplied")
    '                        End If
    '                    End If
    '                    dtr.Close()
    '                End If
    '                If iPCount >= stPromo.Entitle Then GoTo NextRecord
    '            End If
    '            dtr = objDO.ReadRecord("Select Promotion.PromoId, Promotion.PromoName, Priority, PromoCondition.ItemId, ItemName, PromoCondition.UOM, MinQty, MaxQty, Multiply, Entitle, EntitleType from PromoCondition, Promotion, Item where Promotion.PromoId = PromoCondition.PromoId and PromoCondition.ItemId = Item.ItemNo and PromoCondition.LineType = 'Item' and Promotion.PromoId = '" & stPromo.PromoId & "'")
    '            While dtr.Read = True
    '                bItemPromo = True
    '                bCondition = False
    '                iPriority = dtr("Priority")
    '                If dtr("Multiply") = "Incremental" Then
    '                    bMultiply = True
    '                End If
    '                'dgvOrdItem
    '                For iCnt As Integer = 0 To dgvOrdItem.Rows.Count - 1
    '                    If dtr("ItemId").ToString = dgvOrdItem.Item(0, iCnt).Value.ToString And dtr("Uom").ToString = dgvOrdItem.Item(4, iCnt).Value.ToString Then
    '                        If dtr("MinQty") > CDbl(dgvOrdItem.Item(5, iCnt).Value.ToString) Or dtr("MaxQty") < CDbl(dgvOrdItem.Item(5, iCnt).Value.ToString) Then
    '                            bCondition = False
    '                            Exit While
    '                        Else
    '                            If CDbl(dgvOrdItem.Item(6, iCnt).Value.ToString) = 0 Then
    '                                bCondition = False
    '                                Exit While
    '                            End If
    '                            If dgvOrdItem.Item(14, iCnt).Value.ToString <> "" Then
    '                                If CDbl(dgvOrdItem.Item(14, iCnt).Value.ToString) < dtr("Priority") Then
    '                                    bCondition = False
    '                                    Exit While
    '                                End If
    '                            End If
    '                            If bMultiply = True Then
    '                                iMulti = System.Math.Floor(CDbl(dgvOrdItem.Item(5, iCnt).Value.ToString) / dtr("MinQty"))
    '                                iMultiCnt = iMulti
    '                            End If
    '                            bCondition = True
    '                            arrCnt.Add(iCnt)
    '                        End If
    '                    End If
    '                Next
    '                If bCondition = False Then
    '                    Exit While
    '                End If
    '            End While
    '            dtr.Close()
    '            If bItemPromo = True And bCondition = False Then GoTo NextRecord
    '            ' Checking Category
    '            'TODO: Current Problem If multiple promotion with same item in the category will not be given promotion
    '            dtr = objDO.ReadRecord("Select Promotion.PromoId, Promotion.PromoName, Priority, PromoCondition.ItemId, PromoCondition.UOM, MinQty, MaxQty, Multiply  from PromoCondition, Promotion where Promotion.PromoId = PromoCondition.PromoId and PromoCondition.LineType = 'Category' and Promotion.PromoId = '" & stPromo.PromoId & "'")
    '            While dtr.Read = True
    '                bCatPromo = True
    '                bCondition = False
    '                iPriority = dtr("Priority")
    '                If dtr("Multiply") = "Incremental" Then
    '                    bMultiply = True
    '                End If
    '                Dim dPQty As Double = 0
    '                For iCnt As Integer = 0 To dgvOrdItem.Rows.Count - 1
    '                    If dtr("ItemId").ToString = GetCategory(dgvOrdItem.Item(0, iCnt).Value.ToString) And dtr("Uom").ToString = dgvOrdItem.Item(4, iCnt).Value.ToString And CDbl(dgvOrdItem.Item(6, iCnt).Value.ToString) > 0 Then
    '                        If dgvOrdItem.Item(14, iCnt).Value.ToString <> "" Then
    '                            If CDbl(dgvOrdItem.Item(14, iCnt).Value.ToString) >= dtr("Priority") Then
    '                                dPQty += CDbl(dgvOrdItem.Item(5, iCnt).Value.ToString)
    '                                arrCnt.Add(iCnt)
    '                            End If
    '                        Else
    '                            dPQty += CDbl(dgvOrdItem.Item(5, iCnt).Value.ToString)
    '                            arrCnt.Add(iCnt)
    '                        End If
    '                    End If
    '                Next
    '                If dtr("MinQty") > dPQty Or dtr("MaxQty") < dPQty Then
    '                    bCondition = False
    '                    Exit While
    '                End If
    '                bCondition = True
    '                If bMultiply = True Then
    '                    iMulti = System.Math.Floor(CDbl(dPQty) / dtr("MinQty"))
    '                    If iMulti < iMultiCnt Then iMultiCnt = iMulti
    '                End If
    '                If bCondition = False Then
    '                    Exit While
    '                End If
    '            End While
    '            dtr.Close()
    '            If bCatPromo = True And bCondition = False Then GoTo NextRecord
    '            ' Checking Promo Group
    '            Dim bPromoGroup As Boolean = False
    '            Dim arrGroup As New ArrayList
    '            'TODO: Current Problem If multiple promotion with same item in the category will not be given promotion
    '            dtr = objDO.ReadRecord("Select Promotion.PromoId, Promotion.PromoName, Priority, PromoCondition.ItemId, PromoCondition.UOM, MinQty, MaxQty, Multiply  from PromoCondition, Promotion where Promotion.PromoId = PromoCondition.PromoId and PromoCondition.LineType = 'Promotion Group' and Promotion.PromoId = '" & stPromo.PromoId & "'")
    '            While dtr.Read = True
    '                Dim ipdGroup As New ItemPromoDet
    '                ipdGroup.PromoId = dtr("PromoId")
    '                ipdGroup.Priority = dtr("Priority")
    '                ipdGroup.ItemId = dtr("ItemId")
    '                ipdGroup.UOM = dtr("Uom")
    '                ipdGroup.MinQty = dtr("MinQty")
    '                ipdGroup.MaxQty = dtr("MaxQty")
    '                ipdGroup.Multiply = dtr("Multiply")
    '                arrGroup.Add(ipdGroup)
    '            End While
    '            dtr.Close()

    '            For iCount As Integer = 0 To arrGroup.Count - 1
    '                Dim ipdGroup As New ItemPromoDet
    '                ipdGroup = arrGroup(iCount)
    '                bPromoGroup = True
    '                Dim dtrTemp As SqlDataReader
    '                dtrTemp = objDO.ReadRecord("Select ItemId  from PromoGroup where GroupId = '" & ipdGroup.ItemId & "'")
    '                Dim dPQty As Double = 0
    '                While dtrTemp.Read
    '                    bCondition = False
    '                    iPriority = ipdGroup.Priority
    '                    If ipdGroup.Multiply = "Incremental" Then
    '                        bMultiply = True
    '                    End If
    '                    For iCnt As Integer = 0 To dgvOrdItem.Rows.Count - 1
    '                        If dtrTemp("ItemId") = dgvOrdItem.Item(0, iCnt).Value.ToString And ipdGroup.UOM = dgvOrdItem.Item(4, iCnt).Value.ToString And CDbl(dgvOrdItem.Item(6, iCnt).Value.ToString) > 0 Then
    '                            If dgvOrdItem.Item(14, iCnt).Value.ToString <> "" Then
    '                                If CDbl(dgvOrdItem.Item(14, iCnt).Value.ToString) >= ipdGroup.Priority Then
    '                                    dPQty += CDbl(dgvOrdItem.Item(5, iCnt).Value.ToString)
    '                                    arrCnt.Add(iCnt)
    '                                End If
    '                            Else
    '                                dPQty += CDbl(dgvOrdItem.Item(5, iCnt).Value.ToString)
    '                                arrCnt.Add(iCnt)
    '                            End If
    '            If dtr("TotalPromoApplied").ToString <> "" Then
    '                iPCount = dtr("TotalPromoApplied")
    '            End If
    '        End If
    '        dtr.Close()
    '    ElseIf stPromo.EntitleType = "Per Promotion" Then
    '        dtr = objDO.ReadRecord("Select Sum(PromoCount) as TotalPromoApplied from PromoEntitlement where PromoId = " & objDO.SafeSQL(stPromo.PromoId))
    '        If dtr.Read Then
    '            If dtr("TotalPromoApplied").ToString <> "" Then
    '                iPCount = dtr("TotalPromoApplied")
    '            End If
    '        End If
    '        dtr.Close()
    '    End If
    '    If iPCount >= stPromo.Entitle Then GoTo NextRecord
    'End If
    'dtr = objDO.ReadRecordAnother("Select Promotion.PromoId, Promotion.PromoName, Priority, PromoCondition.ItemId, ItemName, PromoCondition.UOM, MinQty, MaxQty, Multiply, Entitle, EntitleType from PromoCondition, Promotion, Item where Promotion.PromoId = PromoCondition.PromoId and PromoCondition.ItemId = Item.ItemNo and PromoCondition.LineType = 'Item' and Promotion.PromoId = '" & stPromo.PromoId & "'")
    'While dtr.Read = True
    '    bItemPromo = True
    '    bCondition = False
    '    iPriority = dtr("Priority")
    '    If dtr("Multiply") = "Incremental" Then
    '        bMultiply = True
    '    End If
    '    'dgvOrdItem
    '    For iCnt As Integer = 0 To dgvOrdItem.Rows.Count - 1
    '        If iCnt = dgvOrdItem.NewRowIndex Then Exit For
    '        If dtr("ItemId").ToString = dgvOrdItem.Item(0, iCnt).Value.ToString And dtr("Uom").ToString = dgvOrdItem.Item(4, iCnt).Value.ToString Then
    '            If dtr("MinQty") > CDbl(dgvOrdItem.Item(5, iCnt).Value.ToString) Or dtr("MaxQty") < CDbl(dgvOrdItem.Item(5, iCnt).Value.ToString) Then
    '                bCondition = False
    '                Exit While
    '            Else
    '                If CDbl(dgvOrdItem.Item(6, iCnt).Value.ToString) = 0 Then
    '                    bCondition = False
    '                    Exit While
    '                End If
    '                'Commented for POS Demo Indonesia later take it
    '                If dgvOrdItem.Item(14, iCnt).Value.ToString <> "" Then
    '                    'If CDbl(dgvOrdItem.Item(14, iCnt).Value.ToString) < dtr("Priority") Then
    '                    '    bCondition = False
    '                    '    Exit While
    '                    'End If
    '                End If
    '                If bMultiply = True Then
    '                    iMulti = System.Math.Floor(CDbl(dgvOrdItem.Item(5, iCnt).Value.ToString) / dtr("MinQty"))
    '                    iMultiCnt = iMulti
    '                End If
    '                bCondition = True
    '                arrCnt.Add(iCnt)
    '            End If
    '        End If
    '    Next
    '    If bCondition = False Then
    '        Exit While
    '    End If
    'End While
    'dtr.Close()
    'If bItemPromo = True And bCondition = False Then GoTo NextRecord
    '' Checking Category
    ''TODO: Current Problem If multiple promotion with same item in the category will not be given promotion
    'dtr = objDO.ReadRecordAnother("Select Promotion.PromoId, Promotion.PromoName, Priority, PromoCondition.ItemId, PromoCondition.UOM, MinQty, MaxQty, Multiply  from PromoCondition, Promotion where Promotion.PromoId = PromoCondition.PromoId and PromoCondition.LineType = 'Category' and Promotion.PromoId = '" & stPromo.PromoId & "'")
    'While dtr.Read = True
    '    bCatPromo = True
    '    bCondition = False
    '    iPriority = dtr("Priority")
    '    If dtr("Multiply") = "Incremental" Then
    '        bMultiply = True
    '    End If
    '    Dim dPQty As Double = 0
    '    For iCnt As Integer = 0 To dgvOrdItem.Rows.Count - 1
    '        If iCnt = dgvOrdItem.NewRowIndex Then Exit For
    '        If dtr("ItemId").ToString = GetCategory(dgvOrdItem.Item(0, iCnt).Value.ToString) And dtr("Uom").ToString = dgvOrdItem.Item(4, iCnt).Value.ToString And CDbl(dgvOrdItem.Item(6, iCnt).Value.ToString) > 0 Then
    '            If dgvOrdItem.Item(16, iCnt).Value.ToString <> "" Then
    '                If CDbl(dgvOrdItem.Item(16, iCnt).Value.ToString) >= dtr("Priority") Then
    '                    dPQty += CDbl(dgvOrdItem.Item(5, iCnt).Value.ToString)
    '                    arrCnt.Add(iCnt)
    '                End If
    '            Else
    '                dPQty += CDbl(dgvOrdItem.Item(5, iCnt).Value.ToString)
    '                arrCnt.Add(iCnt)
    '            End If
    '        End If
    '    Next
    '    If dtr("MinQty") > dPQty Or dtr("MaxQty") < dPQty Then
    '        bCondition = False
    '        Exit While
    '    End If
    '    bCondition = True
    '    If bMultiply = True Then
    '        iMulti = System.Math.Floor(CDbl(dPQty) / dtr("MinQty"))
    '        If iMulti < iMultiCnt Then iMultiCnt = iMulti
    '    End If
    '    If bCondition = False Then
    '        Exit While
    '    End If
    'End While
    'dtr.Close()
    'If bCatPromo = True And bCondition = False Then GoTo NextRecord
    '' Checking Promo Group
    'Dim bPromoGroup As Boolean = False
    'Dim arrGroup As New ArrayList
    ''TODO: Current Problem If multiple promotion with same item in the category will not be given promotion
    'dtr = objDO.ReadRecordAnother("Select Promotion.PromoId, Promotion.PromoName, Priority, PromoCondition.ItemId, PromoCondition.UOM, MinQty, MaxQty, Multiply  from PromoCondition, Promotion where Promotion.PromoId = PromoCondition.PromoId and PromoCondition.LineType = 'Promotion Group' and Promotion.PromoId = '" & stPromo.PromoId & "'")
    'While dtr.Read = True
    '    Dim ipdGroup As New ItemPromoDet
    '    ipdGroup.PromoId = dtr("PromoId")
    '    ipdGroup.Priority = dtr("Priority")
    '    ipdGroup.ItemId = dtr("ItemId")
    '    ipdGroup.UOM = dtr("Uom")
    '    ipdGroup.MinQty = dtr("MinQty")
    '    ipdGroup.MaxQty = dtr("MaxQty")
    '    ipdGroup.Multiply = dtr("Multiply")
    '    arrGroup.Add(ipdGroup)
    'End While
    'dtr.Close()

    '                For iCount As Integer = 0 To arrGroup.Count - 1
    '                    Dim ipdGroup As New ItemPromoDet
    '                    ipdGroup = arrGroup(iCount)
    '                    bPromoGroup = True
    '                    Dim dtrTemp As SqlDataReader
    '                    dtrTemp = objDO.ReadRecord("Select ItemId  from PromoGroup where GroupId = '" & ipdGroup.ItemId & "'")
    '                    Dim dPQty As Double = 0
    '                    While dtrTemp.Read
    '                        bCondition = False
    '                        iPriority = ipdGroup.Priority
    '                        If ipdGroup.Multiply = "Incremental" Then
    '                            bMultiply = True
    '                        End If
    '                        For iCnt As Integer = 0 To dgvOrdItem.Rows.Count - 1
    '                            If iCnt = dgvOrdItem.NewRowIndex Then Exit For
    '                            If dtrTemp("ItemId") = dgvOrdItem.Item(0, iCnt).Value.ToString And ipdGroup.UOM = dgvOrdItem.Item(4, iCnt).Value.ToString And CDbl(dgvOrdItem.Item(6, iCnt).Value.ToString) > 0 Then
    '                                If dgvOrdItem.Item(16, iCnt).Value.ToString <> "" Then
    '                                    If CDbl(dgvOrdItem.Item(16, iCnt).Value.ToString) >= ipdGroup.Priority Then
    '                                        dPQty += CDbl(dgvOrdItem.Item(5, iCnt).Value.ToString)
    '                                        arrCnt.Add(iCnt)
    '                                    End If
    '                                Else
    '                                    dPQty += CDbl(dgvOrdItem.Item(5, iCnt).Value.ToString)
    '                                    arrCnt.Add(iCnt)
    '                                End If
    '                            End If
    '                        Next
    '                    End While
    '                    dtrTemp.Close()
    '                    If ipdGroup.MinQty > dPQty Or ipdGroup.MaxQty < dPQty Then
    '                        bCondition = False
    '                        Exit For
    '                    End If
    '                    bCondition = True
    '                    If bMultiply = True Then
    '                        iMulti = System.Math.Floor(CDbl(dPQty) / ipdGroup.MinQty)
    '                        If iMulti < iMultiCnt Then iMultiCnt = iMulti
    '                    End If
    '                    If bCondition = False Then
    '                        Exit For
    '                    End If

    '                Next

    '                If bPromoGroup = True And bCondition = False Then GoTo NextRecord

    '                If bCondition = True Then
    '                    If stPromo.Entitle > 0 Then
    '                        If iPCount + iMultiCnt > stPromo.Entitle Then
    '                            iMultiCnt = stPromo.Entitle - iPCount
    '                        End If
    '                    End If
    '                    For i As Integer = 0 To arrCnt.Count - 1
    '                        If dgvOrdItem.Item(14, arrCnt(i)).Value.ToString <> "" Then
    '                            RemovePromotion(dgvOrdItem.Item(14, arrCnt(i)).Value.ToString)
    '                        End If
    '                        dgvOrdItem.Item(14, arrCnt(i)).Value = stPromo.PromoId
    '                        dgvOrdItem.Item(16, arrCnt(i)).Value = CStr(iPriority)
    '                    Next
    '                    dtrOffer = objDO.ReadRecordAnother("Select PromoOffer.ItemId, ItemName, PromoOffer.Uom, PromoOffer.FOcQty, PromoOffer.DisPrice, PromoOffer.Discount, Category from PromoOffer, Item  where PromoOffer.ItemId = Item.ItemNo and PromoID = '" & stPromo.PromoId & "'")
    '                    While dtrOffer.Read

    '                        If dtrOffer("FocQty") > 0 Then  '
    '                            Dim row As String() = New String() _
    '                            {dtrOffer("ItemiD").ToString, "", dtrOffer("ItemName").ToString, sLoc, "", _
    '                            CStr(iMultiCnt * dtrOffer("FocQty")), "0.00", "0.00", "", "", "1", "", "", "", "0", "0"}
    '                            dgvOrdItem.Rows.Add(row)

    '                            'Dim cb As New DataGridViewTextBoxCell
    '                            Dim cb As New DataGridViewComboBoxCell
    '                            cb = dgvOrdItem.Item(4, dgvOrdItem.RowCount - 2)
    '                            dgvOrdItem.Item(4, dgvOrdItem.RowCount - 2).Value = ""
    '                            cb.Items.Clear()
    '                            Dim bCur As Boolean = False
    '                            'MsgBox(dgvOrdItem.Item(0, e.RowIndex).Value)
    '                            dtr = objDO.ReadRecord("Select Uom from UOM where ItemNo=" & objDO.SafeSQL(dgvOrdItem.Item(0, dgvOrdItem.RowCount - 2).Value))
    '                            While dtr.Read
    '                                cb.Items.Add(dtr("Uom"))
    '                                bCur = True
    '                            End While
    '                            dtr.Close()
    '                            dgvOrdItem.Item(4, dgvOrdItem.RowCount - 2).Value = dtrOffer("UOM").ToString

    '                        ElseIf dtrOffer("DisPrice") > 0 Then
    '                            For iCnt As Integer = 0 To dgvOrdItem.Rows.Count - 1
    '                                If iCnt = dgvOrdItem.NewRowIndex Then Exit For
    '                                If dtrOffer("ItemId").ToString = dgvOrdItem.Item(0, iCnt).Value.ToString And dtrOffer("Uom").ToString = dgvOrdItem.Item(4, iCnt).Value.ToString Then
    '                                    dgvOrdItem.Item(6, iCnt).Value = dtrOffer("DisPrice").ToString
    '                                    dgvOrdItem.Item(5, iCnt).Value = Format(CDbl(dgvOrdItem.Item(5, iCnt).Value.ToString) * CDbl(dgvOrdItem.Item(6, iCnt).Value.ToString), "0.0000")
    '                                End If
    '                            Next
    '                        ElseIf dtrOffer("Discount") > 0 Then
    '                            For iCnt As Integer = 0 To dgvOrdItem.Rows.Count - 1
    '                                If iCnt = dgvOrdItem.NewRowIndex Then Exit For
    '                                If dtrOffer("ItemId").ToString = dgvOrdItem.Item(0, iCnt).Value.ToString And dtrOffer("Uom").ToString = dgvOrdItem.Item(4, iCnt).Value.ToString Then
    '                                    dgvOrdItem.Item(6, iCnt).Value = Format((1 - dtrOffer("Discount") / 100) * CDbl(dgvOrdItem.Item(9, iCnt).Value.ToString), "0.0000")
    '                                    dgvOrdItem.Item(5, iCnt).Value = Format(CDbl(dgvOrdItem.Item(5, iCnt).Value.ToString) * CDbl(dgvOrdItem.Item(6, iCnt).Value.ToString), "0.00")
    '                                End If
    '                            Next
    '                        End If
    '                        Dim stP As PromoCount
    '                        For iCnt As Integer = ArrPromoCount.Count - 1 To 0 Step -1
    '                            stP = ArrPromoCount(iCnt)
    '                            If stP.PromoId = stPromo.PromoId Then
    '                                ArrPromoCount.RemoveAt(iCnt)
    '                            End If
    '                        Next
    '                        stP.PromoId = stPromo.PromoId
    '                        stP.Count = iMultiCnt
    '                        ArrPromoCount.Add(stP)
    '                    End While
    '                    dtrOffer.Close()
    '                    Calculate()
    '                    Exit For
    '                End If
    'NextRecord:
    '            Next
    '        Catch ex As Exception
    '            MsgBox(ex.Message)
    '        End Try
    '    End Sub

    Private Function CheckInvoicePromotion(ByVal dAmt As Double) As Boolean
        Dim dtr As SqlDataReader
        Dim dtrOffer As SqlDataReader
        Dim Arr As New ArrayList
        Dim bApplied As Boolean = False
        dtr = objDO.ReadRecord("Select Promotion.PromoId, MinAmt, MaxAmt, CatBased, ItemCondition, DisPer, DisAmt, Multiply, Entitle, EntitleType from Promotion, PromoApply Where PromoName <> 'Special Price' and PromoType = 'Invoice Promotion' and Promotion.MinAmt <= " & dAmt & " and Promotion.MaxAmt >= " & dAmt & " and (Promotion.PromoId = PromoApply.PromoId) and FromDate <= '" & Format(Date.Now, "yyyyMMdd") & "' and ToDate >= '" & Format(Date.Now, "yyyyMMdd") & "' and ((Promotion.ApType = 'All') or (PromoApply.Id = '" & CustIdTextBox.Text & "' and Promotion.ApType = 'Customer') or (PromoApply.Id = '" & sZone & "' and Promotion.ApType = 'Zone') or (PromoApply.Id = '" & AgentIDTextBox.Text & "' and Promotion.ApType = 'Agent') or (PromoApply.Id = '" & sPrGroup & "' and Promotion.ApType = 'Price Group')) order by Priority")
        While dtr.Read
            Dim ipdDet As InvPromoDet
            ipdDet.PromoId = dtr("PromoId")
            ipdDet.MinAmt = dtr("MinAmt")
            ipdDet.MaxAmt = dtr("MaxAmt")
            ipdDet.CatBased = dtr("CatBased")
            ipdDet.ItemCondition = dtr("ItemCondition")
            ipdDet.DisPer = dtr("DisPer")
            ipdDet.DisAmt = dtr("DisAmt")
            ipdDet.Multiply = dtr("Multiply")
            ipdDet.Entitle = dtr("Entitle")
            ipdDet.EntitleType = dtr("EntitleType")
            Arr.Add(ipdDet)
        End While
        dtr.Close()
        If Arr.Count = 0 Then Exit Function

        Dim bMultiply As Boolean = False
        Dim bMultiStarted As Boolean = False
        Dim iMultiCnt As Integer = 1
        Dim iMulti As Integer = 1

        For iIndex As Integer = 0 To Arr.Count - 1

            Dim ipdDet As InvPromoDet
            ipdDet = Arr(iIndex)
            If ipdDet.Multiply = "Incremental" Then
                bMultiply = True
            Else
                bMultiply = False
            End If
            Dim iPCount As Integer = 0
            'Checking Items
            If ipdDet.Entitle > 0 Then
                If ipdDet.EntitleType = "Per Day" Then
                    dtr = objDO.ReadRecord("Select Sum(PromoCount) as TotalPromoApplied from PromoEntitlement where PromoId = " & objDO.SafeSQL(ipdDet.PromoId) & " and OrderDate = '" & Format(Date.Now, "yyyyMMdd") & "'")
                    If dtr.Read Then
                        If dtr("TotalPromoApplied").ToString <> "" Then
                            iPCount = dtr("TotalPromoApplied")
                        End If
                    End If
                    dtr.Close()
                ElseIf ipdDet.EntitleType = "Per Week" Then
                    dtr = objDO.ReadRecord("Select Sum(PromoCount) as TotalPromoApplied from PromoEntitlement where PromoId = " & objDO.SafeSQL(ipdDet.PromoId) & " and OrderDate >= '" & Format(DateAdd(DateInterval.Day, Date.Now.DayOfWeek * -1, Date.Now), "yyyyMMdd") & "'")
                    If dtr.Read Then
                        If dtr("TotalPromoApplied").ToString <> "" Then
                            iPCount = dtr("TotalPromoApplied")
                        End If
                    End If
                    dtr.Close()
                ElseIf ipdDet.EntitleType = "Per Month" Then
                    dtr = objDO.ReadRecord("Select Sum(PromoCount) as TotalPromoApplied from PromoEntitlement where PromoId = " & objDO.SafeSQL(ipdDet.PromoId) & " and OrderDate >= '" & Format(Date.Now, "yyyy") & Format(Date.Now, "MM") & "01" & "'")
                    If dtr.Read Then
                        If dtr("TotalPromoApplied").ToString <> "" Then
                            iPCount = dtr("TotalPromoApplied")
                        End If
                    End If
                    dtr.Close()
                ElseIf ipdDet.EntitleType = "Per Promotion" Then
                    dtr = objDO.ReadRecord("Select Sum(PromoCount) as TotalPromoApplied from PromoEntitlement where PromoId = " & objDO.SafeSQL(ipdDet.PromoId))
                    If dtr.Read Then
                        If dtr("TotalPromoApplied").ToString <> "" Then
                            iPCount = dtr("TotalPromoApplied")
                        End If
                    End If
                    dtr.Close()
                End If
                If iPCount >= ipdDet.Entitle Then GoTo NextRecord
            End If
            If ipdDet.CatBased = True Then
                Dim dTotAmt As Double = 0
                dtr = objDO.ReadRecord("Select PromoId, CategoryId from PromoCategory where PromoId = '" & ipdDet.PromoId & "'")
                While dtr.Read = True
                    For iCnt As Integer = 0 To dgvOrdItem.Rows.Count - 1
                        If dtr("CategoryId").ToString = GetCategory(dgvOrdItem.Item(0, iCnt).Value.ToString) Then
                            dTotAmt += CDbl(dgvOrdItem.Item(5, iCnt).Value.ToString)
                        End If
                    Next
                End While
                dtr.Close()
                If dTotAmt < ipdDet.MinAmt Or dTotAmt > ipdDet.MaxAmt Then
                    GoTo NextRecord
                End If
                If bMultiply = True Then
                    iMulti = System.Math.Floor(CDbl(dTotAmt) / ipdDet.MinAmt)
                    If bMultiStarted = False Then
                        iMultiCnt = iMulti
                        bMultiStarted = True
                    End If
                    If iMulti < iMultiCnt Then iMultiCnt = iMulti
                End If
            End If

            'Checking Items
            Dim bCondition As Boolean = False
            Dim bItemPromo As Boolean = False
            Dim bCatPromo As Boolean = False

            dtr = objDO.ReadRecord("Select Promotion.PromoId, Promotion.PromoName, Priority, PromoCondition.ItemId, ItemName, PromoCondition.UOM, MinQty, MaxQty  from PromoCondition, Promotion, Item where Promotion.PromoId = PromoCondition.PromoId and PromoCondition.ItemId = Item.ItemNo and PromoCondition.LineType = 'Item' and Promotion.PromoId = '" & ipdDet.PromoId & "'")
            While dtr.Read = True
                bItemPromo = True
                bCondition = False
                For iCnt As Integer = 0 To dgvOrdItem.Rows.Count - 1
                    If dtr("ItemId").ToString = dgvOrdItem.Item(0, iCnt).Value.ToString And dtr("Uom").ToString = dgvOrdItem.Item(4, iCnt).Value.ToString Then
                        If dtr("MinQty") > CDbl(dgvOrdItem.Item(5, iCnt).Value.ToString) Or dtr("MaxQty") < CDbl(dgvOrdItem.Item(5, iCnt).Value.ToString) Then
                            bCondition = False
                            Exit While
                        Else
                            If CDbl(dgvOrdItem.Item(6, iCnt).Value.ToString) = 0 Then
                                bCondition = False
                                Exit While
                            End If
                            If bMultiply = True Then
                                iMulti = System.Math.Floor(CDbl(dgvOrdItem.Item(5, iCnt).Value.ToString) / dtr("MinQty"))
                                If bMultiStarted = False Then
                                    iMultiCnt = iMulti
                                    bMultiStarted = True
                                End If
                                If iMulti < iMultiCnt Then iMultiCnt = iMulti
                            End If
                            bCondition = True
                        End If
                    End If
                Next
                If bCondition = False Then
                    Exit While
                End If
            End While
            dtr.Close()
            If bItemPromo = True And bCondition = False Then GoTo NextRecord

            ' Checking Category
            dtr = objDO.ReadRecord("Select Promotion.PromoId, Promotion.PromoName, Priority, PromoCondition.ItemId, PromoCondition.UOM, MinQty, MaxQty  from PromoCondition, Promotion where Promotion.PromoId = PromoCondition.PromoId and PromoCondition.LineType = 'Category' and Promotion.PromoId = '" & ipdDet.PromoId & "'")
            While dtr.Read = True
                bCatPromo = True
                bCondition = False
                Dim dPQty As Double = 0
                For iCnt As Integer = 0 To dgvOrdItem.Rows.Count - 1
                    If dtr("ItemId").ToString = GetCategory(dgvOrdItem.Item(0, iCnt).Value.ToString) And dtr("Uom").ToString = dgvOrdItem.Item(4, iCnt).Value.ToString And CDbl(dgvOrdItem.Item(6, iCnt).Value.ToString) > 0 Then
                        dPQty += CDbl(dgvOrdItem.Item(5, iCnt).Value.ToString)
                    End If
                Next
                If dtr("MinQty") > dPQty Or dtr("MaxQty") < dPQty Then
                    bCondition = False
                    Exit While
                End If
                If bMultiply = True Then
                    iMulti = System.Math.Floor(dPQty / dtr("MinQty"))
                    If bMultiStarted = False Then
                        iMultiCnt = iMulti
                        bMultiStarted = True
                    End If
                    If iMulti < iMultiCnt Then iMultiCnt = iMulti
                End If
                bCondition = True
            End While
            dtr.Close()
            If bCatPromo = True And bCondition = False Then GoTo NextRecord

            If ipdDet.Entitle > 0 Then
                If iPCount + iMultiCnt > ipdDet.Entitle Then
                    iMultiCnt = ipdDet.Entitle - iPCount
                End If
            End If

            dtrOffer = objDO.ReadRecord("Select PromoOffer.ItemId, ItemName, PromoOffer.Uom, PromoOffer.FOcQty, PromoOffer.DisPrice, PromoOffer.Discount, Category from PromoOffer, Item where PromoOffer.ItemId = Item.ItemNo and PromoID = '" & ipdDet.PromoId & "'")
            While dtrOffer.Read
                If dtrOffer("FocQty") > 0 Then
                    Dim row As String() = New String() _
                        {dtrOffer("ItemiD").ToString, "", dtrOffer("ItemName").ToString, "", "", _
                        CStr(iMultiCnt * dtrOffer("FocQty")), "0.00", "0.00", "", "", "1", "", "", "", "0", "0"}
                    dgvOrdItem.Rows.Add(row)
                End If
            End While
            dtrOffer.Close()

            Dim stP As PromoCount
            For iCnt As Integer = ArrPromoCount.Count - 1 To 0 Step -1
                stP = ArrPromoCount(iCnt)
                If stP.PromoId = ipdDet.PromoId Then
                    ArrPromoCount.RemoveAt(iCnt)
                End If
            Next
            stP.PromoId = ipdDet.PromoId
            stP.Count = iMultiCnt
            ArrPromoCount.Add(stP)
            CalculateDiscount(ipdDet.DisPer, ipdDet.DisAmt)

            bApplied = True
NextRecord:
        Next

        Return bApplied
    End Function

    Public Function SavePromoEntitlement(ByVal sOrdNo As String, ByVal dtOrdDt As Date, ByVal sCustNo As String, ByVal sPromoId As String, ByVal iPromoCount As Integer) As Boolean
        objDO.ExecuteSQL("Insert into PromoEntitlement (CustId, OrdNo, OrderDate, PromoId, PromoCount) values (" & objDO.SafeSQL(sCustNo) & "," & objDO.SafeSQL(sOrdNo) & ",'" & Format(dtOrdDt, "yyyyMMdd") & "'," & objDO.SafeSQL(sPromoId) & "," & iPromoCount & ")")
    End Function

    Public Function GetCategory(ByVal sItemNo As String) As String
        Dim dtr As SqlDataReader
        Dim sCat As String = ""
        dtr = objDO.ReadRecord("Select Category from Item where ItemNo = " & objDO.SafeSQL(sItemNo))
        If dtr.Read Then
            sCat = dtr("Category").ToString
        End If
        dtr.Close()
        Return sCat
    End Function

    Private Sub RemovePromotion(ByVal PromoID As String)
        For iCnt As Integer = 0 To dgvOrdItem.Rows.Count - 1
            If dgvOrdItem.Item(14, iCnt).Value.ToString = PromoID Then
                dgvOrdItem.Item(14, iCnt).Value = ""
                dgvOrdItem.Item(16, iCnt).Value = ""
            End If
            If dgvOrdItem.Item(14, iCnt).Value.ToString = PromoID Then
                If dgvOrdItem.Item(6, iCnt).Value.ToString = 0 Then
                    dgvOrdItem.Item(5, iCnt).Value = 0
                Else
                    dgvOrdItem.Item(6, iCnt).Value = dgvOrdItem.Item(9, iCnt).Value
                End If
            End If
        Next
    End Sub

    '=====================================================================================
    'Search Function
    Public Sub ReturnSearch(ByVal SQL As String) Implements SalesInterface.MobileSales.ISalesBase.ReturnSearch
        If SQL = "" Then
            GetOrders("select * from PO order by PONo")
        Else
            GetOrders("select * from PO Where " & SQL & " order by PONo")
        End If
    End Sub

    Public Event SearchField(ByVal FormName As String, ByVal FieldName As String, ByVal FieldType As String, ByVal CurValue As String) Implements SalesInterface.MobileSales.ISalesBase.SearchField

    Private Sub GetTextBox()
        Dim ctl As Object
        Dim tb As TextBox
        For Each ctl In Me.Controls
            If (ctl.GetType() Is GetType(TabControl)) Then
                GetTabControl(ctl)
            End If
            If (ctl.GetType() Is GetType(Panel)) Then
                GetPanel(ctl)
            End If
            If (ctl.GetType() Is GetType(TextBox)) Then
                tb = ctl
                AddHandler tb.Enter, AddressOf TextBox_Enter
            End If
        Next
    End Sub

    Private Sub GetTabControl(ByVal obj As Object)
        Dim ctl As Object
        Dim tb As TextBox
        For Each ctl In obj.Controls
            If (ctl.GetType() Is GetType(Panel)) Then
                GetPanel(ctl)
            End If
            If (ctl.GetType() Is GetType(TabPage)) Then
                GetTabPage(ctl)
            End If
            If (ctl.GetType() Is GetType(TextBox)) Then
                tb = ctl
                AddHandler tb.Enter, AddressOf TextBox_Enter
            End If
        Next
    End Sub

    Private Sub GetTabPage(ByVal obj As Object)
        Dim ctl As Object
        Dim tb As TextBox
        For Each ctl In obj.Controls
            If (ctl.GetType() Is GetType(Panel)) Then
                GetPanel(ctl)
            End If
            If (ctl.GetType() Is GetType(TextBox)) Then
                tb = ctl
                AddHandler tb.Enter, AddressOf TextBox_Enter
            End If
        Next
    End Sub

    Private Sub GetPanel(ByVal obj As Object)
        Dim ctl As Object
        Dim tb As TextBox
        For Each ctl In obj.Controls
            If (ctl.GetType() Is GetType(TextBox)) Then
                tb = ctl
                AddHandler tb.Enter, AddressOf TextBox_Enter
            End If
        Next
    End Sub

    Private Sub TextBox_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim tb As TextBox = sender
        Try
            Dim sField As String
            Dim sFieldType As String
            sField = tb.DataBindings.Item(0).BindingMemberInfo.BindingField.ToString
            sFieldType = tb.DataBindings.Item(0).BindingMemberInfo.BindingField.GetType.Name
            If sField <> "" Then RaiseEvent SearchField(Me.ProductName & "." & Me.Name, sField, sFieldType, tb.Text)
        Catch ex As Exception

        End Try

    End Sub

    'End Search
    '=====================================================================================

    Private Sub InvNoTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles InvNoTextBox.TextChanged
        Dim rs As SqlDataReader
        Dim iIndex As Integer = 0
        Dim sPrCode As String = ""
        rs = objDO.ReadRecord("Select * from POItem where PoNo = '" & InvNoTextBox.Text & "'")
        dgvOrdItem.Rows.Clear()
        Dim cnt As Integer = 0
        bView = True
        While rs.Read
            Dim row As String() = New String() _
            {rs("ItemNo").ToString, "", rs("Description").ToString, "", "", _
            rs("Qty").ToString, rs("Price").ToString, rs("DisPr").ToString, rs("DisPer").ToString, rs("SubAmt").ToString, "", "", rs("Foc").ToString, rs("DeliveryDate").ToString, rs("PromoId").ToString, rs("PromoOffer").ToString, rs("Priority").ToString, rs("DeliQty").ToString, "0", rs("Discount").ToString}
            dgvOrdItem.Rows.Add(row)

            Dim cb2 As New DataGridViewTextBoxCell
            dgvOrdItem.Item(1, cnt) = cb2
            dgvOrdItem.Item(1, cnt).Value = rs("VariantCode")

            Dim cb As New DataGridViewTextBoxCell
            dgvOrdItem.Item(4, cnt) = cb
            dgvOrdItem.Item(4, cnt).Value = rs("UOM")

            Dim cb1 As New DataGridViewTextBoxCell
            dgvOrdItem.Item(3, cnt) = cb1
            dgvOrdItem.Item(3, cnt).Value = rs("Location")

            cnt = cnt + 1
            dgvOrdItem.Item(4, iIndex).Value = rs("Uom").ToString
            dgvOrdItem.Item(3, iIndex).Value = rs("Location").ToString
            iIndex = iIndex + 1
        End While
        bView = False
        rs.Close()
        cnt = 0
    End Sub

    'Private Sub btnInvno_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInvno.Click
    '    RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\Invoice.dll", "Invoice.OrdList", "Invoice.SalesInvoice", "OrdNoTextBox", 0, 0)
    'End Sub

    'Localization
    Private Sub Localization()
        Try
            'cInfo = New CultureInfo(sLang)
            'Thread.CurrentThread.CurrentCulture = cInfo
            'Thread.CurrentThread.CurrentUICulture = cInfo
            'rMgr = New ResourceManager("Invoice.Invoice", [Assembly].GetExecutingAssembly())
            'Me.Text = rMgr.GetString("SalesInvoice")
            'dgvOrdItem.Columns("ItemNo").HeaderText = rMgr.GetString("Col_DataForm_ItemNo")
            'dgvOrdItem.Columns("VariantName").HeaderText = rMgr.GetString("Col_DataForm_VariantName")
            'dgvOrdItem.Columns("ItemName").HeaderText = rMgr.GetString("Col_DataForm_ItemName")
            'dgvOrdItem.Columns("WareHouse").HeaderText = rMgr.GetString("Col_DataForm_Warehouse")
            'dgvOrdItem.Columns("UOM").HeaderText = rMgr.GetString("Col_DataForm_Uom")
            'dgvOrdItem.Columns("Quantity").HeaderText = rMgr.GetString("Col_DataForm_Qty")
            'dgvOrdItem.Columns("Price").HeaderText = rMgr.GetString("Col_DataForm_Price")
            'dgvOrdItem.Columns("Amount").HeaderText = rMgr.GetString("Col_DataForm_Amount")
            'dgvOrdItem.Columns("BaseUOM").HeaderText = rMgr.GetString("Col_DataForm_BaseUOM")
            'dgvOrdItem.Columns("BasePrice").HeaderText = rMgr.GetString("Col_DataForm_BaseUOM")
            'dgvOrdItem.Columns("FOC").HeaderText = rMgr.GetString("Col_DataForm_FOc")
            'dgvOrdItem.Columns("DeliveredDate").HeaderText = rMgr.GetString("Col_DataForm_DelDate")
            'dgvOrdItem.Columns("Promotion").HeaderText = rMgr.GetString("Col_DataForm_Promotion")
            'dgvOrdItem.Columns("Offer").HeaderText = rMgr.GetString("Col_DataForm_Offer")
            'dgvOrdItem.Columns("Priority").HeaderText = rMgr.GetString("Col_DataForm_Prioty")
            'dgvOrdItem.Columns("QtyShip").HeaderText = rMgr.GetString("Col_DataForm_QtyShip")
            'dgvOrdItem.Columns("Explode").HeaderText = rMgr.GetString("Col_DataForm_Explode")
            'InvNoLabel.Text = rMgr.GetString("InvNoLabel")
            'OrdDtLabel.Text = rMgr.GetString("InvDtLabel")
            ''  OrdNoLabel1.Text = rMgr.GetString("OrdNoLabel")
            'CustIdLabel1.Text = rMgr.GetString("CustIdLabel1")
            'lblCustName.Text = rMgr.GetString("CustNameLabel")
            'RefNoLabel.Text = rMgr.GetString("lblCustRef")
            'PayTermsLabel1.Text = rMgr.GetString("PayTermsLabel1")
            'AgentIDLabel1.Text = rMgr.GetString("AgentIDLabel1")
            'lblVoid.Text = rMgr.GetString("lblVoid")
            'DiscountLabel1.Text = rMgr.GetString("DiscountLabel1")
            'SubTotalLabel1.Text = rMgr.GetString("SubTotalLabel1")
            'GSTAmtLabel1.Text = rMgr.GetString("GSTAmtLabel1")
            'TotalAmtLabel1.Text = rMgr.GetString("TotalAmtLabel1")
            'CurCodeLabel1.Text = rMgr.GetString("CurCodeLabel1")
            'CurExRateLabel1.Text = rMgr.GetString("CurExRateLabel1")

            'Sales_lblAvailable.Text = rMgr.GetString("Sales_lblAvailable")
            'Sales_lblPOqty.Text = rMgr.GetString("Sales_lblPOqty")
            'Sales_lblReserved.Text = rMgr.GetString("Sales_lblReserved")
            'Sales_lblWareHouse.Text = rMgr.GetString("Sales_lblWareHouse")
            'Sales_SOQty.Text = rMgr.GetString("Sales_SOQty")

            'TabControl2.TabPages(0).Text = rMgr.GetString("TabCol_SalInv_Order")
            'TabControl2.TabPages(1).Text = rMgr.GetString("TabCol_DataForm_Foriegn")

            'gbAmount.Text = rMgr.GetString("gbAmount")
            ''gbItemCard.Text = rMgr.GetString("gbItemCard")

            'btnCancel.Text = rMgr.GetString("btnCancel")
            'btnSave.Text = rMgr.GetString("btnSave")
            'btnNew.Text = rMgr.GetString("btnNew")
            'btnVoid.Text = rMgr.GetString("btnVoid")
            'btnModify.Text = rMgr.GetString("btnModify")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub SetCulture(ByVal CultureName As String) Implements SalesInterface.MobileSales.ISalesBase.SetCulture
        'sLang = CultureName
        'Localization()
    End Sub
    Public Sub LoadReport(ByVal InvNo As String)
        Try
            objDO.ExecuteSQL("Delete from InvReport")
            Dim strSql As String = "SELECT Item.Description, POItem.[LineNo] as Line, PoItem.ItemNo, Customer.CustNo, Customer.CustName, Customer.Address, Customer.Address2, Customer.Address3, Customer.Address4, Customer.Phone, Customer.ContactPerson, Customer.FaxNo, PO.PONo as DoNo, PO.PONo AS Expr1, PO.PoDt, PO.PoNo, PO.AgentId, PO.CustRefNo, PO.PayTerms, POItem.Qty, POItem.Price, POItem.SubAmt, PO.SubTotal, PO.GST, PO.GstAmt, PO.TotalAmt, Item.AssemblyBOM FROM Customer INNER JOIN PO ON Customer.CustNo = PO.CustId INNER JOIN POItem ON PO.PONo = POItem.PONo INNER JOIN Item ON POItem.ItemNo = Item.ItemNo and PO.PONo=" & objDO.SafeSQL(InvNo)
            Dim rs As SqlDataReader
            rs = objDO.ReadRecord(strSql)

            While rs.Read = True
                objDO.ExecuteSQLAnother("Insert into InvReport(InvNo, InvDate, CustNo, AgentID, CustName, Add1, Add2, Add3, Add4, Phone, Fax, Attn, DoNo, RefNo, PONo, Terms, Line, ItemNo, Description, Qty, price, SubAmt, SubTotal, GSTAmt, TotalAmt) values (" & objDO.SafeSQL(rs("Expr1").ToString) & "," & objDO.SafeSQL(Format(CDate(rs("InvDt")), "dd-MMM-yyyy")) & "," & objDO.SafeSQL(rs("CustNo").ToString) & "," & objDO.SafeSQL(rs("AgentID").ToString) & "," & objDO.SafeSQL(rs("CustName").ToString) & "," & objDO.SafeSQL(rs("Address").ToString) & "," & objDO.SafeSQL(rs("Address2").ToString) & "," & objDO.SafeSQL(rs("Address3").ToString) & "," & objDO.SafeSQL(rs("Address4").ToString) & "," & objDO.SafeSQL(rs("Phone").ToString) & "," & objDO.SafeSQL(rs("FaxNo").ToString) & "," & objDO.SafeSQL(rs("ContactPerson").ToString) & "," & objDO.SafeSQL(rs("DONo").ToString) & "," & objDO.SafeSQL(rs("CustRefNo").ToString) & "," & objDO.SafeSQL(rs("PONo").ToString) & "," & objDO.SafeSQL(rs("PayTerms").ToString) & "," & rs("Line").ToString & "," & objDO.SafeSQL(rs("ItemNo").ToString) & "," & objDO.SafeSQL(rs("Description").ToString) & "," & objDO.SafeSQL(rs("Qty").ToString) & "," & objDO.SafeSQL(Format(rs("Price"), "0.00")) & "," & objDO.SafeSQL(Format(rs("SubAmt"), "0.00")) & "," & objDO.SafeSQL(Format(rs("SubTotal"), "0.00")) & "," & objDO.SafeSQL(Format(rs("GstAmt"), "0.00")) & "," & objDO.SafeSQL(Format(rs("TotalAmt"), "0.00")) & ")")
                If CBool(rs("AssemblyBOM")) = True Then
                    Dim rsAss As SqlDataReader
                    Dim sSpace As String = "  "
                    rsAss = objDO.ReadRecordAnother("Select * from AssemblyBOM where ParentItemNo = " & objDO.SafeSQL(rs("ItemNo")))
                    While rsAss.Read
                        objDO.ExecuteSQLAnother("Insert into InvReport(InvNo, InvDate, CustNo, AgentID, CustName, Add1, Add2, Add3, Add4, Phone, Fax, Attn, DoNo, RefNo, PONo, Terms, ItemNo, Description, Qty, price, SubAmt, SubTotal, GSTAmt, TotalAmt) values (" & objDO.SafeSQL(rs("Expr1").ToString) & "," & objDO.SafeSQL(Format(CDate(rs("InvDt")), "dd-MMM-yyyy")) & "," & objDO.SafeSQL(rs("CustNo").ToString) & "," & objDO.SafeSQL(rs("AgentID").ToString) & "," & objDO.SafeSQL(rs("CustName").ToString) & "," & objDO.SafeSQL(rs("Address").ToString) & "," & objDO.SafeSQL(rs("Address2").ToString) & "," & objDO.SafeSQL(rs("Address3").ToString) & "," & objDO.SafeSQL(rs("Address4").ToString) & "," & objDO.SafeSQL(rs("Phone").ToString) & "," & objDO.SafeSQL(rs("FaxNo").ToString) & "," & objDO.SafeSQL(rs("ContactPerson").ToString) & "," & objDO.SafeSQL(rs("DONo").ToString) & "," & objDO.SafeSQL(rs("CustRefNo").ToString) & "," & objDO.SafeSQL(rs("PONo").ToString) & "," & objDO.SafeSQL(rs("PayTerms").ToString) & "," & objDO.SafeSQL("") & "," & objDO.SafeSQL(sSpace & rsAss("Qty").ToString & " x " & rsAss("Description").ToString) & ", '', '', '', '', '', '')")
                    End While
                    rsAss.Close()
                End If
            End While
            rs.Close()
            'ExecuteReport("Select * from InvReport", "EverHomeInvoiceRep")
            strSql = "Select * from InvReport"
            ' Dim strSql As String = "SELECT Invoice.InvNo, Invoice.InvDt, Customer.CustNo, Customer.CustName, Invoice.PayTerms, Invoice.AgentId, InvItem.ItemNo, Item.Description, InvItem.UOM, InvItem.Qty, InvItem.Price, InvItem.SubAmt, InvItem.GstAmt, Invoice.TotalAmt FROM  Item INNER JOIN InvItem ON Item.ItemNo = InvItem.ItemNo INNER JOIN Invoice ON InvItem.InvNo = Invoice.InvNo INNER JOIN Customer ON Invoice.CustId = Customer.CustNo where Invoice.InvDt Between " & objDO.SafeSql(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & objDO.SafeSql(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strTerms & strPay
            'ExecuteReport(strSql, "DailySalesRep")
            Dim RptName As String = "EverHomeInvoiceRep"
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
            Dim frm As New ViewReport
            frm.crvReport.ShowRefreshButton = False
            frm.crvReport.ShowCloseButton = False
            frm.crvReport.ShowGroupTreeButton = False
            frm.crvReport.ReportSource = rptDocument
            frm.Show()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        If IsNewRecord = True Then Exit Sub
        If InvNoTextBox.Text = String.Empty Then
            MessageBox.Show("Select Invoice No to Print", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            InvNoTextBox.Select()
            Return
        Else
            LoadReport(InvNoTextBox.Text)
        End If
    End Sub

End Class