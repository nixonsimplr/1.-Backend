Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.odbc
Imports DataInterface
Imports SalesInterface.MobileSales
Public Class ConsignmentOrder
    Implements ISalesBase
    Dim bCmbCustLoading As Boolean = False
    Private aAgent As New ArrayList()
    Private objDO As New DataInterface.IbizDO
    Private arrOrdNo As New ArrayList()
    Dim strPay As String = " "
    Dim ArList As New ArrayList
    Dim sPrGroup, sAgent, sPayTerms As String
    Dim sMDT As String = "", sLoc As String = ""
    Dim dGstPercent As Double = 0

    Private Structure ArrOrdItem
        Dim ItemNo As String
        Dim Qty As Double
        Dim UOM As String
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

    Private Sub ConsignmentOrder_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        objDO.DisconnectDB()
    End Sub

    Private Sub ConsignmentOrder_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objDO.ConnectDB()
        loadCombo()
        chkAllCheque.Checked = True
        LoadDataGrid()
    End Sub
    Public Sub loadCombo()
        bCmbCustLoading = True
        Dim dtr As SqlDataReader
        '  dtr = objDO.ReadRecord("Select MDT.MDTNo as Code, MDT.Description from MDT order by Description")
        dtr = objDO.ReadRecord("Select CustNo, CustName from Customer Where Consignment=1 order by CustNo")
        cmbAgent.DataSource = Nothing
        aAgent.Clear()
        'aAgent.Add(New ComboValues("ALL", "ALL"))
        While dtr.Read()
            aAgent.Add(New ComboValues(dtr("CustNo").ToString, dtr("CustName").ToString))

            '    iSelIndex = iIndex
            'End If
            'iIndex = iIndex + 1
        End While
        dtr.Close()
        cmbAgent.DataSource = aAgent
        cmbAgent.DisplayMember = "Desc"
        cmbAgent.ValueMember = "Code"
        dtpFromDate.Value = Date.Now
        dtpToDate.Value = Date.Now
        bCmbCustLoading = False
    End Sub
    Private Sub LoadDataGrid()
        Dim rs As SqlDataReader
        If chkAllCheque.Checked = True Then
            rs = objDO.ReadRecord("Select OrdNo, OrdDT, AgentID, TotalAmt as Amount from OrderHdr, Customer where Customer.CustNo=OrderHdr.CustID and (ExportMYOB =0 or ExportMYOB is Null) and (Void =0 or Void is Null) and OrderHdr.Exported=1 " & strPay & " Order by OrdNo")
        Else
            rs = objDO.ReadRecord("Select OrdNo, OrdDT, AgentID, TotalAmt as Amount from OrderHdr, Customer where Customer.CustNo=OrderHdr.CustID and (ExportMYOB =0 or ExportMYOB is Null) and (Void =0 or Void is Null) and OrderHdr.Exported=1 and OrdDt between " & objDO.SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00.000")) & " and " & objDO.SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59.000")) & strPay & " Order by OrdDT")
        End If
        dgvPDC.Rows.Clear()
        While rs.Read
            Dim row As String() = New String() _
                           {False, rs("OrdNo").ToString, Format(rs("OrdDT"), "dd/MMM/yyyy"), rs("AgentID").ToString, Format(rs("Amount"), "0.00")}
            dgvPDC.Rows.Add(row)
        End While
        rs.Close()
        rs.Dispose()
    End Sub
    Private Sub cmbAgent_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbAgent.SelectedIndexChanged
        If bCmbCustLoading = True Then Exit Sub
        If cmbAgent.Text = "ALL" Then
            strPay = " "
        Else
            strPay = "and OrderHdr.CustID = " & "'" & cmbAgent.SelectedValue & "'"
        End If
        LoadDataGrid()
    End Sub

    Private Sub dtpToDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpToDate.ValueChanged
        If cmbAgent.Text = "ALL" Then
            strPay = ""
        Else
            strPay = "and OrderHdr.CustID  = " & "'" & cmbAgent.SelectedValue & "'"
        End If
        LoadDataGrid()
    End Sub

    Private Sub dtpFromDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpFromDate.ValueChanged
        If cmbAgent.Text = "ALL" Then
            strPay = " "
        Else
            strPay = "and OrderHdr.CustID  = " & "'" & cmbAgent.SelectedValue & "'"
        End If
        LoadDataGrid()
    End Sub

    Private Sub chkSelAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSelAll.CheckedChanged
        Dim i As Integer
        If chkSelAll.Checked = True Then
            For i = 0 To dgvPDC.Rows.Count - 2
                dgvPDC.Item(0, i).Value = True
            Next
        Else
            For i = 0 To dgvPDC.Rows.Count - 2
                dgvPDC.Item(0, i).Value = False
            Next
        End If
    End Sub
    Private Sub chkAllCheque_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAllCheque.CheckedChanged
        If chkAllCheque.Checked = True Then
            dtpFromDate.Enabled = False
            dtpToDate.Enabled = False
            LoadDataGrid()
        Else
            dtpFromDate.Enabled = True
            dtpToDate.Enabled = True
            LoadDataGrid()
        End If
    End Sub

    Private Sub btnCrOrder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCrOrder.Click
        Dim strSql As String = ""
        arrOrdNo = New ArrayList
        If dgvPDC.Rows.Count = 0 Then Exit Sub
        If dgvPDC.Rows.Count = 1 And dgvPDC.Item(1, 1).Value = "" Then
            Exit Sub
        End If
        Dim i As Integer
        Dim iCnt As Integer = 0
        For i = 0 To dgvPDC.Rows.Count - 2
            If Not dgvPDC.Item(0, i).Value Is Nothing Then
                If Not dgvPDC.Item(1, i).Value Is Nothing Then
                    If dgvPDC.Item(1, i).Value <> "" Then
                        If dgvPDC.Item(0, i).Value = True Then
                            iCnt += 1
                        End If
                    End If
                End If
            End If
        Next
        If iCnt = 0 Then
            MsgBox("No rows selected")
            Exit Sub
        End If
        If MsgBox("You are about to Create an Order for : " & cmbAgent.Text & vbCrLf & "Do you want to continue?", MsgBoxStyle.YesNo, "Create Order?") = MsgBoxResult.Yes Then
            For i = 0 To dgvPDC.Rows.Count - 2
                If Not dgvPDC.Item(0, i).Value Is Nothing Then
                    If Not dgvPDC.Item(1, i).Value Is Nothing Then
                        If dgvPDC.Item(1, i).Value <> "" Then
                            If dgvPDC.Item(0, i).Value = True Then
                                objDO.ExecuteSQL("Update OrderHdr Set ExportMYOB = 1 where OrdNo=" & objDO.SafeSQL(dgvPDC.Item(1, i).Value.ToString))
                                arrOrdNo.Add(dgvPDC.Item(1, i).Value.ToString)
                            End If
                        End If
                    End If
                End If
            Next
            Dim dtr As SqlDataReader
            Dim ordStr As String = "'"
            Dim cnt As Integer = arrOrdNo.Count - 1
            For i = 0 To arrOrdNo.Count - 1
                If i = cnt Then
                    ordStr = ordStr & arrOrdNo(i).ToString & "'"
                Else
                    ordStr = ordStr & arrOrdNo(i).ToString & "','"
                End If
            Next
            Dim sOrdNo As String = ""
            Dim rs As SqlDataReader
                rs = objDO.ReadRecord("Select MDTNo, GST from System")
            If rs.Read = True Then
                sMDT = rs("MDTNo").ToString
                dGSTPercent = rs("GST")
            End If
            rs.Close()
            rs = objDO.ReadRecord("Select * from MDT where MdtNo = " & objDO.SafeSQL(sMDT) & "")
            If rs.Read = True Then
                Dim sPre As String = rs("PreOrdNo")
                Dim iLen As Int32 = rs("LenOrdNo")
                Dim iOrdNo As Int32 = rs("LastOrdNo") + 1
                sOrdNo = sPre & StrDup(iLen - Len(sPre) - Len(CStr(iOrdNo)), "0") & CStr(iOrdNo)
                sLoc = rs("Location").ToString
            End If
            rs.Close()
            'dDiscount = "0.00"
            'dSubTotal = "0.00"
            'dGSTAmt = "0.00"
            'dTotalAmt = "0.00"
            dtr = objDO.ReadRecord("SELECT PriceGroup, SalesAgent, PaymentTerms from Customer where CustNo =  " & objDO.SafeSQL(cmbAgent.SelectedValue))
            While dtr.Read
                '  If IsRecExists(dtr("ItemNo").ToString, "") Then
                sPrGroup = dtr("PriceGroup").ToString
                sAgent = dtr("SalesAgent")
                sPayTerms = dtr("PaymentTerms").ToString
            End While
            dtr.Close()

            dtr = objDO.ReadRecord("SELECT ItemNo, UOM, Sum(Qty) from OrdItem where OrdNo in  " & objDO.SafeSQL(arrOrdNo(i).ToString) & " group by ItemNo, UOM")
            While dtr.Read
                '  If IsRecExists(dtr("ItemNo").ToString, "") Then
                Dim aPr1 As ArrOrdItem
                aPr1.ItemNo = dtr("ItemNo").ToString
                aPr1.Qty = dtr("Qty")
                aPr1.UOM = dtr("UOM").ToString
                ArList.Add(aPr1)
                'End If
                'objDO.ExecuteSQLAnother("Update Invoice Set PaidAmt = PaidAmt + " & dtr("AmtPaid") & " Where InvNo=" & objDO.SafeSQL(dtr("InvNo").ToString))
            End While
            dtr.Close()
            '''''''To Gaurav Insert OrdItem Code
            For i = 0 To ArList.Count - 1

                Dim aPr As ArrOrdItem
                aPr = ArList(i)
                '    Dim dPrice As Double = GetPrice()
                'Dim dSubAmt As Double = aPr.Qty * dPrice
                'objDO.ExecuteSQL("Insert into OrdItem(OrdNo, ItemNo, UOM, Qty, Price, SubAmt, Remarks, Description, DisPer, DisPr, PromoID, PromoOffer, ActPrice, ReasonCode, ExpiryDate, LineNum, [LineNo]) values (" & objDO.SafeSQL(sOrdNo) & "," & objDO.SafeSQL(aPr.ItemNo) & "," & objDO.SafeSQL(aPr.UOM) & "," & CStr(aPr.Qty) & "," & CStr(dPrice) & "," & CStr(dSubAmt) & "," & objDO.SafeSQL("") & "," & objDO.SafeSQL("") & "," & CStr(0) & "," & CStr(0) & "," & objDO.SafeSQL("") & "," & objDO.SafeSQL("") & "," & CStr(0) & "," & objDO.SafeSQL("") & "," & objDO.SafeSQL(Format(Date.Now, "yyyyMMdd HH:mm:ss")) & "," & i + 1 & ")")
            Next

            '''''''To Gaurav To do Order Header Code


            LoadDataGrid()
        End If
    End Sub
    Public Function GetPrice(ByVal sItemId As String, ByVal sUOM As String, ByVal sCustNo As String, ByVal sPrGroup As String) As Double
        objDO.ConnectAnotherDB()
        Dim dtr As SqlDataReader
        Dim sDefPrGroup As String = ""
        dtr = objDO.ReadRecordAnother("Select TradeDealPriceGroup from System")
        If dtr.Read Then
            sDefPrGroup = dtr("TradeDealPriceGroup")
        End If
        dtr.Close()

        Dim dPr As Double = 0
        'dtr = objDO.ReadRecordAnother("Select UnitPrice from ItemPr where ItemNo = '" & sItemId & "' and UOM = '& sUOM & "' and PriceGroup = '" & sCustNo & "'and SalesType = 'Customer' Order by MinQty")
        'If dtr.Read Then
        '    dPr = dtr("UnitPrice")
        'End If
        'dtr.Close()
        If dPr > 0 Then GoTo ExitNow
        dtr = objDO.ReadRecordAnother("Select UnitPrice from ItemPr where ItemNo = '" & sItemId & "' and PriceGroup = '" & sPrGroup & "' and SalesType = 'Customer Price Group' Order by MinQty")
        If dtr.Read Then
            dPr = dtr("UnitPrice")
        End If
        dtr.Close()
        If dPr > 0 Then GoTo ExitNow
        dtr = objDO.ReadRecordAnother("Select UnitPrice from ItemPr where ItemNo = '" & sItemId & "' and PriceGroup = '" & sDefPrGroup & "' and SalesType = 'Customer Price Group' Order by MinQty")
        If dtr.Read Then
            dPr = dtr("UnitPrice")
        End If
        dtr.Close()
        If dPr > 0 Then GoTo ExitNow
        dtr = objDO.ReadRecordAnother("Select UnitPrice from ItemPr where ItemNo = '" & sItemId & "' and SalesType = 'All Customers' Order by MinQty")
        If dtr.Read Then
            dPr = dtr("UnitPrice")
        End If
        dtr.Close()
        If dPr > 0 Then GoTo ExitNow
        dtr = objDO.ReadRecordAnother("Select UnitPrice from Item where ItemNo = '" & sItemId & "'")
        If dtr.Read Then
            dPr = dtr("UnitPrice")
        End If
        dtr.Close()
ExitNow:
        objDO.DisconnectAnotherDB()
        Return dPr
    End Function
    Private Function IsRecExists(ByVal sItemNo As String, ByVal sAgent As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean = False
        dtr = objDO.ReadRecord("Select * from SalesTargetQty where ItemNo = " & objDO.SafeSQL(sItemNo) & " and AgentID = " & objDO.SafeSQL(sAgent))
        If dtr.Read Then
            bAns = True
        End If
        dtr.Close()
        Return bAns
    End Function
End Class