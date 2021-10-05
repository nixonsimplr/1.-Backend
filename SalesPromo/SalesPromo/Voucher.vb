Imports System
Imports System.Resources
Imports System.Threading
Imports System.Reflection
Imports System.Globalization
Imports DataInterface.IbizDO
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Public Class Voucher
    Implements ISalesBase

    Private myAdapter As SqlDataAdapter
    Private myDataSet As DataSet
    Private myDataView As DataView
    Private RowNo As Integer
    Private objDO As New DataInterface.IbizDO
    Private IsNewRecord As Boolean = False
    Private IsModify As Boolean = True
    Dim iVouNo As Int32
    Dim rs As SqlDataReader
    Private rInd As Integer
    Private sMDT As String = ""
    Dim sVouType As String
    Dim sAmtType As String
    Private sPrGroup, sItemNo, sCustNo As String
    Private Sub OrdNoLabel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        IsNewRecord = True
        Dim sVouNo As String = ""
        '    dgvExchange.ReadOnly = False
        'btnNew.Enabled = False
        'btnSave.Enabled = True
        clear()
        rs = objDO.ReadRecord("Select MDTNo from System")
        If rs.Read = True Then
            sMDT = rs("MDTNo").ToString
        End If
        rs.Close()
        rs = objDO.ReadRecord("Select * from MDT where MdtNo = " & objDO.SafeSQL(sMDT) & "")
        If rs.Read = True Then
            Dim sPre As String = rs("PreVouNo")
            Dim iLen As Int32 = rs("LenVouNo")
            iVouNo = rs("LastVouNo") + 1
            sVouNo = sPre & StrDup(iLen - Len(sPre) - Len(CStr(iVouNo)), "0") & CStr(iVouNo)
        End If
        rs.Close()
        txtVouNo.Text = sVouNo
    End Sub
    Private Sub clear()
        Try
            txtVouNo.Text = String.Empty
            txtFromVouNo.Text = String.Empty
            txtToVouNo.Text = String.Empty
            txtVouPrefix.Text = String.Empty
            txtMinAmt.Text = String.Empty
            txtAmount.Text = String.Empty
            dtpExpDate.Value = Date.Now
            dtpVoucherDate.Value = Date.Now
            rbAmount.Checked = True
            cbGift.Checked = True
            chkMulUse.Checked = False
            chkActive.Checked = False
            rbAmount.Checked = True
            txtVouPrefix.Focus()
            dgvCustomerPrice.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Function GetListViewForm() As String Implements SalesInterface.MobileSales.ISalesBase.GetListViewForm
        RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\SalesPromo.dll", "SalesPromo.VoucherList", "SalesPromo.Voucher", "txtVoucherNo", 0, 0)
        Return Windows.Forms.Application.StartupPath & "\SalesPromo.dll,SalesPromo.VoucherList"
    End Function

    Public Sub ListViewClick() Implements SalesInterface.MobileSales.ISalesBase.ListViewClick

    End Sub


    Public Sub MoveFirstClick() Implements SalesInterface.MobileSales.ISalesBase.MoveFirstClick
        If IsNewRecord = True Then
            If MsgBox("Record Not Saved", MsgBoxStyle.YesNo, "Information") = MsgBoxResult.No Then
                Exit Sub
            End If
        End If
        RowNo = 0
        LoadRow()
    End Sub

    Public Sub MoveLastClick() Implements SalesInterface.MobileSales.ISalesBase.MoveLastClick
        If IsNewRecord = True Then
            If MsgBox("Record Not Saved", MsgBoxStyle.YesNo, "Information") = MsgBoxResult.No Then
                Exit Sub
            End If
        End If
        RowNo = myDataView.Count - 1
        LoadRow()
    End Sub

    Public Sub MoveNextClick() Implements SalesInterface.MobileSales.ISalesBase.MoveNextClick
        If IsNewRecord = True Then
            If MsgBox("Record Not Saved", MsgBoxStyle.YesNo, "Information") = MsgBoxResult.No Then
                Exit Sub
            End If
        End If
        If RowNo >= myDataView.Count - 1 Then Exit Sub
        RowNo = RowNo + 1
        LoadRow()
    End Sub

    Public Sub MovePositionClick(ByVal Position As Long) Implements SalesInterface.MobileSales.ISalesBase.MovePositionClick

    End Sub

    Public Sub MovePreviousClick() Implements SalesInterface.MobileSales.ISalesBase.MovePreviousClick
        If IsNewRecord = True Then
            If MsgBox("Record Not Saved", MsgBoxStyle.YesNo, "Information") = MsgBoxResult.No Then
                Exit Sub
            End If
        End If
        If RowNo <= 0 Then Exit Sub
        RowNo = RowNo - 1
        LoadRow()
    End Sub

    Public Sub Print() Implements SalesInterface.MobileSales.ISalesBase.Print

    End Sub

    Public Sub Print(ByVal PageNo As Integer) Implements SalesInterface.MobileSales.ISalesBase.Print

    End Sub

    Public Event ResultData(ByVal ChildLoadType As String, ByVal Value As String) Implements SalesInterface.MobileSales.ISalesBase.ResultData

    Public Sub ReturnedData(ByVal Value As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.ReturnedData
        If ResultTo = "txtVoucherNo" Then
            Dim sSql As String
            sSql = "SELECT * from VoucherHeader where VoucherNo ='" & Value & "' order by VoucherNo"
            GetVoucher(sSql)
            GetTextBox()
        End If
        If ResultTo = "txtItemNo" Then
            txtItemNo.Text = Value
        End If

        If ResultTo = "dgvCustPrice" Then
            While InStr(Value, Chr(1)) > 0
                If dgvCustomerPrice.Rows.Count - 1 = rInd Then dgvCustomerPrice.Rows.Add()
                dgvCustomerPrice.Item(0, rInd).Value = Mid(Value, 1, InStr(Value, Chr(1)) - 1)
                Value = Mid(Value, InStr(Value, Chr(1)) + 1)
                rInd = rInd + 1
            End While
            If dgvCustomerPrice.Rows.Count - 1 = rInd Then dgvCustomerPrice.Rows.Add()
            dgvCustomerPrice.Item(0, rInd).Value = Value
            rInd = rInd + 1
        End If
    End Sub

    Public Event SearchField(ByVal FormName As String, ByVal FieldName As String, ByVal FieldType As String, ByVal CurValue As String) Implements SalesInterface.MobileSales.ISalesBase.SearchField
    Public Sub SetCulture(ByVal CultureName As String) Implements SalesInterface.MobileSales.ISalesBase.SetCulture

    End Sub
    Private Sub SaveVoucher()
        If txtVouNo.Text = String.Empty Then
            MsgBox("Please Enter Voucher No")
            Me.txtVouNo.Select()
            Return
        End If
        If txtFromVouNo.Text = String.Empty Then
            MsgBox("Please Enter Voucher No")
            Me.txtFromVouNo.Select()
            Return
        End If
        If txtToVouNo.Text = String.Empty Then
            MsgBox("Please Enter Voucher No")
            Me.txtToVouNo.Select()
            Return
        End If
        If IsNumeric(txtFromVouNo.Text) = False Then
            MsgBox("Please Enter Numeric No only")
            Me.txtFromVouNo.Select()
            Return
        End If
        If IsNumeric(txtToVouNo.Text) = False Then
            MsgBox("Please Enter Numeric No only")
            Me.txtToVouNo.Select()
            Return
        End If
        If CDbl(txtToVouNo.Text) < CDbl(txtFromVouNo.Text) Then
            MsgBox("Please Check the Voucher No")
            Me.txtToVouNo.Select()
            Return
        End If
        'If txtToVNo.Text = String.Empty Then
        '    MsgBox("Please Enter Voucher No")
        '    Me.txtToVNo.Select()
        '    Return
        'End If

        If txtAmount.Text = String.Empty Then
            MsgBox("Please Enter the Amount")
            Me.txtAmount.Select()
            Return
        End If

        If rbPercentage.Checked = True Then
            sAmtType = "Percentage"
        ElseIf rbAmount.Checked = True Then
            sAmtType = "Amount"
        Else
            sAmtType = "FOC"
        End If

        Dim iMulUse As Integer
        If chkMulUse.Checked = True Then
            iMulUse = 1
        Else
            iMulUse = 0
        End If

        Dim iAct As Integer
        If chkActive.Checked = True Then
            iAct = 1
        Else
            iAct = 0
        End If

        Dim strSql As String
        Try
            Dim iCnt As Integer
            If txtMinAmt.Text = "" Then txtMinAmt.Text = 0
            For iCnt = CInt(txtFromVouNo.Text) To CInt(txtToVouNo.Text)
                If IsVoucherExists(iCnt) = False Then
                    strSql = "Insert into Voucher(VouHeaderNo, VoucherNo, VoucherType, VoucherDate, ExpiryDate, Vouchervalue, Active, AmountType, MultipleUse, MinAmt, ItemNo, Redeemed) Values (" & objDO.SafeSQL(txtVouNo.Text) & "," & objDO.SafeSQL(txtVouPrefix.Text & iCnt) & ", " & objDO.SafeSQL(sVouType) & ", " & objDO.SafeSQL(Format(dtpVoucherDate.Value, "dd-MMM-yyyy")) & ", " & objDO.SafeSQL(Format(dtpExpDate.Value, "dd-MMM-yyyy")) & ", " & txtAmount.Text & "," & iAct & "," & objDO.SafeSQL(sAmtType) & "," & iMulUse & "," & txtMinAmt.Text & "," & objDO.SafeSQL(txtItemNo.Text) & ",0)"
                    objDO.ExecuteSQL(strSql)
                End If
            Next

            strSql = "Insert into VoucherHeader(VoucherNo, VouPrefix, VoucherFromNo, VoucherToNo, VoucherType, VoucherDate, ExpiryDate, Vouchervalue, Active, AmountType, MultipleUse, MinAmt, ItemNo) Values (" & objDO.SafeSQL(txtVouNo.Text) & "," & objDO.SafeSQL(txtVouPrefix.Text) & ", " & objDO.SafeSQL(txtFromVouNo.Text) & ", " & objDO.SafeSQL(txtToVouNo.Text) & "," & objDO.SafeSQL(sVouType) & ", " & objDO.SafeSQL(Format(dtpVoucherDate.Value, "dd-MMM-yyyy")) & ", " & objDO.SafeSQL(Format(dtpExpDate.Value, "dd-MMM-yyyy")) & ", " & txtAmount.Text & "," & iAct & "," & objDO.SafeSQL(sAmtType) & "," & iMulUse & "," & txtMinAmt.Text & "," & objDO.SafeSQL(txtItemNo.Text) & ")"
            objDO.ExecuteSQL(strSql)

            Dim cnt As Integer
            For cnt = 0 To dgvCustomerPrice.Rows.Count - 2
                strSql = "Insert into VoucherLineItem(VoucherNo, ItemNo, UOM, Qty) Values (" & objDO.SafeSQL(txtVouNo.Text) & "," & objDO.SafeSQL(dgvCustomerPrice.Item(0, cnt).Value.ToString) & "," & objDO.SafeSQL(dgvCustomerPrice.Item(2, cnt).Value.ToString) & "," & dgvCustomerPrice.Item(3, cnt).Value.ToString & ")"
                objDO.ExecuteSQL(strSql)
            Next
            IsNewRecord = False
            MsgBox("New Voucher Added Successfully")
            objDO.ExecuteSQL("Update MDT set LastVouNo=" & iVouNo & " where MdtNo = " & objDO.SafeSQL(sMDT))
            GetVoucher("SELECT * from VoucherHeader order by VoucherNo")
            RowNo = 0
            LoadRow()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Function IsVoucherExists(ByVal sVouNo As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = objDO.ReadRecord("Select VoucherNo from Voucher where VoucherNo = " & objDO.SafeSQL(sVouNo))
        bAns = dtr.Read
        dtr.Close()
        Return bAns
    End Function
    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim StrSql As String
        If txtVouNo.Text = "" Then
            MsgBox("Enter Voucher No to Delete")
            Exit Sub
        End If
        If MsgBox("Do you want to Delete the Voucher Y/N", MsgBoxStyle.YesNo, "Delete?") = MsgBoxResult.Yes Then
            Try
                StrSql = "Delete from Voucher where VouHeaderNo=" & objDO.SafeSQL(txtVouNo.Text)
                objDO.ExecuteSQL(StrSql)
                StrSql = "Delete from VoucherHeader where VoucherNo=" & objDO.SafeSQL(txtVouNo.Text)
                objDO.ExecuteSQL(StrSql)
                StrSql = "Delete from VoucherLineItem where VoucherNo=" & objDO.SafeSQL(txtVouNo.Text)
                objDO.ExecuteSQL(StrSql)
                MessageBox.Show("Voucher Deleted")
                clear()
                GetVoucher("Select * from VoucherHeader order by VoucherNo")
                RowNo = 0
                LoadRow()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        IsNewRecord = False
        RowNo = 0
        LoadRow()
    End Sub

    Private Sub Voucher_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        objDO.DisconnectDB()
    End Sub

    Private Sub Voucher_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objDO.ConnectionString = My.Settings.ConnectionString
        objDO.ConnectDB()
        Dim rsP As SqlDataReader
        rsP = objDO.ReadRecord("select IsPanel from System")
        If rsP.Read = True Then
            If CBool(rsP("IsPanel")) = True Then
                PnlButton.Visible = True
            Else
                PnlButton.Visible = False
            End If
        End If
        rsP.Close()
        Dim sSql As String
        cbGift.Checked = True
        sSql = "SELECT * from VoucherHeader"
        GetVoucher(sSql)
        GetTextBox()
    End Sub
    Private Sub GetVoucher(ByVal sSQL As String)
        myAdapter = objDO.GetDataAdapter(sSQL)
        myDataSet = New DataSet
        myAdapter.Fill(myDataSet, "VoucherHeader")
        myDataView = New DataView(myDataSet.Tables("VoucherHeader"))
        If myDataView.Count = 0 Then
            RowNo = -1
        Else
            RowNo = 0
        End If
        LoadRow()
    End Sub
    Private Sub LoadRow()
        If RowNo < 0 Or RowNo >= myDataView.Count Then Exit Sub
        txtVouNo.Text = myDataView(RowNo).Item("VoucherNo").ToString
        txtFromVouNo.Text = myDataView(RowNo).Item("VoucherFromNo").ToString
        txtToVouNo.Text = myDataView(RowNo).Item("VoucherToNo").ToString

        ' rsP = objDO.ReadRecord("select * from Voucher where VouHeaderNo=" & objDO.SafeSQL(txtVouNo.Text))
        'If rsP.Read = True Then
        dtpVoucherDate.Value = myDataView(RowNo).Item("VoucherDate").ToString
        dtpExpDate.Value = myDataView(RowNo).Item("ExpiryDate").ToString
        txtAmount.Text = myDataView(RowNo).Item("VoucherValue").ToString
        txtVouPrefix.Text = myDataView(RowNo).Item("VouPrefix").ToString
        txtItemNo.Text = myDataView(RowNo).Item("ItemNo").ToString
        If IsDBNull(myDataView(RowNo).Item("MinAmt").ToString) = True Then
            txtMinAmt.Text = 0
        Else
            txtMinAmt.Text = myDataView(RowNo).Item("MinAMt").ToString
        End If
        If myDataView(RowNo).Item("VoucherType").ToString = "Gift" Then
            cbGift.Checked = True
        Else
            cbPayment.Checked = True
        End If
        If myDataView(RowNo).Item("AmountType").ToString = "Amount" Then
            rbAmount.Checked = True
        ElseIf myDataView(RowNo).Item("AmountType").ToString = "Percentage" Then
            rbPercentage.Checked = True
        Else
            rbFOC.Checked = True
        End If
        '   txtReedeem.Text = myDataView(RowNo).Item("Reedeemed").ToString
        If CDbl(myDataView(RowNo).Item("MultipleUse")) = True Then
            chkMulUse.Checked = True
        Else
            chkMulUse.Checked = False
        End If
        If CDbl(myDataView(RowNo).Item("Active")) = True Then
            chkActive.Checked = True
        Else
            chkActive.Checked = False
        End If
    End Sub

    Public Sub ReturnSearch(ByVal SQL As String) Implements SalesInterface.MobileSales.ISalesBase.ReturnSearch
        If SQL = "" Then
            GetVoucher("Select * from VoucherHeader order by VoucherNo")
        Else
            GetVoucher("Select * from VoucherHeader  and " & SQL & "  order by VoucherNo")
        End If
    End Sub

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

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        SaveVoucher()
    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles pnlControls.Paint

    End Sub

    Private Sub rbPercentage_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbPercentage.CheckedChanged
        pnlGrid.Visible = False
        dgvCustomerPrice.Visible = False
        If rbPercentage.Checked = True Then
            sAmtType = "Percentage"
        ElseIf rbFOC.Checked = True Then
            sAmtType = "FOC"
        Else
            sAmtType = "Amount"
        End If
    End Sub

    Private Sub cbGift_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbGift.CheckedChanged
        If cbGift.Checked = True Then
            txtMinAmt.Visible = True
            txtItemNo.Visible = False
            btnItemNo.Visible = False
            lbMinAmt.Visible = True
            lbItemNo.Visible = False
            sVouType = "Gift"
            cbPayment.Checked = False
        Else
            txtItemNo.Visible = True
            btnItemNo.Visible = True
            txtMinAmt.Visible = False
            lbItemNo.Visible = True
            lbMinAmt.Visible = False
            sVouType = "Payment"
            cbGift.Checked = False
        End If
    End Sub

    Private Sub cbPayment_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbPayment.CheckedChanged
        If cbPayment.Checked = True Then
            txtItemNo.Visible = True
            btnItemNo.Visible = True
            txtMinAmt.Visible = False
            lbItemNo.Visible = True
            lbMinAmt.Visible = False
            cbGift.Checked = False
        Else
            txtMinAmt.Visible = True
            txtItemNo.Visible = False
            btnItemNo.Visible = False
            lbMinAmt.Visible = True
            lbItemNo.Visible = False
            cbPayment.Checked = False
        End If
    End Sub

    Private Sub rbFOC_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbFOC.CheckedChanged
        pnlGrid.Visible = True
        dgvCustomerPrice.Visible = True
        txtAmount.Text = "0"
    End Sub

    Private Sub rbAmount_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbAmount.CheckedChanged
        pnlGrid.Visible = False
        dgvCustomerPrice.Visible = False
    End Sub

    Private Sub dgvCustomerPrice_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCustomerPrice.CellContentClick

    End Sub

    Private Sub dgvCustomerPrice_CellMouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvCustomerPrice.CellMouseDown
        If e.RowIndex < 0 Then Exit Sub
        If e.ColumnIndex < 0 Then Exit Sub
        rInd = e.RowIndex
        If e.ColumnIndex = 0 Then
            RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\DataList.dll", "DataList.ItemList", "SalesPromo.Voucher", "dgvCustPrice", 0, 0)
        End If
    End Sub

    Private Sub dgvCustomerPrice_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCustomerPrice.CellValueChanged
        Try
            If e.RowIndex < 0 Then Exit Sub
            If e.ColumnIndex < 0 Then Exit Sub
            '  If IsNewRecord = False Then Exit Sub
            If e.ColumnIndex = 0 And dgvCustomerPrice.Item(0, e.RowIndex).Value <> "" Then
                sItemNo = dgvCustomerPrice.Item(0, e.RowIndex).Value.ToString
                Dim dtr As SqlDataReader
                dtr = objDO.ReadRecord("Select * from Item where ItemNo=" & objDO.SafeSQL(dgvCustomerPrice.Item(0, e.RowIndex).Value))
                If dtr.Read = True Then
                    dgvCustomerPrice.Item(1, e.RowIndex).Value = dtr("Description").ToString
                    Dim cb1 As New DataGridViewTextBoxCell
                    dgvCustomerPrice.Item(2, e.RowIndex) = cb1
                    dgvCustomerPrice.Item(2, e.RowIndex).Value = dtr("BaseUOM")
                    dgvCustomerPrice.Item(3, e.RowIndex).Value = "1"
                End If
                dtr.Close()
                'dgvCustomerPrice.Item(3, e.RowIndex).Value = "0"
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Label8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub txtVouNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVouNo.TextChanged
        If IsNewRecord = True Then Exit Sub
        Dim rsP As SqlDataReader
        dgvCustomerPrice.Rows.Clear()
        rsP = objDO.ReadRecord("select VoucherLineItem.ItemNo, Item.Description, VoucherLineItem.UOM, VoucherLineItem.Qty  from VoucherLineItem, Item where VoucherLineItem.ItemNo=Item.ItemNo and VoucherLineItem.VoucherNo=" & objDO.SafeSQL(txtVouNo.Text))
        Dim cnt As Integer
        While rsP.Read = True
            pnlGrid.Visible = True
            dgvCustomerPrice.Visible = True
            Dim row As String() = New String() _
            {rsP("ItemNo").ToString, rsP("Description").ToString, "", rsP("Qty").ToString}
            dgvCustomerPrice.Rows.Add(row)
            Dim cb As New DataGridViewTextBoxCell
            dgvCustomerPrice.Item(2, cnt) = cb
            dgvCustomerPrice.Item(2, cnt).Value = rsP("UOM").ToString
            dgvCustomerPrice.Item(3, cnt).Value = rsP("Qty").ToString
            cnt = cnt + 1
        End While
        rsP.Close()
    End Sub

    Private Sub btnItemNo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnItemNo.Click
        RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\DataList.dll", "DataList.ItemList", "SalesPromo.Voucher", "txtItemNo", 0, 0)
    End Sub

    Public Event LoadDataForm(ByVal Location As String, ByVal LoadType As String, ByVal ParentLoadType As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.LoadDataForm

    Public Event LoadForm(ByVal Location As String, ByVal LoadType As String) Implements SalesInterface.MobileSales.ISalesBase.LoadForm
End Class