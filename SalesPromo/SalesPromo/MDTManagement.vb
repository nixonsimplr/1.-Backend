Imports System
Imports System.Resources
Imports System.Threading
Imports System.Reflection
Imports System.Globalization
Imports DataInterface.IbizDO
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales

Public Class MDTManagement
    Implements ISalesBase



#Region "Localization Variants"
    Private rMgr As ResourceManager
    Private cInfo As CultureInfo
    Private sLang As String
#End Region
    Private aAgent As New ArrayList()
    Private IsNewRecord As Boolean = False
    Private IsProcess As Boolean = False
    Private myAdapter As SqlDataAdapter
    Private myDataSet As DataSet
    Private myDataView As DataView
    Private RowNo As Integer
    Private objDO As New DataInterface.IbizDO
    Dim sCode As String = ""
    Dim iInventory As Integer = 1
    Dim iMulUom As Integer = 1
    Dim iIsLive As Integer = 1
    Dim iPrintWLAN As Integer = 1
    Dim dHistory As Double = 0

    Private Sub MDTManagement_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        objDO.DisconnectDB()
    End Sub

    Private Sub MDTManagement_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objDO.ConnectionString = My.Settings.ConnectionString
        objDO.ConnectDB()
        chkWLAN.Checked = True
        LoadCombo()
        GetMDT("Select * from MDT")
    End Sub

    Private Sub cmbMDTID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim dtr As SqlDataReader
            dtr = objDO.ReadRecord("Select * from MDT where MDTNo=" & objDO.SafeSQL(txtMDTID.Text))
            While dtr.Read = True
                txtDescription.Text = dtr("Description").ToString
                cmbAgent.Text = dtr("AgentId").ToString
                cmbOrderSeq.Text = dtr("OrderBy").ToString
                cmbInvSeq.Text = dtr("PrintInvOrder").ToString
                cmbPrinterName.Text = dtr("PrinterName").ToString
                txtVehicle.Text = dtr("VehicleID").ToString
                txtCashCustNo.Text = dtr("CashCustNo").ToString
                txtLastOrdNo.Text = dtr("LastOrdNo").ToString
                txtOrdPrefix.Text = dtr("PreOrdNo").ToString
                txtOrdLen.Text = dtr("LenOrdNo").ToString
                txtLInvNo.Text = dtr("LastInvNo").ToString
                txtInvPrefix.Text = dtr("PreInvNo").ToString
                txtInvLen.Text = dtr("LenInvNo").ToString
                txtLRcptNo.Text = dtr("LastRcptNo").ToString
                txtRcptPrefix.Text = dtr("PreRcptNo").ToString
                txtRcptLen.Text = dtr("LenRcptNo").ToString
                txtLRetNo.Text = dtr("LastRtnNo").ToString
                txtRetPrefix.Text = dtr("PreRtnNo").ToString
                txtRetLen.Text = dtr("LenRtnNo").ToString
                txtLExNo.Text = dtr("LastExNo").ToString
                txtExPrefix.Text = dtr("PreExNo").ToString
                txtExLen.Text = dtr("LenExNo").ToString
                txtLCustNo.Text = dtr("LastCustNo").ToString
                txtCustPrefix.Text = dtr("PreCustNo").ToString
                txtCustLen.Text = dtr("LenCustNo").ToString
                txtLItemNo.Text = dtr("LastItNo").ToString
                txtItemPrefix.Text = dtr("PreItNo").ToString
                txtItemLen.Text = dtr("LenItNo").ToString
                txtLCrNo.Text = dtr("LastCrNo").ToString
                txtCrPrefix.Text = dtr("PreCrNo").ToString
                txtCrLen.Text = dtr("LenCrNo").ToString
                txtPreSerNo.Text = dtr("PreSerNo").ToString
                txtLenSerNo.Text = dtr("LenSerNo").ToString
                txtLSerNo.Text = dtr("LastSerNo").ToString
                txtPreStock.Text = dtr("PreReqNo").ToString
                txtLenStock.Text = dtr("LenReqNo").ToString
                txtLStockNo.Text = dtr("LastReqNo").ToString
                txtPreAsset.Text = dtr("PreAstNo").ToString
                txtLenAsset.Text = dtr("LenAstNo").ToString
                txtLAstNo.Text = dtr("LastAstNo").ToString
                txtPrintNetPort.Text = dtr("PrintNetPort").ToString
                ' MsgBox(dtr("PrintNetPort")).ToString()
                txtPrintIP.Text = dtr("PrintIP").ToString
                chkWLAN.Checked = CBool(dtr("PrintWLAN"))
                chkMulUOM.Checked = CBool(dtr("MultipleUOM"))
                ChkInventory.Checked = CBool(dtr("Inventory"))
                chkIsLive.Checked = CBool(dtr("ISLive"))
                'MsgBox(CBool(dtr("PrintWLAN")))
                txtPrinter.Text = dtr("Printer").ToString
                txtPrintPort.Text = dtr("PrintPort").ToString
                txtPrintBaud.Text = dtr("PrintBaud").ToString
            End While
            dtr.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub LoadCombo()
        IsProcess = True
        Dim dtr As SqlDataReader
        dtr = objDO.ReadRecord("Select HandheldType from Handheld")
        cmbHandheldType.Items.Clear()
        While dtr.Read
            cmbHandheldType.Items.Add(dtr("HandheldType").ToString)
        End While
        dtr.Close()
        'dtr = objDO.ReadRecord("Select Code from InvType")
        'cmbInvType.Items.Clear()
        'While dtr.Read
        '    cmbInvType.Items.Add(dtr("Code").ToString)
        'End While
        'dtr.Close()
        dtr = objDO.ReadRecord("Select Code, Name from SalesAgent")
        cmbAgent.DataSource = Nothing
        aAgent.Clear()
        'aAgent.Add(New ComboValues("ALL", "ALL"))
        While dtr.Read()
            aAgent.Add(New ComboValues(dtr("Code").ToString, dtr("Name").ToString))
            '    iSelIndex = iIndex
            'End If
            'iIndex = iIndex + 1
        End While
        dtr.Close()
        'MsgBox("Cmb")
        cmbAgent.DataSource = aAgent
        cmbAgent.DisplayMember = "Desc"
        cmbAgent.ValueMember = "Code"
        cmbAgent.SelectedIndex = 0
        dtr = objDO.ReadRecord("Select Code from Location")
        Do While dtr.Read = True
            cmbLocation.Items.Add(dtr("Code"))
        Loop
        dtr.Close()
        'If cmbAgent.Items.Count > 0 Then
        cmbOrderSeq.SelectedIndex = 0
        cmbInvSeq.SelectedIndex = 0
        'End If
        If cmbLocation.Items.Count > 0 Then
            cmbLocation.SelectedIndex = 0
        End If
        IsProcess = False
        If cmbPrinterName.Items.Count > 0 Then
            cmbPrinterName.SelectedIndex = 0
        End If
    End Sub
    Private Sub Clear()
        'Dim txtBox As Type = txtAgent.GetType

        'For Each ctrl As Control In Me.Controls
        '    If ctrl.GetType() Is txtBox Then
        '        CType(ctrl, TextBox).Clear()
        '    End If
        'Next
        txtMDTID.Select()
        txtMDTID.Text = ""
        txtVehicle.Text = String.Empty
        txtCashCustNo.Text = String.Empty
        txtDescription.Text = String.Empty
        txtLastOrdNo.Text = String.Empty
        txtOrdPrefix.Text = String.Empty
        txtOrdLen.Text = String.Empty
        txtLInvNo.Text = String.Empty
        txtInvPrefix.Text = String.Empty
        txtInvLen.Text = String.Empty
        txtLRcptNo.Text = String.Empty
        txtRcptPrefix.Text = String.Empty
        txtRcptLen.Text = String.Empty

        txtLRetNo.Text = String.Empty
        txtRetPrefix.Text = String.Empty
        txtRetLen.Text = String.Empty
        txtLExNo.Text = String.Empty
        txtExPrefix.Text = String.Empty
        txtExLen.Text = String.Empty
        txtLCustNo.Text = String.Empty
        txtCustPrefix.Text = String.Empty
        txtCustLen.Text = String.Empty
        txtLItemNo.Text = String.Empty
        txtItemPrefix.Text = String.Empty
        txtItemLen.Text = String.Empty
        txtLCrNo.Text = String.Empty
        txtCrPrefix.Text = String.Empty
        txtCrLen.Text = String.Empty
        txtPrintNetPort.Text = String.Empty
        txtPrintIP.Text = String.Empty
        chkWLAN.Checked = True
        chkMulUOM.Checked = True
        ChkInventory.Checked = True
        chkIsLive.Checked = False
        txtLSerNo.Text = String.Empty
        txtLenSerNo.Text = String.Empty
        txtPreSerNo.Text = String.Empty
        txtLStockNo.Text = String.Empty
        txtLenStock.Text = String.Empty
        txtPreStock.Text = String.Empty
        txtLAstNo.Text = String.Empty
        txtLenAsset.Text = String.Empty
        txtPreAsset.Text = String.Empty
        '  cmbAgent.Items.Clear()
        cmbLocation.Items.Clear()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim strUpt As String = ""
        Dim IsUpt As Boolean = False
        Dim dtr1 As SqlDataReader
        dtr1 = objDO.ReadRecord("Select MDTNo from MDT where MDTNo=" & objDO.SafeSQL(txtMDTID.Text))
        If dtr1.Read = True Then
            IsUpt = True
        Else
            IsUpt = False
        End If
        dtr1.Close()
        If IsUpt = True Then
            If ChkInventory.Checked = True Then
                iInventory = 1
            Else
                iInventory = 0
            End If
            If chkIsLive.Checked = True Then
                iIsLive = 1
            Else
                iIsLive = 0
            End If
            If chkMulUOM.Checked = True Then
                iMulUom = 1
            Else
                iMulUom = 0
            End If
            If chkWLAN.Checked = True Then
                iPrintWLAN = 1
            Else
                iPrintWLAN = 0
            End If

            If Me.txtDescription.Text = String.Empty Then
                MessageBox.Show("Please Enter Description", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtDescription.Select()
                Return
            End If
            'If Me.txtVehicle.Text = String.Empty Then
            '    MessageBox.Show("Please Enter Vehicle No", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
            '    Me.txtVehicle.Select()
            '    Return
            'End If
            If Me.txtLastOrdNo.Text = String.Empty Then
                MessageBox.Show("Please Enter Last Order No", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtLastOrdNo.Select()
                Return
            End If
            If Me.txtOrdPrefix.Text = String.Empty Then
                MessageBox.Show("Please Enter Order Prefix ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtOrdPrefix.Select()
                Return
            End If
            If Me.txtOrdLen.Text = String.Empty Then
                MessageBox.Show("Please Enter Order Length", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtOrdLen.Select()
                Return
            End If
            If Me.txtLInvNo.Text = String.Empty Then
                MessageBox.Show("Please Enter Last Invoice No", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtLInvNo.Select()
                Return
            End If
            If Me.txtLRcptNo.Text = String.Empty Then
                MessageBox.Show("Please Enter Last Receipt No", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtLRcptNo.Select()
                Return
            End If

            If Me.txtLRetNo.Text = String.Empty Then
                MessageBox.Show("Please Enter Last Return No", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtLRetNo.Select()
                Return
            End If
            If Me.txtLExNo.Text = String.Empty Then
                MessageBox.Show("Please Enter Last Exchange No", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtLExNo.Select()
                Return
            End If

            If Me.txtLCustNo.Text = String.Empty Then
                MessageBox.Show("Please Enter Last Customer No", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtLCustNo.Select()
                Return
            End If
            If Me.txtLItemNo.Text = String.Empty Then
                MessageBox.Show("Please Enter Last Item No", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtLItemNo.Select()
                Return
            End If
            If Me.txtLCrNo.Text = String.Empty Then
                MessageBox.Show("Please Enter Last CR No", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtLCrNo.Select()
                Return
            End If
            If Me.txtInvPrefix.Text = String.Empty Then
                MessageBox.Show("Please Enter Invoice Prefix", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtInvPrefix.Select()
                Return
            End If
            If Me.txtRcptPrefix.Text = String.Empty Then
                MessageBox.Show("Please Enter Receipt Prefix", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtRcptPrefix.Select()
                Return
            End If

            If Me.txtRetPrefix.Text = String.Empty Then
                MessageBox.Show("Please Enter Return Prefix", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtRetPrefix.Select()
                Return
            End If
            If Me.txtExPrefix.Text = String.Empty Then
                MessageBox.Show("Please Enter Exchange Prefix", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtExPrefix.Select()
                Return
            End If

            If Me.txtCustPrefix.Text = String.Empty Then
                MessageBox.Show("Please Enter Customer Prefix", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtCustPrefix.Select()
                Return
            End If

            If Me.txtItemPrefix.Text = String.Empty Then
                MessageBox.Show("Please Enter Item Prefix", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtItemPrefix.Select()
                Return
            End If
            If Me.txtCrPrefix.Text = String.Empty Then
                MessageBox.Show("Please Enter CR Prefix", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtCrPrefix.Select()
                Return
            End If
            If Me.txtInvLen.Text = String.Empty Then
                MessageBox.Show("Please Enter Invoice Length", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtInvLen.Select()
                Return
            End If
            If Me.txtRcptLen.Text = String.Empty Then
                MessageBox.Show("Please Enter Receipt Length", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtRcptLen.Select()
                Return
            End If

            If Me.txtRetLen.Text = String.Empty Then
                MessageBox.Show("Please Enter Return Length", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtRetLen.Select()
                Return
            End If
            If Me.txtExLen.Text = String.Empty Then
                MessageBox.Show("Please Enter Exchange Length", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtExLen.Select()
                Return
            End If

            If Me.txtCustLen.Text = String.Empty Then
                MessageBox.Show("Please Enter Customer Length", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtCustLen.Select()
                Return
            End If

            If Me.txtItemLen.Text = String.Empty Then
                MessageBox.Show("Please Enter Item Length", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtItemLen.Select()
                Return
            End If
            If Me.txtCrLen.Text = String.Empty Then
                MessageBox.Show("Please Enter CR Len", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtCrLen.Select()
                Return
            End If
            If Me.txtPrinter.Text = String.Empty Then
                MessageBox.Show("Please Enter Printer", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtPrinter.Select()
                Return
            End If
            If Me.txtPrintPort.Text = String.Empty Then
                MessageBox.Show("Please Enter Print Port", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtPrintPort.Select()
                Return
            End If
            If Me.txtPrintBaud.Text = String.Empty Then
                MessageBox.Show("Please Enter Print Baud", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtPrintBaud.Select()
                Return
            End If
            If Me.txtDotPrintPort.Text = String.Empty Then
                MessageBox.Show("Please Enter Dot Print Port", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtDotPrintPort.Select()
                Return
            End If
            If Me.txtDotPrintBaud.Text = String.Empty Then
                MessageBox.Show("Please Enter Dot Print Baud", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtDotPrintBaud.Select()
                Return
            End If

            Dim LastPromoNo, LenPromoNo, LastSerNo, LenSerNo, LastReqNo, LenReqNo, LenAstNo, LastAstNo As Integer

            If txtLPromoNo.Text <> String.Empty Then LastPromoNo = txtLPromoNo.Text
            If txtLenPromoNo.Text <> String.Empty Then LenPromoNo = txtLenPromoNo.Text
            If txtLSerNo.Text <> String.Empty Then LastSerNo = txtLSerNo.Text
            If txtLenSerNo.Text <> String.Empty Then LenSerNo = txtLenSerNo.Text
            If txtLenStock.Text <> String.Empty Then LenReqNo = txtLenStock.Text
            If txtLStockNo.Text <> String.Empty Then LastReqNo = txtLStockNo.Text
            If txtLenAsset.Text <> String.Empty Then LenAstNo = txtLenAsset.Text
            If txtLAstNo.Text <> String.Empty Then LastAstNo = txtLAstNo.Text
            If txtHistory.Text <> String.Empty Then
                dHistory = CDbl(txtHistory.Text)
            Else
                dHistory = 0
            End If


            strUpt = "update MDT Set MDTNo=" & objDO.SafeSQL(txtMDTID.Text) & ", " & _
            "Description=" & objDO.SafeSQL(txtDescription.Text) & ", VehicleID=" & objDO.SafeSQL(txtVehicle.Text) & ", LastOrdNo=" & txtLastOrdNo.Text & "," & _
            "PreOrdNo=" & objDO.SafeSQL(txtOrdPrefix.Text) & ", LenOrdNo=" & txtOrdLen.Text & "," & _
            "LastInvNo=" & txtLInvNo.Text & ",PreInvNo=" & objDO.SafeSQL(txtInvPrefix.Text) & "," & _
            "LenInvNo=" & txtInvLen.Text & ",LastRcptNo=" & txtLRcptNo.Text & "," & _
            "PreRcptNo=" & objDO.SafeSQL(txtRcptPrefix.Text) & ", LenRcptNo=" & txtRcptLen.Text & "," & _
            "AgentId=" & objDO.SafeSQL(cmbAgent.SelectedValue) & ", LastRtnNo=" & txtLRetNo.Text & "," & _
            "PreRtnNo=" & objDO.SafeSQL(txtRetPrefix.Text) & ", LenRtnNo=" & txtRetLen.Text & "," & _
            "LastExNo=" & txtLExNo.Text & ", PreExNo=" & objDO.SafeSQL(txtExPrefix.Text) & "," & _
            "LenExNo=" & txtExLen.Text & ", LastCustNo=" & txtLCustNo.Text & "," & _
            "PreCustNo=" & objDO.SafeSQL(txtCustPrefix.Text) & ", LenCustNo=" & txtCustLen.Text & "," & _
            "Location=" & objDO.SafeSQL(cmbLocation.Text) & ", PreITNo=" & objDO.SafeSQL(txtItemPrefix.Text) & "," & _
            "LenITNo=" & txtItemLen.Text & ", LastITNo=" & txtLItemNo.Text & "," & _
            "PreCRNo=" & objDO.SafeSQL(txtCrPrefix.Text) & ", LenCRNo=" & txtCrLen.Text & "," & _
            "LastCRNo=" & txtLCrNo.Text & ", DotPrintBaud=" & txtDotPrintBaud.Text & " ," & _
            "DotPrintPort=" & txtDotPrintPort.Text & " , PrintBaud=" & txtPrintBaud.Text & "," & _
            "PrintPort=" & txtPrintPort.Text & ", Printer=" & txtPrinter.Text & "," & _
            "PrinterName=" & objDO.SafeSQL(cmbPrinterName.Text) & ", PrintNetPort=" & txtPrintNetPort.Text & ", PrintIP=" & objDO.SafeSQL(txtPrintIP.Text) & "," & _
            "MultipleUOM=" & iMulUom & ",Inventory=" & iInventory & ",IsLive=" & iIsLive & ",PrintWLAN=" & iPrintWLAN & ", LastPrNo=" & LastPromoNo & "," & _
            "LenPrNo=" & LenPromoNo & ", PrePrNo=" & objDO.SafeSQL(txtPrePromoNo.Text) & "," & _
            "LastSerNo=" & LastSerNo & ", LenSerNo=" & LenSerNo & "," & _
            "PreSerNo=" & objDO.SafeSQL(txtPreSerNo.Text) & ", LastAstNo=" & LastAstNo & ", LenAstNo=" & LenAstNo & ", PreAstNo=" & objDO.SafeSQL(txtPreAsset.Text) & ", LastReqNo=" & LastReqNo & ", LenReqNo=" & LenReqNo & ", PreReqNo=" & objDO.SafeSQL(txtPreStock.Text) & ", PrintInvOrder=" & objDO.SafeSQL(cmbInvSeq.Text) & _
            ", OrderBy=" & objDO.SafeSQL(cmbOrderSeq.Text) & ", CashCustNo=" & objDO.SafeSQL(txtCashCustNo.Text) & ", HandheldType = " & objDO.SafeSQL(cmbHandheldType.Text) & ", InvType = " & objDO.SafeSQL("") & ", InvHistory = " & dHistory & " where MDTNo=" & objDO.SafeSQL(txtMDTID.Text)
            objDO.ExecuteSQL(strUpt)
            MessageBox.Show("MDT Modified", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Dim iRowNo As Integer = RowNo
            GetMDT("Select * from MDT")
            RowNo = iRowNo
            If RowNo <> 0 Then
                LoadRow()
            End If
        Else
            SaveMDT()
            GetMDT("Select * from MDT")
        End If
    End Sub

    Private Sub SaveMDT()
        Try
            If Me.txtDescription.Text = String.Empty Then
                MessageBox.Show("Please Enter Description", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtDescription.Select()
                Return
            End If
            If Me.txtVehicle.Text = String.Empty Then
                MessageBox.Show("Please Enter Vehicle No", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtVehicle.Select()
                Return
            End If
            If Me.txtLSerNo.Text = String.Empty Then
                MessageBox.Show("Please Enter Last Schedule No", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtLSerNo.Select()
                Return
            End If
            If Me.txtPreSerNo.Text = String.Empty Then
                MessageBox.Show("Please Enter Schedule Prefix ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtPreSerNo.Select()
                Return
            End If
            If Me.txtLenSerNo.Text = String.Empty Then
                MessageBox.Show("Please Enter Schedule Length", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtLenSerNo.Select()
                Return
            End If
            If Me.txtLInvNo.Text = String.Empty Then
                MessageBox.Show("Please Enter Last Invoice No", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtLInvNo.Select()
                Return
            End If
            If Me.txtLRcptNo.Text = String.Empty Then
                MessageBox.Show("Please Enter Last Receipt No", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtLRcptNo.Select()
                Return
            End If

            If Me.txtLRetNo.Text = String.Empty Then
                MessageBox.Show("Please Enter Last Return No", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtLRetNo.Select()
                Return
            End If
            If Me.txtLExNo.Text = String.Empty Then
                MessageBox.Show("Please Enter Last Exchange No", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtLExNo.Select()
                Return
            End If

            If Me.txtLCustNo.Text = String.Empty Then
                MessageBox.Show("Please Enter Last Customer No", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtLCustNo.Select()
                Return
            End If
            If Me.txtLItemNo.Text = String.Empty Then
                MessageBox.Show("Please Enter Last Item No", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtLItemNo.Select()
                Return
            End If
            If Me.txtLCrNo.Text = String.Empty Then
                MessageBox.Show("Please Enter Last CR No", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtLCrNo.Select()
                Return
            End If
            If Me.txtInvPrefix.Text = String.Empty Then
                MessageBox.Show("Please Enter Invoice Prefix", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtInvPrefix.Select()
                Return
            End If
            If Me.txtRcptPrefix.Text = String.Empty Then
                MessageBox.Show("Please Enter Receipt Prefix", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtRcptPrefix.Select()
                Return
            End If
            If Me.txtRetPrefix.Text = String.Empty Then
                MessageBox.Show("Please Enter Return Prefix", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtRetPrefix.Select()
                Return
            End If
            If Me.txtExPrefix.Text = String.Empty Then
                MessageBox.Show("Please Enter Exchange Prefix", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtExPrefix.Select()
                Return
            End If

            If Me.txtCustPrefix.Text = String.Empty Then
                MessageBox.Show("Please Enter Customer Prefix", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtCustPrefix.Select()
                Return
            End If

            If Me.txtItemPrefix.Text = String.Empty Then
                MessageBox.Show("Please Enter Item Prefix", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtItemPrefix.Select()
                Return
            End If
            If Me.txtCrPrefix.Text = String.Empty Then
                MessageBox.Show("Please Enter CR Prefix", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtCrPrefix.Select()
                Return
            End If
            If Me.txtInvLen.Text = String.Empty Then
                MessageBox.Show("Please Enter Invoice Length", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtInvLen.Select()
                Return
            End If
            If Me.txtRcptLen.Text = String.Empty Then
                MessageBox.Show("Please Enter Receipt Length", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtRcptLen.Select()
                Return
            End If

            If Me.txtRetLen.Text = String.Empty Then
                MessageBox.Show("Please Enter Return Length", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtRetLen.Select()
                Return
            End If
            If Me.txtExLen.Text = String.Empty Then
                MessageBox.Show("Please Enter Exchange Length", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtExLen.Select()
                Return
            End If

            If Me.txtCustLen.Text = String.Empty Then
                MessageBox.Show("Please Enter Customer Length", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtCustLen.Select()
                Return
            End If

            If Me.txtItemLen.Text = String.Empty Then
                MessageBox.Show("Please Enter Item Length", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtItemLen.Select()
                Return
            End If
            If Me.txtCrLen.Text = String.Empty Then
                MessageBox.Show("Please Enter CR Len", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtCrLen.Select()
                Return
            End If
            If Me.txtPrinter.Text = String.Empty Then
                MessageBox.Show("Please Enter Printer", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtPrinter.Select()
                Return
            End If
            If Me.txtPrintPort.Text = String.Empty Then
                MessageBox.Show("Please Enter Print Port", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtPrintPort.Select()
                Return
            End If
            If Me.txtPrintBaud.Text = String.Empty Then
                MessageBox.Show("Please Enter Print Baud", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtPrintBaud.Select()
                Return
            End If
            If Me.txtDotPrintPort.Text = String.Empty Then
                MessageBox.Show("Please Enter Dot Print Port", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtDotPrintPort.Select()
                Return
            End If
            If Me.txtDotPrintBaud.Text = String.Empty Then
                MessageBox.Show("Please Enter Dot Print Baud", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.txtDotPrintBaud.Select()
                Return
            End If
            If ChkInventory.Checked = True Then
                iInventory = 1
            Else
                iInventory = 0
            End If
            Dim strsql As String
            Dim LastPromoNo, LenPromoNo, LastSerNo, LenSerNo, LenReqNo, LastReqNo, LenAstNo, LastAstNo As Integer

            If ChkInventory.Checked = True Then
                iInventory = 1
            Else
                iInventory = 0
            End If
            If chkIsLive.Checked = True Then
                iIsLive = 1
            Else
                iIsLive = 0
            End If
            If chkMulUOM.Checked = True Then
                iMulUom = 1
            Else
                iMulUom = 0
            End If
            If chkWLAN.Checked = True Then
                iPrintWLAN = 1
            Else
                iPrintWLAN = 0
            End If

            If txtLPromoNo.Text <> String.Empty Then LastPromoNo = txtLPromoNo.Text
            If txtLenPromoNo.Text <> String.Empty Then LenPromoNo = txtLenPromoNo.Text
            If txtLSerNo.Text <> String.Empty Then LastSerNo = txtLSerNo.Text
            If txtLenSerNo.Text <> String.Empty Then LenSerNo = txtLenSerNo.Text
            If txtLenStock.Text <> String.Empty Then LenReqNo = txtLenStock.Text
            If txtLStockNo.Text <> String.Empty Then LastReqNo = txtLStockNo.Text
            If txtLenAsset.Text <> String.Empty Then LenAstNo = txtLenAsset.Text
            If txtLAstNo.Text <> String.Empty Then LastAstNo = txtLAstNo.Text
            strsql = "Insert into MDT  (MDTNo, Description, VehicleID, AgentId, Location, LastOrdNo," & _
            "PreOrdNo, LenOrdNo, LastInvNo, PreInvNo, LenInvNo, LastRcptNo, PreRcptNo," & _
            "LenRcptNo,  LastRtnNo, PreRtnNo, LenRtnNo, LenExNo, LastExNo, PreExNo," & _
            "LastCustNo, PreCustNo, LenCustNo, PreITNo, LenITNo, LastITNo, PreCRNo," & _
            "LenCRNo, LastCRNo,DotPrintBaud, DotPrintPort, PrintBaud, PrintPort," & _
            "Printer, PrintNetPort, PrintIP, PrintWLAN, LastPrNo, LenPrNo, PrePrNo, LastSerNo, LenSerNo, PreSerNo, LastReqNo, LenReqNo, PreReqNo, LastAstNo, LenAstNo, PreAstNo, CashCustNo, PrinterName, OrderBy, PrintInvOrder, MultipleUOM, IsLive, Inventory, HandheldType, InvType, InvHistory) Values(" & _
            "" & objDO.SafeSQL(txtMDTID.Text) & "," & objDO.SafeSQL(txtDescription.Text) & "," & objDO.SafeSQL(txtVehicle.Text) & "," & _
            "" & objDO.SafeSQL(cmbAgent.SelectedValue) & "," & objDO.SafeSQL(cmbLocation.Text) & "," & _
            "" & txtLastOrdNo.Text & "," & objDO.SafeSQL(txtOrdPrefix.Text) & "," & txtOrdLen.Text & "," & _
            "" & txtLInvNo.Text & "," & objDO.SafeSQL(txtInvPrefix.Text) & "," & _
            "" & txtInvLen.Text & "," & txtLRcptNo.Text & "," & objDO.SafeSQL(txtRcptPrefix.Text) & "," & _
            "" & txtRcptLen.Text & "," & txtLRetNo.Text & "," & objDO.SafeSQL(txtRetPrefix.Text) & "," & _
            "" & txtRetLen.Text & "," & txtExLen.Text & "," & txtLExNo.Text & "," & _
            "" & objDO.SafeSQL(txtExPrefix.Text) & "," & txtLCustNo.Text & "," & _
            "" & objDO.SafeSQL(txtCustPrefix.Text) & "," & txtCustLen.Text & "," & _
            "" & objDO.SafeSQL(txtItemPrefix.Text) & "," & txtItemLen.Text & "," & _
            "" & txtLItemNo.Text & "," & objDO.SafeSQL(txtCrPrefix.Text) & "," & _
            "" & txtCrLen.Text & "," & txtLCrNo.Text & "," & txtDotPrintBaud.Text & "," & _
            "" & txtDotPrintPort.Text & "," & txtPrintBaud.Text & "," & txtPrintPort.Text & "," & _
            "" & txtPrinter.Text & "," & txtPrintNetPort.Text & "," & objDO.SafeSQL(txtPrintIP.Text) & "," & _
            "" & iPrintWLAN & "," & LastPromoNo & "," & _
            "" & LenPromoNo & "," & objDO.SafeSQL(txtPrePromoNo.Text) & "," & _
            "" & LastSerNo & "," & LenSerNo & "," & objDO.SafeSQL(txtPreSerNo.Text) & "," & LastReqNo & "," & LenReqNo & "," & _
            "" & objDO.SafeSQL(txtPreStock.Text) & "," & LastAstNo & "," & LenAstNo & "," & objDO.SafeSQL(txtPreAsset.Text) & "," & objDO.SafeSQL(txtCashCustNo.Text) & "," & objDO.SafeSQL(cmbPrinterName.Text) & "," & _
            "" & objDO.SafeSQL(cmbOrderSeq.Text) & "," & objDO.SafeSQL(cmbInvSeq.Text) & "," & iMulUom & "," & iIsLive & "," & iInventory & "," & objDO.SafeSQL(cmbHandheldType.Text) & "," & objDO.SafeSQL("") & "," & dHistory & ")"
            objDO.ExecuteSQL(strsql)
            MessageBox.Show("New MDT Added", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Clear()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub GetMDT(ByVal sSQL As String)
        myAdapter = objDO.GetDataAdapter(sSQL)
        myDataSet = New DataSet
        myAdapter.Fill(myDataSet, "MDT")
        myDataView = New DataView(myDataSet.Tables("MDT"))
        If myDataView.Count = 0 Then
            RowNo = -1
        Else
            RowNo = 0
        End If
        LoadRow()
    End Sub

    Private Sub LoadRow()
        If RowNo < 0 Or RowNo >= myDataView.Count Then Exit Sub
        txtMDTID.Text = myDataView(RowNo).Item("MDTNo").ToString
        txtDescription.Text = myDataView(RowNo).Item("Description").ToString
        txtVehicle.Text = myDataView(RowNo).Item("VehicleID").ToString
        txtCashCustNo.Text = myDataView(RowNo).Item("CashCustNo").ToString
        cmbAgent.SelectedValue = myDataView(RowNo).Item("AgentId").ToString
        cmbLocation.Text = myDataView(RowNo).Item("Location").ToString
        cmbPrinterName.Text = myDataView(RowNo).Item("PrinterName").ToString
        cmbOrderSeq.Text = myDataView(RowNo).Item("OrderBy").ToString
        cmbInvSeq.Text = myDataView(RowNo).Item("PrintInvOrder").ToString
        cmbHandheldType.Text = myDataView(RowNo).Item("HandheldType").ToString
        cmbInvType.Text = myDataView(RowNo).Item("InvType").ToString
        txtHistory.Text = myDataView(RowNo).Item("InvHistory").ToString
        txtLastOrdNo.Text = myDataView(RowNo).Item("LastOrdNo").ToString
        txtOrdPrefix.Text = myDataView(RowNo).Item("PreOrdNo").ToString
        txtOrdLen.Text = myDataView(RowNo).Item("LenOrdNo").ToString
        txtLInvNo.Text = myDataView(RowNo).Item("LastInvNo").ToString
        txtInvPrefix.Text = myDataView(RowNo).Item("PreInvNo").ToString
        txtInvLen.Text = myDataView(RowNo).Item("LenInvNo").ToString
        txtLRcptNo.Text = myDataView(RowNo).Item("LastRcptNo").ToString
        txtRcptPrefix.Text = myDataView(RowNo).Item("PreRcptNo").ToString
        txtRcptLen.Text = myDataView(RowNo).Item("LenRcptNo").ToString
        txtLRetNo.Text = myDataView(RowNo).Item("LastRtnNo").ToString
        txtRetPrefix.Text = myDataView(RowNo).Item("PreRtnNo").ToString
        txtRetLen.Text = myDataView(RowNo).Item("LenRtnNo").ToString
        txtExLen.Text = myDataView(RowNo).Item("LenExNo").ToString
        txtLExNo.Text = myDataView(RowNo).Item("LastExNo").ToString
        txtExPrefix.Text = myDataView(RowNo).Item("PreExNo").ToString
        txtLCustNo.Text = myDataView(RowNo).Item("LastCustNo").ToString
        txtCustPrefix.Text = myDataView(RowNo).Item("PreCustNo").ToString
        txtCustLen.Text = myDataView(RowNo).Item("LenCustNo").ToString
        txtItemPrefix.Text = myDataView(RowNo).Item("PreITNo").ToString
        txtItemLen.Text = myDataView(RowNo).Item("LenITNo").ToString
        txtLItemNo.Text = myDataView(RowNo).Item("LastITNo").ToString
        txtCrPrefix.Text = myDataView(RowNo).Item("PreCRNo").ToString
        txtCrLen.Text = myDataView(RowNo).Item("LenCRNo").ToString
        txtLCrNo.Text = myDataView(RowNo).Item("LastCRNo").ToString
        txtPrinter.Text = myDataView(RowNo).Item("Printer").ToString
        txtPrintBaud.Text = myDataView(RowNo).Item("PrintBaud").ToString
        txtPrintPort.Text = myDataView(RowNo).Item("PrintPort").ToString
        txtDotPrintBaud.Text = myDataView(RowNo).Item("DotPrintBaud").ToString
        txtDotPrintPort.Text = myDataView(RowNo).Item("DotPrintPort").ToString
        txtPrintNetPort.Text = myDataView(RowNo).Item("PrintNetPort").ToString
        txtPrintIP.Text = myDataView(RowNo).Item("PrintIP").ToString
        chkWLAN.Checked = CBool(myDataView(RowNo).Item("PrintWLAN"))
        chkMulUOM.Checked = CBool(myDataView(RowNo).Item("MultipleUOM"))
        ChkInventory.Checked = CBool(myDataView(RowNo).Item("Inventory"))
        chkIsLive.Checked = CBool(myDataView(RowNo).Item("IsLive"))
        txtLPromoNo.Text = myDataView(RowNo).Item("LastPrNo").ToString
        txtLenPromoNo.Text = myDataView(RowNo).Item("LenPrNo").ToString
        txtPrePromoNo.Text = myDataView(RowNo).Item("PrePrNo").ToString
        txtLSerNo.Text = myDataView(RowNo).Item("LastSerNo").ToString
        txtLenSerNo.Text = myDataView(RowNo).Item("LenSerNo").ToString
        txtPreSerNo.Text = myDataView(RowNo).Item("PreSerNo").ToString
        txtLStockNo.Text = myDataView(RowNo).Item("LastReqNo").ToString
        txtLenStock.Text = myDataView(RowNo).Item("LenReqNo").ToString
        txtPreStock.Text = myDataView(RowNo).Item("PreReqNo").ToString
        txtLAstNo.Text = myDataView(RowNo).Item("LastAstNo").ToString
        txtLenAsset.Text = myDataView(RowNo).Item("LenAstNo").ToString
        txtPreAsset.Text = myDataView(RowNo).Item("PreAstNo").ToString
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Clear()
    End Sub

    Private Sub pnlTop_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles pnlTop.Paint

    End Sub

    Public Function GetListViewForm() As String Implements SalesInterface.MobileSales.ISalesBase.GetListViewForm
        RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\SalesPromo.dll", "SalesPromo.MDTList", "SalesPromo.MDTManagement", "txtMDTID", 0, 0)
        Return Windows.Forms.Application.StartupPath & "\SalesPromo.dll,SalesPromo.MDTList"
    End Function

    Public Sub ListViewClick() Implements SalesInterface.MobileSales.ISalesBase.ListViewClick

    End Sub

    Public Event LoadForm(ByVal Location As String, ByVal LoadType As String) Implements SalesInterface.MobileSales.ISalesBase.LoadForm

    Public Sub MoveFirstClick() Implements SalesInterface.MobileSales.ISalesBase.MoveFirstClick
        RowNo = 0
        LoadRow()
    End Sub

    Public Sub MoveLastClick() Implements SalesInterface.MobileSales.ISalesBase.MoveLastClick
        RowNo = myDataView.Count - 1
        LoadRow()
    End Sub

    Public Sub MoveNextClick() Implements SalesInterface.MobileSales.ISalesBase.MoveNextClick
        If RowNo >= myDataView.Count - 1 Then Exit Sub
        RowNo = RowNo + 1
        LoadRow()
    End Sub

    Public Sub MovePositionClick(ByVal Position As Long) Implements SalesInterface.MobileSales.ISalesBase.MovePositionClick

    End Sub

    Public Sub MovePreviousClick() Implements SalesInterface.MobileSales.ISalesBase.MovePreviousClick
        If RowNo <= 0 Then Exit Sub
        RowNo = RowNo - 1
        LoadRow()
    End Sub

    Public Sub Print() Implements SalesInterface.MobileSales.ISalesBase.Print

    End Sub

    Public Sub Print(ByVal PageNo As Integer) Implements SalesInterface.MobileSales.ISalesBase.Print

    End Sub

    Private Sub cmbAgent_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAgent.SelectedIndexChanged
        If IsProcess = True Then Exit Sub
        Dim dtr As SqlDataReader
        'MsgBox("CmbAgent")
        dtr = objDO.ReadRecordAnother("Select AgentCode, AgentID from Team where AgentID=" & objDO.SafeSQL(cmbAgent.SelectedValue))
        If dtr.Read = True Then
            txtVehicle.Text = dtr("AgentID").ToString
            cmbLocation.Text = dtr("AgentCode").ToString
            txtCashCustNo.Text = ""
        End If
        dtr.Close()
        'dtr = objDO.ReadRecordAnother("Select Name from SalesAgent where Code=" & objDO.SafeSQL(cmbAgent.Text))
        'If dtr.Read = True Then
        '    cmbAgent.Text = dtr("Name").ToString
        'End If
        'dtr.Close()
    End Sub

    Private Sub gbOrderNo_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gbOrderNo.Enter

    End Sub

    Public Event LoadDataForm(ByVal Location As String, ByVal LoadType As String, ByVal ParentLoadType As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.LoadDataForm

    Public Event ResultData(ByVal ChildLoadType As String, ByVal Value As String) Implements SalesInterface.MobileSales.ISalesBase.ResultData

    Public Sub ReturnedData(ByVal Value As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.ReturnedData
        If ResultTo = "txtMDTID" Then
            GetMDT("SELECT * from MDT where MDTNo='" & Value & "' order by MDTNo")
            Exit Sub
        End If
    End Sub


    Private Sub gbItem_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gbItem.Enter

    End Sub

    Private Sub pnlMiddle_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles pnlMiddle.Paint

    End Sub

    Private Sub pnlBottom_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles pnlBottom.Paint

    End Sub

    '=====================================================================================
    'Search Function
    Public Sub ReturnSearch(ByVal SQL As String) Implements SalesInterface.MobileSales.ISalesBase.ReturnSearch
        If SQL = "" Then
            GetMDT("Select * from MDT")
        Else
            GetMDT("Select * from MDT Where " & SQL)
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

    'Localization
    Private Sub Localization()
        Try
            cInfo = New CultureInfo(sLang)
            Thread.CurrentThread.CurrentCulture = cInfo
            Thread.CurrentThread.CurrentUICulture = cInfo
            rMgr = New ResourceManager("SalesPromo.SalesPromo", [Assembly].GetExecutingAssembly())

            Me.Text = rMgr.GetString("MDTManagement")
            btnSave.Text = rMgr.GetString("btnSave")
            btnCancel.Text = rMgr.GetString("btnCancel")

            MDT_lblMDTID.Text = rMgr.GetString("MDT_lblMDTID")
            MDT_lblDesc.Text = rMgr.GetString("MDT_lblDesc")
            MDT_lblAgent.Text = rMgr.GetString("MDT_lblAgent")
            '        MDT_Location.Text = rMgr.GetString("Inv_lblLocation")

            gbOrderNo.Text = rMgr.GetString("gbOrderNo")
            MDT_lblLastOrdNo.Text = rMgr.GetString("MDT_lblLastOrdNo")
            MDT_lblOrdPrefix.Text = rMgr.GetString("MDT_lblOrdPrefix")
            MDT_lblOrdLen.Text = rMgr.GetString("MDT_lblOrdLen")

            gbInvoice.Text = rMgr.GetString("gbInvoice")
            MDT_lblInvLast.Text = rMgr.GetString("MDT_lblInvLast")
            MDT_lblInvPrefix.Text = rMgr.GetString("MDT_lblOrdPrefix")
            MDT_lblInvLen.Text = rMgr.GetString("MDT_lblOrdLen")

            gbReceipt.Text = rMgr.GetString("gbReceipt")
            MDT_lblRecptLast.Text = rMgr.GetString("MDT_lblRecptLast")
            MDT_lblRecptPrefix.Text = rMgr.GetString("MDT_lblOrdPrefix")
            MDT_lblRecptLen.Text = rMgr.GetString("MDT_lblOrdLen")

            gbReturn.Text = rMgr.GetString("gbReturn")
            MDT_lblRetLast.Text = rMgr.GetString("MDT_lblRetLast")
            MDT_lblRetPref.Text = rMgr.GetString("MDT_lblOrdPrefix")
            MDT_lblRetLen.Text = rMgr.GetString("MDT_lblOrdLen")

            gbCustomer.Text = rMgr.GetString("gbCustomer")
            MDT_lblCustLast.Text = rMgr.GetString("MDT_lblCustLast")
            MDT_lblCustPref.Text = rMgr.GetString("MDT_lblOrdPrefix")
            MDT_lblCustLen.Text = rMgr.GetString("MDT_lblOrdLen")

            gbExchange.Text = rMgr.GetString("gbExchange")
            MDT_lblExchLast.Text = rMgr.GetString("MDT_lblExchLast")
            MDT_lblExchPref.Text = rMgr.GetString("MDT_lblOrdPrefix")
            MDT_lblExchLen.Text = rMgr.GetString("MDT_lblOrdLen")

            gbItem.Text = rMgr.GetString("gbItem")
            MDT_lblItemLast.Text = rMgr.GetString("MDT_lblItemLast")
            MDT_lblItemPref.Text = rMgr.GetString("MDT_lblOrdPrefix")
            MDT_lblItemLen.Text = rMgr.GetString("MDT_lblOrdLen")

            gbCreditNote.Text = rMgr.GetString("gbCreditNote")
            MDT_lblCreditLast.Text = rMgr.GetString("MDT_lblCreditLast")
            MDT_lblCreditPref.Text = rMgr.GetString("MDT_lblOrdPrefix")
            MDT_lblCreditLen.Text = rMgr.GetString("MDT_lblOrdLen")

            gbPrinterSettings.Text = rMgr.GetString("gbPrinterSettings")
            MDT_lblPrintPort.Text = rMgr.GetString("MDT_lblPrintPort")
            MDT_lblPrintBaud.Text = rMgr.GetString("MDT_lblPrintBaud")
            MDT_lblDotPrint.Text = rMgr.GetString("MDT_lblDotPrint")
            MDT_lblDotPrintBaud.Text = rMgr.GetString("MDT_lblDotPrintBaud")
            MDT_lblPrinter.Text = rMgr.GetString("MDT_lblPrinter")

            gbItemTrans.Text = rMgr.GetString("gbItemTrans")
            MDT_lblPrinterNet.Text = rMgr.GetString("MDT_lblPrinterNet")
            MDT_lblPrinterIP.Text = rMgr.GetString("MDT_lblPrinterIP")
            MDT_lblAreYou.Text = rMgr.GetString("MDT_lblAreYou")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub SetCulture(ByVal CultureName As String) Implements SalesInterface.MobileSales.ISalesBase.SetCulture
        sLang = CultureName
        '  Localization()
    End Sub

    Private Sub btnSave_Layout(ByVal sender As Object, ByVal e As System.Windows.Forms.LayoutEventArgs) Handles btnSave.Layout

    End Sub

    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim StrSql As String
        If MsgBox("Do you want to Delete Y/N", MsgBoxStyle.YesNo, "Delete?") = MsgBoxResult.Yes Then
            Try
                StrSql = "Delete from MDT where MDTNo=" & objDO.SafeSQL(txtMDTID.Text)
                objDO.ExecuteSQL(StrSql)
                MessageBox.Show("MDT Deleted")
                '  MessageBox.Show(rMgr.GetString("Message_RecDel"), rMgr.GetString("Message_RecDel_Caption"), MessageBoxButtons.OK)
                '   objDO.ExecuteSQL("Update MDT set LastExNo=" & iExNo & " where MdtNo = " & objDO.SafeSQL(sMDT))
                Clear()
                GetMDT("Select * from MDT")
                RowNo = 0
                LoadRow()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Clear()
        IsNewRecord = True
    End Sub

    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub cmbPrinterName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbPrinterName.SelectedIndexChanged

    End Sub

    Private Sub MDT_lblAgent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        If MsgBox("Note: It will clear the Salesman Data permanently" & vbCrLf & "Do you want to Clear the Salesman :" & cmbAgent.Text & " Data Yes/No", MsgBoxStyle.YesNo, "Delete?") = MsgBoxResult.Yes Then
            objDO.ExecuteSQL("Delete from Invoice where AgentID=" & objDO.SafeSQL(cmbAgent.SelectedValue))
            objDO.ExecuteSQL("Delete from InvItem Where InvNo in (Select InvNo from Invoice where AgentID=" & objDO.SafeSQL(cmbAgent.SelectedValue) & ")")
            objDO.ExecuteSQL("Delete from OrderHdr where AgentID=" & objDO.SafeSQL(cmbAgent.SelectedValue))
            objDO.ExecuteSQL("Delete from OrdItem Where OrdNo in (Select OrdNo from OrderHdr where AgentID=" & objDO.SafeSQL(cmbAgent.SelectedValue) & ")")
            objDO.ExecuteSQL("Delete from GoodsExchange where SalesPersonCode=" & objDO.SafeSQL(cmbAgent.SelectedValue))
            objDO.ExecuteSQL("Delete from GoodsExchangeItem Where ExchangeNo  in (Select ExchangeNo from GoodsExchange where SalesPersonCode=" & objDO.SafeSQL(cmbAgent.SelectedValue) & ")")
            objDO.ExecuteSQL("Delete from GoodsReturn where SalesPersonCode=" & objDO.SafeSQL(cmbAgent.SelectedValue))
            objDO.ExecuteSQL("Delete from GoodsReturnItem Where ReturnNo in (Select ReturnNo from GoodsReturn where SalesPersonCode=" & objDO.SafeSQL(cmbAgent.SelectedValue) & ")")
            objDO.ExecuteSQL("Delete from CreditNote where SalesPersonCode=" & objDO.SafeSQL(cmbAgent.SelectedValue))
            objDO.ExecuteSQL("Delete from CreditNoteDet Where CreditNoteNo in (Select CreditNoteNo from CreditNote where SalesPersonCode=" & objDO.SafeSQL(cmbAgent.SelectedValue) & ")")
            objDO.ExecuteSQL("Delete from Receipt where AgentID=" & objDO.SafeSQL(cmbAgent.SelectedValue))
            objDO.ExecuteSQL("Delete from RcptItem Where RcptNo in (Select RcptNo from Receipt where AgentID=" & objDO.SafeSQL(cmbAgent.SelectedValue) & ")")
            objDO.ExecuteSQL("Delete from StockOrder where AgentID=" & objDO.SafeSQL(cmbAgent.SelectedValue))
            objDO.ExecuteSQL("Delete from StockOrderItem Where StockNo in (Select StockNo from StockOrder where AgentID=" & objDO.SafeSQL(cmbAgent.SelectedValue) & ")")
            objDO.ExecuteSQL("Delete from ItemTrans where Location=" & objDO.SafeSQL(cmbLocation.Text))
            objDO.ExecuteSQL("Delete from GoodsInvn where Location=" & objDO.SafeSQL(cmbLocation.Text))
            objDO.ExecuteSQL("Delete from AssetTrans where AgentID=" & objDO.SafeSQL(cmbAgent.SelectedValue))
            objDO.ExecuteSQL("Delete from CustVisit where AgentID=" & objDO.SafeSQL(cmbAgent.SelectedValue))
            MsgBox(cmbAgent.Text & " Data cleared Successfully")
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbInvType.SelectedIndexChanged

    End Sub

    Private Sub Label14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub txtMDTID_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtMDTID.MouseDown
        RaiseEvent SearchField(Me.ProductName & "." & Me.Name, "MDTNo", "String", txtMDTID.Text)
    End Sub
End Class