<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Voucher
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim OrdDtLabel As System.Windows.Forms.Label
        Dim lblVouNo As System.Windows.Forms.Label
        Dim Label2 As System.Windows.Forms.Label
        Dim Label3 As System.Windows.Forms.Label
        Dim Label4 As System.Windows.Forms.Label
        Dim Label5 As System.Windows.Forms.Label
        Dim Label6 As System.Windows.Forms.Label
        Dim Label7 As System.Windows.Forms.Label
        Dim lblItemNo As System.Windows.Forms.Label
        Dim lblMinAmt As System.Windows.Forms.Label
        Dim Label8 As System.Windows.Forms.Label
        Dim Label9 As System.Windows.Forms.Label
        Me.dtpExpDate = New System.Windows.Forms.DateTimePicker
        Me.txtVouNo = New System.Windows.Forms.TextBox
        Me.pnlControls = New System.Windows.Forms.Panel
        Me.txtVouPrefix = New System.Windows.Forms.TextBox
        Me.txtFromVouNo = New System.Windows.Forms.TextBox
        Me.rbPayment = New System.Windows.Forms.RadioButton
        Me.Label1 = New System.Windows.Forms.Label
        Me.rbGift = New System.Windows.Forms.RadioButton
        Me.txtToVouNo = New System.Windows.Forms.TextBox
        Me.txtItemNo = New System.Windows.Forms.TextBox
        Me.btnItemNo = New System.Windows.Forms.Button
        Me.txtMinAmt = New System.Windows.Forms.TextBox
        Me.chkMulUse = New System.Windows.Forms.CheckBox
        Me.chkActive = New System.Windows.Forms.CheckBox
        Me.dtpVoucherDate = New System.Windows.Forms.DateTimePicker
        Me.txtAmount = New System.Windows.Forms.TextBox
        Me.lbMinAmt = New System.Windows.Forms.Label
        Me.lbItemNo = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.cbPayment = New System.Windows.Forms.CheckBox
        Me.cbGift = New System.Windows.Forms.CheckBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.rbFOC = New System.Windows.Forms.RadioButton
        Me.rbAmount = New System.Windows.Forms.RadioButton
        Me.rbPercentage = New System.Windows.Forms.RadioButton
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.PnlButton = New System.Windows.Forms.Panel
        Me.pnlGrid = New System.Windows.Forms.Panel
        Me.dgvCustomerPrice = New System.Windows.Forms.DataGridView
        Me.ItemNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ItemName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.UOM = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        OrdDtLabel = New System.Windows.Forms.Label
        lblVouNo = New System.Windows.Forms.Label
        Label2 = New System.Windows.Forms.Label
        Label3 = New System.Windows.Forms.Label
        Label4 = New System.Windows.Forms.Label
        Label5 = New System.Windows.Forms.Label
        Label6 = New System.Windows.Forms.Label
        Label7 = New System.Windows.Forms.Label
        lblItemNo = New System.Windows.Forms.Label
        lblMinAmt = New System.Windows.Forms.Label
        Label8 = New System.Windows.Forms.Label
        Label9 = New System.Windows.Forms.Label
        Me.pnlControls.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.PnlButton.SuspendLayout()
        Me.pnlGrid.SuspendLayout()
        CType(Me.dgvCustomerPrice, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'OrdDtLabel
        '
        OrdDtLabel.AutoSize = True
        OrdDtLabel.Location = New System.Drawing.Point(7, 182)
        OrdDtLabel.Name = "OrdDtLabel"
        OrdDtLabel.Size = New System.Drawing.Size(161, 13)
        OrdDtLabel.TabIndex = 64
        OrdDtLabel.Text = "Expiry Date . . . . . . . . . . . . . ."
        '
        'lblVouNo
        '
        lblVouNo.AutoSize = True
        lblVouNo.Location = New System.Drawing.Point(7, 14)
        lblVouNo.Name = "lblVouNo"
        lblVouNo.Size = New System.Drawing.Size(153, 13)
        lblVouNo.TabIndex = 62
        lblVouNo.Text = "Voucher No . . . . . . . . . . . . ."
        AddHandler lblVouNo.Click, AddressOf Me.OrdNoLabel1_Click
        '
        'Label2
        '
        Label2.AutoSize = True
        Label2.Location = New System.Drawing.Point(7, 211)
        Label2.Name = "Label2"
        Label2.Size = New System.Drawing.Size(160, 13)
        Label2.TabIndex = 85
        Label2.Text = "Amount. . . . . . . . . . . . . . . . ."
        '
        'Label3
        '
        Label3.AutoSize = True
        Label3.Location = New System.Drawing.Point(7, 334)
        Label3.Name = "Label3"
        Label3.Size = New System.Drawing.Size(162, 13)
        Label3.TabIndex = 87
        Label3.Text = "Amount Type . . . . . . . . . . . . ."
        AddHandler Label3.Click, AddressOf Me.Label3_Click
        '
        'Label4
        '
        Label4.AutoSize = True
        Label4.Location = New System.Drawing.Point(7, 151)
        Label4.Name = "Label4"
        Label4.Size = New System.Drawing.Size(149, 13)
        Label4.TabIndex = 90
        Label4.Text = "Voucher Date . . . . . . . . . . ."
        '
        'Label5
        '
        Label5.AutoSize = True
        Label5.Location = New System.Drawing.Point(7, 268)
        Label5.Name = "Label5"
        Label5.Size = New System.Drawing.Size(149, 13)
        Label5.TabIndex = 93
        Label5.Text = "Active . . . . . . . . . . . . . . . ."
        '
        'Label6
        '
        Label6.AutoSize = True
        Label6.Location = New System.Drawing.Point(7, 239)
        Label6.Name = "Label6"
        Label6.Size = New System.Drawing.Size(148, 13)
        Label6.TabIndex = 96
        Label6.Text = "Multiple Use . . . . . . . . . . . ."
        '
        'Label7
        '
        Label7.AutoSize = True
        Label7.Location = New System.Drawing.Point(7, 53)
        Label7.Name = "Label7"
        Label7.Size = New System.Drawing.Size(143, 13)
        Label7.TabIndex = 101
        Label7.Text = "Voucher Type . . . . . . . . . ."
        '
        'lblItemNo
        '
        lblItemNo.AutoSize = True
        lblItemNo.Location = New System.Drawing.Point(9, 247)
        lblItemNo.Name = "lblItemNo"
        lblItemNo.Size = New System.Drawing.Size(0, 13)
        lblItemNo.TabIndex = 108
        lblItemNo.Visible = False
        '
        'lblMinAmt
        '
        lblMinAmt.AutoSize = True
        lblMinAmt.Location = New System.Drawing.Point(10, 245)
        lblMinAmt.Name = "lblMinAmt"
        lblMinAmt.Size = New System.Drawing.Size(0, 13)
        lblMinAmt.TabIndex = 107
        lblMinAmt.Visible = False
        '
        'Label8
        '
        Label8.AutoSize = True
        Label8.Location = New System.Drawing.Point(7, 119)
        Label8.Name = "Label8"
        Label8.Size = New System.Drawing.Size(163, 13)
        Label8.TabIndex = 115
        Label8.Text = "Voucher No From. . . . . . . . . . ."
        AddHandler Label8.Click, AddressOf Me.Label8_Click
        '
        'Label9
        '
        Label9.AutoSize = True
        Label9.Location = New System.Drawing.Point(7, 89)
        Label9.Name = "Label9"
        Label9.Size = New System.Drawing.Size(161, 13)
        Label9.TabIndex = 209
        Label9.Text = "Voucher Prefix . . . . . . . . . . . ."
        '
        'dtpExpDate
        '
        Me.dtpExpDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpExpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpExpDate.Location = New System.Drawing.Point(157, 178)
        Me.dtpExpDate.Name = "dtpExpDate"
        Me.dtpExpDate.Size = New System.Drawing.Size(101, 21)
        Me.dtpExpDate.TabIndex = 204
        '
        'txtVouNo
        '
        Me.txtVouNo.Location = New System.Drawing.Point(156, 11)
        Me.txtVouNo.Name = "txtVouNo"
        Me.txtVouNo.Size = New System.Drawing.Size(100, 21)
        Me.txtVouNo.TabIndex = 63
        '
        'pnlControls
        '
        Me.pnlControls.Controls.Add(Me.txtVouPrefix)
        Me.pnlControls.Controls.Add(Label9)
        Me.pnlControls.Controls.Add(Me.txtFromVouNo)
        Me.pnlControls.Controls.Add(Label8)
        Me.pnlControls.Controls.Add(Me.rbPayment)
        Me.pnlControls.Controls.Add(Me.Label1)
        Me.pnlControls.Controls.Add(Me.rbGift)
        Me.pnlControls.Controls.Add(Me.txtToVouNo)
        Me.pnlControls.Controls.Add(Me.txtItemNo)
        Me.pnlControls.Controls.Add(Me.btnItemNo)
        Me.pnlControls.Controls.Add(Label7)
        Me.pnlControls.Controls.Add(Me.txtMinAmt)
        Me.pnlControls.Controls.Add(Me.chkMulUse)
        Me.pnlControls.Controls.Add(Label6)
        Me.pnlControls.Controls.Add(Me.chkActive)
        Me.pnlControls.Controls.Add(Label5)
        Me.pnlControls.Controls.Add(Me.dtpVoucherDate)
        Me.pnlControls.Controls.Add(Label4)
        Me.pnlControls.Controls.Add(Me.txtAmount)
        Me.pnlControls.Controls.Add(Label2)
        Me.pnlControls.Controls.Add(Me.dtpExpDate)
        Me.pnlControls.Controls.Add(Me.txtVouNo)
        Me.pnlControls.Controls.Add(OrdDtLabel)
        Me.pnlControls.Controls.Add(lblVouNo)
        Me.pnlControls.Controls.Add(lblItemNo)
        Me.pnlControls.Controls.Add(lblMinAmt)
        Me.pnlControls.Controls.Add(Me.lbMinAmt)
        Me.pnlControls.Controls.Add(Me.lbItemNo)
        Me.pnlControls.Controls.Add(Me.GroupBox1)
        Me.pnlControls.Controls.Add(Me.GroupBox2)
        Me.pnlControls.Controls.Add(Label3)
        Me.pnlControls.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlControls.Location = New System.Drawing.Point(0, 0)
        Me.pnlControls.Name = "pnlControls"
        Me.pnlControls.Size = New System.Drawing.Size(805, 629)
        Me.pnlControls.TabIndex = 93
        '
        'txtVouPrefix
        '
        Me.txtVouPrefix.Location = New System.Drawing.Point(156, 86)
        Me.txtVouPrefix.Name = "txtVouPrefix"
        Me.txtVouPrefix.Size = New System.Drawing.Size(100, 21)
        Me.txtVouPrefix.TabIndex = 190
        '
        'txtFromVouNo
        '
        Me.txtFromVouNo.Location = New System.Drawing.Point(156, 116)
        Me.txtFromVouNo.Name = "txtFromVouNo"
        Me.txtFromVouNo.Size = New System.Drawing.Size(100, 21)
        Me.txtFromVouNo.TabIndex = 200
        '
        'rbPayment
        '
        Me.rbPayment.AutoSize = True
        Me.rbPayment.Location = New System.Drawing.Point(505, 115)
        Me.rbPayment.Name = "rbPayment"
        Me.rbPayment.Size = New System.Drawing.Size(67, 17)
        Me.rbPayment.TabIndex = 103
        Me.rbPayment.TabStop = True
        Me.rbPayment.Text = "Payment"
        Me.rbPayment.UseVisualStyleBackColor = True
        Me.rbPayment.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(271, 116)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(17, 13)
        Me.Label1.TabIndex = 114
        Me.Label1.Text = "to"
        '
        'rbGift
        '
        Me.rbGift.AutoCheck = False
        Me.rbGift.AutoSize = True
        Me.rbGift.Location = New System.Drawing.Point(416, 115)
        Me.rbGift.Name = "rbGift"
        Me.rbGift.Size = New System.Drawing.Size(42, 17)
        Me.rbGift.TabIndex = 102
        Me.rbGift.TabStop = True
        Me.rbGift.Text = "Gift"
        Me.rbGift.UseVisualStyleBackColor = True
        Me.rbGift.Visible = False
        '
        'txtToVouNo
        '
        Me.txtToVouNo.Location = New System.Drawing.Point(297, 113)
        Me.txtToVouNo.Name = "txtToVouNo"
        Me.txtToVouNo.Size = New System.Drawing.Size(100, 21)
        Me.txtToVouNo.TabIndex = 201
        '
        'txtItemNo
        '
        Me.txtItemNo.Location = New System.Drawing.Point(158, 297)
        Me.txtItemNo.Name = "txtItemNo"
        Me.txtItemNo.Size = New System.Drawing.Size(100, 21)
        Me.txtItemNo.TabIndex = 208
        Me.txtItemNo.Visible = False
        '
        'btnItemNo
        '
        Me.btnItemNo.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnItemNo.Location = New System.Drawing.Point(260, 295)
        Me.btnItemNo.Name = "btnItemNo"
        Me.btnItemNo.Size = New System.Drawing.Size(24, 23)
        Me.btnItemNo.TabIndex = 104
        Me.btnItemNo.Text = "..."
        Me.btnItemNo.UseVisualStyleBackColor = True
        Me.btnItemNo.Visible = False
        '
        'txtMinAmt
        '
        Me.txtMinAmt.Location = New System.Drawing.Point(159, 296)
        Me.txtMinAmt.Name = "txtMinAmt"
        Me.txtMinAmt.Size = New System.Drawing.Size(100, 21)
        Me.txtMinAmt.TabIndex = 100
        Me.txtMinAmt.Visible = False
        '
        'chkMulUse
        '
        Me.chkMulUse.AutoSize = True
        Me.chkMulUse.Location = New System.Drawing.Point(158, 239)
        Me.chkMulUse.Name = "chkMulUse"
        Me.chkMulUse.Size = New System.Drawing.Size(15, 14)
        Me.chkMulUse.TabIndex = 206
        Me.chkMulUse.UseVisualStyleBackColor = True
        '
        'chkActive
        '
        Me.chkActive.AutoSize = True
        Me.chkActive.Location = New System.Drawing.Point(158, 268)
        Me.chkActive.Name = "chkActive"
        Me.chkActive.Size = New System.Drawing.Size(15, 14)
        Me.chkActive.TabIndex = 207
        Me.chkActive.UseVisualStyleBackColor = True
        '
        'dtpVoucherDate
        '
        Me.dtpVoucherDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpVoucherDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpVoucherDate.Location = New System.Drawing.Point(157, 147)
        Me.dtpVoucherDate.Name = "dtpVoucherDate"
        Me.dtpVoucherDate.Size = New System.Drawing.Size(101, 21)
        Me.dtpVoucherDate.TabIndex = 203
        '
        'txtAmount
        '
        Me.txtAmount.Location = New System.Drawing.Point(158, 208)
        Me.txtAmount.Name = "txtAmount"
        Me.txtAmount.Size = New System.Drawing.Size(100, 21)
        Me.txtAmount.TabIndex = 205
        '
        'lbMinAmt
        '
        Me.lbMinAmt.AutoSize = True
        Me.lbMinAmt.Location = New System.Drawing.Point(7, 301)
        Me.lbMinAmt.Name = "lbMinAmt"
        Me.lbMinAmt.Size = New System.Drawing.Size(175, 13)
        Me.lbMinAmt.TabIndex = 110
        Me.lbMinAmt.Text = "Minimum Amount. . . . . . . . . . . . ."
        Me.lbMinAmt.Visible = False
        '
        'lbItemNo
        '
        Me.lbItemNo.AutoSize = True
        Me.lbItemNo.Location = New System.Drawing.Point(7, 300)
        Me.lbItemNo.Name = "lbItemNo"
        Me.lbItemNo.Size = New System.Drawing.Size(161, 13)
        Me.lbItemNo.TabIndex = 109
        Me.lbItemNo.Text = "Item No. . . . . . . . . . . . . . . . ."
        Me.lbItemNo.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cbPayment)
        Me.GroupBox1.Controls.Add(Me.cbGift)
        Me.GroupBox1.Location = New System.Drawing.Point(154, 38)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(240, 35)
        Me.GroupBox1.TabIndex = 111
        Me.GroupBox1.TabStop = False
        '
        'cbPayment
        '
        Me.cbPayment.AutoSize = True
        Me.cbPayment.Location = New System.Drawing.Point(103, 13)
        Me.cbPayment.Name = "cbPayment"
        Me.cbPayment.Size = New System.Drawing.Size(68, 17)
        Me.cbPayment.TabIndex = 116
        Me.cbPayment.Text = "Payment"
        Me.cbPayment.UseVisualStyleBackColor = True
        '
        'cbGift
        '
        Me.cbGift.AutoSize = True
        Me.cbGift.Location = New System.Drawing.Point(6, 13)
        Me.cbGift.Name = "cbGift"
        Me.cbGift.Size = New System.Drawing.Size(43, 17)
        Me.cbGift.TabIndex = 115
        Me.cbGift.Text = "Gift"
        Me.cbGift.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.rbFOC)
        Me.GroupBox2.Controls.Add(Me.rbAmount)
        Me.GroupBox2.Controls.Add(Me.rbPercentage)
        Me.GroupBox2.Location = New System.Drawing.Point(156, 320)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(351, 35)
        Me.GroupBox2.TabIndex = 112
        Me.GroupBox2.TabStop = False
        '
        'rbFOC
        '
        Me.rbFOC.AutoSize = True
        Me.rbFOC.Location = New System.Drawing.Point(171, 12)
        Me.rbFOC.Name = "rbFOC"
        Me.rbFOC.Size = New System.Drawing.Size(46, 17)
        Me.rbFOC.TabIndex = 211
        Me.rbFOC.TabStop = True
        Me.rbFOC.Text = "FOC"
        Me.rbFOC.UseVisualStyleBackColor = True
        '
        'rbAmount
        '
        Me.rbAmount.AutoSize = True
        Me.rbAmount.Location = New System.Drawing.Point(6, 12)
        Me.rbAmount.Name = "rbAmount"
        Me.rbAmount.Size = New System.Drawing.Size(62, 17)
        Me.rbAmount.TabIndex = 209
        Me.rbAmount.TabStop = True
        Me.rbAmount.Text = "Amount"
        Me.rbAmount.UseVisualStyleBackColor = True
        '
        'rbPercentage
        '
        Me.rbPercentage.AutoSize = True
        Me.rbPercentage.Location = New System.Drawing.Point(85, 12)
        Me.rbPercentage.Name = "rbPercentage"
        Me.rbPercentage.Size = New System.Drawing.Size(80, 17)
        Me.rbPercentage.TabIndex = 210
        Me.rbPercentage.TabStop = True
        Me.rbPercentage.Text = "Percentage"
        Me.rbPercentage.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDelete.Location = New System.Drawing.Point(609, 6)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(93, 23)
        Me.btnDelete.TabIndex = 79
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(510, 6)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(93, 23)
        Me.btnSave.TabIndex = 76
        Me.btnSave.Text = "&Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnNew.Location = New System.Drawing.Point(411, 6)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(93, 23)
        Me.btnNew.TabIndex = 75
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(708, 6)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(93, 23)
        Me.btnCancel.TabIndex = 77
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'PnlButton
        '
        Me.PnlButton.Controls.Add(Me.btnDelete)
        Me.PnlButton.Controls.Add(Me.btnSave)
        Me.PnlButton.Controls.Add(Me.btnCancel)
        Me.PnlButton.Controls.Add(Me.btnNew)
        Me.PnlButton.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PnlButton.Location = New System.Drawing.Point(0, 629)
        Me.PnlButton.Name = "PnlButton"
        Me.PnlButton.Size = New System.Drawing.Size(805, 35)
        Me.PnlButton.TabIndex = 92
        '
        'pnlGrid
        '
        Me.pnlGrid.Controls.Add(Me.dgvCustomerPrice)
        Me.pnlGrid.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlGrid.Location = New System.Drawing.Point(0, 379)
        Me.pnlGrid.Name = "pnlGrid"
        Me.pnlGrid.Size = New System.Drawing.Size(805, 250)
        Me.pnlGrid.TabIndex = 94
        '
        'dgvCustomerPrice
        '
        Me.dgvCustomerPrice.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvCustomerPrice.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ItemNo, Me.ItemName, Me.UOM, Me.Qty})
        Me.dgvCustomerPrice.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvCustomerPrice.Location = New System.Drawing.Point(0, 0)
        Me.dgvCustomerPrice.MultiSelect = False
        Me.dgvCustomerPrice.Name = "dgvCustomerPrice"
        Me.dgvCustomerPrice.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvCustomerPrice.Size = New System.Drawing.Size(805, 250)
        Me.dgvCustomerPrice.TabIndex = 220
        Me.dgvCustomerPrice.Visible = False
        '
        'ItemNo
        '
        Me.ItemNo.HeaderText = "Item No"
        Me.ItemNo.Name = "ItemNo"
        Me.ItemNo.ReadOnly = True
        '
        'ItemName
        '
        Me.ItemName.HeaderText = "Description"
        Me.ItemName.Name = "ItemName"
        Me.ItemName.ReadOnly = True
        Me.ItemName.Width = 250
        '
        'UOM
        '
        Me.UOM.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.[Nothing]
        Me.UOM.HeaderText = "UOM"
        Me.UOM.Name = "UOM"
        Me.UOM.ReadOnly = True
        Me.UOM.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.UOM.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'Qty
        '
        Me.Qty.HeaderText = "Quantity"
        Me.Qty.Name = "Qty"
        '
        'Voucher
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(805, 664)
        Me.Controls.Add(Me.pnlGrid)
        Me.Controls.Add(Me.pnlControls)
        Me.Controls.Add(Me.PnlButton)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "Voucher"
        Me.Text = "Voucher"
        Me.pnlControls.ResumeLayout(False)
        Me.pnlControls.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.PnlButton.ResumeLayout(False)
        Me.pnlGrid.ResumeLayout(False)
        CType(Me.dgvCustomerPrice, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dtpExpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtVouNo As System.Windows.Forms.TextBox
    Friend WithEvents pnlControls As System.Windows.Forms.Panel
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents PnlButton As System.Windows.Forms.Panel
    Friend WithEvents txtAmount As System.Windows.Forms.TextBox
    Friend WithEvents rbPercentage As System.Windows.Forms.RadioButton
    Friend WithEvents rbAmount As System.Windows.Forms.RadioButton
    Friend WithEvents dtpVoucherDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents chkActive As System.Windows.Forms.CheckBox
    Friend WithEvents chkMulUse As System.Windows.Forms.CheckBox
    Friend WithEvents txtMinAmt As System.Windows.Forms.TextBox
    Friend WithEvents rbPayment As System.Windows.Forms.RadioButton
    Friend WithEvents rbGift As System.Windows.Forms.RadioButton
    Friend WithEvents txtItemNo As System.Windows.Forms.TextBox
    Friend WithEvents btnItemNo As System.Windows.Forms.Button
    Friend WithEvents lbMinAmt As System.Windows.Forms.Label
    Friend WithEvents lbItemNo As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtToVouNo As System.Windows.Forms.TextBox
    Friend WithEvents rbFOC As System.Windows.Forms.RadioButton
    Friend WithEvents pnlGrid As System.Windows.Forms.Panel
    Friend WithEvents cbPayment As System.Windows.Forms.CheckBox
    Friend WithEvents cbGift As System.Windows.Forms.CheckBox
    Friend WithEvents dgvCustomerPrice As System.Windows.Forms.DataGridView
    Friend WithEvents txtFromVouNo As System.Windows.Forms.TextBox
    Friend WithEvents txtVouPrefix As System.Windows.Forms.TextBox
    Friend WithEvents ItemNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ItemName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents UOM As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Qty As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
