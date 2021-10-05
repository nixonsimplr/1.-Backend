<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PO
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
        Dim lblTAmt As System.Windows.Forms.Label
        Dim lblGST As System.Windows.Forms.Label
        Dim lblSub As System.Windows.Forms.Label
        Dim lblDis As System.Windows.Forms.Label
        Dim lblCustName As System.Windows.Forms.Label
        Dim lblOrdDt As System.Windows.Forms.Label
        Dim lblCustNo As System.Windows.Forms.Label
        Dim lblRefNo As System.Windows.Forms.Label
        Dim lblAgent As System.Windows.Forms.Label
        Dim lblPTerms As System.Windows.Forms.Label
        Dim lblPoNo As System.Windows.Forms.Label
        Dim CurCodeLabel1 As System.Windows.Forms.Label
        Dim CurExRateLabel1 As System.Windows.Forms.Label
        Me.TabControl2 = New System.Windows.Forms.TabControl
        Me.tpInvoice = New System.Windows.Forms.TabPage
        Me.InvNoTextBox = New System.Windows.Forms.TextBox
        Me.lblVoid = New System.Windows.Forms.Label
        Me.btnPay = New System.Windows.Forms.Button
        Me.btnAgent = New System.Windows.Forms.Button
        Me.btnCustNo = New System.Windows.Forms.Button
        Me.gbAmount = New System.Windows.Forms.GroupBox
        Me.TotalAmtLabel = New System.Windows.Forms.Label
        Me.GSTAmtLabel = New System.Windows.Forms.Label
        Me.SubTotalLabel = New System.Windows.Forms.Label
        Me.DiscountLabel = New System.Windows.Forms.Label
        Me.CustNameTextBox = New System.Windows.Forms.TextBox
        Me.InvDtDateTimePicker = New System.Windows.Forms.DateTimePicker
        Me.CustIdTextBox = New System.Windows.Forms.TextBox
        Me.PONoTextBox = New System.Windows.Forms.TextBox
        Me.AgentIDTextBox = New System.Windows.Forms.TextBox
        Me.PayTermsTextBox = New System.Windows.Forms.TextBox
        Me.tpCur = New System.Windows.Forms.TabPage
        Me.Button4 = New System.Windows.Forms.Button
        Me.CurCodeTextBox = New System.Windows.Forms.TextBox
        Me.CurExRateTextBox = New System.Windows.Forms.TextBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.btnPrint = New System.Windows.Forms.Button
        Me.btnVoid = New System.Windows.Forms.Button
        Me.btnModify = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.dgvOrdItem = New System.Windows.Forms.DataGridView
        Me.ItemNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.VariantName = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ItemName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.WareHouse = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.UOM = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.Quantity = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Price = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DisPr = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DisPer = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Amount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BaseUOM = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BasePrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.FOC = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DeliveredDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Promotion = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Offer = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Priority = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.QtyShip = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Explode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Discount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Panel3 = New System.Windows.Forms.Panel
        lblTAmt = New System.Windows.Forms.Label
        lblGST = New System.Windows.Forms.Label
        lblSub = New System.Windows.Forms.Label
        lblDis = New System.Windows.Forms.Label
        lblCustName = New System.Windows.Forms.Label
        lblOrdDt = New System.Windows.Forms.Label
        lblCustNo = New System.Windows.Forms.Label
        lblRefNo = New System.Windows.Forms.Label
        lblAgent = New System.Windows.Forms.Label
        lblPTerms = New System.Windows.Forms.Label
        lblPoNo = New System.Windows.Forms.Label
        CurCodeLabel1 = New System.Windows.Forms.Label
        CurExRateLabel1 = New System.Windows.Forms.Label
        Me.TabControl2.SuspendLayout()
        Me.tpInvoice.SuspendLayout()
        Me.gbAmount.SuspendLayout()
        Me.tpCur.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.dgvOrdItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl2
        '
        Me.TabControl2.Controls.Add(Me.tpInvoice)
        Me.TabControl2.Controls.Add(Me.tpCur)
        Me.TabControl2.Dock = System.Windows.Forms.DockStyle.Top
        Me.TabControl2.Location = New System.Drawing.Point(0, 0)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(731, 257)
        Me.TabControl2.TabIndex = 8
        '
        'tpInvoice
        '
        Me.tpInvoice.AutoScroll = True
        Me.tpInvoice.Controls.Add(Me.InvNoTextBox)
        Me.tpInvoice.Controls.Add(Me.lblVoid)
        Me.tpInvoice.Controls.Add(Me.btnPay)
        Me.tpInvoice.Controls.Add(Me.btnAgent)
        Me.tpInvoice.Controls.Add(Me.btnCustNo)
        Me.tpInvoice.Controls.Add(Me.gbAmount)
        Me.tpInvoice.Controls.Add(Me.CustNameTextBox)
        Me.tpInvoice.Controls.Add(Me.InvDtDateTimePicker)
        Me.tpInvoice.Controls.Add(Me.CustIdTextBox)
        Me.tpInvoice.Controls.Add(Me.PONoTextBox)
        Me.tpInvoice.Controls.Add(Me.AgentIDTextBox)
        Me.tpInvoice.Controls.Add(Me.PayTermsTextBox)
        Me.tpInvoice.Controls.Add(lblCustName)
        Me.tpInvoice.Controls.Add(lblOrdDt)
        Me.tpInvoice.Controls.Add(lblCustNo)
        Me.tpInvoice.Controls.Add(lblRefNo)
        Me.tpInvoice.Controls.Add(lblAgent)
        Me.tpInvoice.Controls.Add(lblPTerms)
        Me.tpInvoice.Controls.Add(lblPoNo)
        Me.tpInvoice.Location = New System.Drawing.Point(4, 22)
        Me.tpInvoice.Name = "tpInvoice"
        Me.tpInvoice.Padding = New System.Windows.Forms.Padding(3)
        Me.tpInvoice.Size = New System.Drawing.Size(723, 231)
        Me.tpInvoice.TabIndex = 0
        Me.tpInvoice.Text = "Ordering"
        Me.tpInvoice.UseVisualStyleBackColor = True
        '
        'InvNoTextBox
        '
        Me.InvNoTextBox.Location = New System.Drawing.Point(151, 18)
        Me.InvNoTextBox.Name = "InvNoTextBox"
        Me.InvNoTextBox.Size = New System.Drawing.Size(100, 21)
        Me.InvNoTextBox.TabIndex = 37
        '
        'lblVoid
        '
        Me.lblVoid.AutoSize = True
        Me.lblVoid.Font = New System.Drawing.Font("Arial Black", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVoid.ForeColor = System.Drawing.Color.Red
        Me.lblVoid.Location = New System.Drawing.Point(500, 3)
        Me.lblVoid.Name = "lblVoid"
        Me.lblVoid.Size = New System.Drawing.Size(94, 27)
        Me.lblVoid.TabIndex = 35
        Me.lblVoid.Text = "VOIDED"
        Me.lblVoid.Visible = False
        '
        'btnPay
        '
        Me.btnPay.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPay.Location = New System.Drawing.Point(251, 167)
        Me.btnPay.Name = "btnPay"
        Me.btnPay.Size = New System.Drawing.Size(24, 23)
        Me.btnPay.TabIndex = 31
        Me.btnPay.Text = "..."
        Me.btnPay.UseVisualStyleBackColor = True
        '
        'btnAgent
        '
        Me.btnAgent.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAgent.Location = New System.Drawing.Point(251, 196)
        Me.btnAgent.Name = "btnAgent"
        Me.btnAgent.Size = New System.Drawing.Size(24, 23)
        Me.btnAgent.TabIndex = 30
        Me.btnAgent.Text = "..."
        Me.btnAgent.UseVisualStyleBackColor = True
        '
        'btnCustNo
        '
        Me.btnCustNo.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCustNo.Location = New System.Drawing.Point(251, 80)
        Me.btnCustNo.Name = "btnCustNo"
        Me.btnCustNo.Size = New System.Drawing.Size(24, 23)
        Me.btnCustNo.TabIndex = 29
        Me.btnCustNo.Text = "..."
        Me.btnCustNo.UseVisualStyleBackColor = True
        '
        'gbAmount
        '
        Me.gbAmount.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbAmount.Controls.Add(lblTAmt)
        Me.gbAmount.Controls.Add(Me.TotalAmtLabel)
        Me.gbAmount.Controls.Add(lblGST)
        Me.gbAmount.Controls.Add(Me.GSTAmtLabel)
        Me.gbAmount.Controls.Add(lblSub)
        Me.gbAmount.Controls.Add(Me.SubTotalLabel)
        Me.gbAmount.Controls.Add(lblDis)
        Me.gbAmount.Controls.Add(Me.DiscountLabel)
        Me.gbAmount.Location = New System.Drawing.Point(388, 23)
        Me.gbAmount.Name = "gbAmount"
        Me.gbAmount.Size = New System.Drawing.Size(313, 166)
        Me.gbAmount.TabIndex = 28
        Me.gbAmount.TabStop = False
        Me.gbAmount.Text = "Calculation"
        '
        'lblTAmt
        '
        lblTAmt.AutoSize = True
        lblTAmt.Location = New System.Drawing.Point(16, 124)
        lblTAmt.Name = "lblTAmt"
        lblTAmt.Size = New System.Drawing.Size(141, 13)
        lblTAmt.TabIndex = 6
        lblTAmt.Text = "Total Amount . . . . . . . . . ."
        '
        'TotalAmtLabel
        '
        Me.TotalAmtLabel.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TotalAmtLabel.Location = New System.Drawing.Point(170, 124)
        Me.TotalAmtLabel.Name = "TotalAmtLabel"
        Me.TotalAmtLabel.Size = New System.Drawing.Size(122, 17)
        Me.TotalAmtLabel.TabIndex = 7
        Me.TotalAmtLabel.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblGST
        '
        lblGST.AutoSize = True
        lblGST.Location = New System.Drawing.Point(16, 94)
        lblGST.Name = "lblGST"
        lblGST.Size = New System.Drawing.Size(143, 13)
        lblGST.TabIndex = 4
        lblGST.Text = "GST Amount . . . . . . . . . . ."
        '
        'GSTAmtLabel
        '
        Me.GSTAmtLabel.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GSTAmtLabel.Location = New System.Drawing.Point(170, 94)
        Me.GSTAmtLabel.Name = "GSTAmtLabel"
        Me.GSTAmtLabel.Size = New System.Drawing.Size(122, 17)
        Me.GSTAmtLabel.TabIndex = 5
        Me.GSTAmtLabel.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblSub
        '
        lblSub.AutoSize = True
        lblSub.Location = New System.Drawing.Point(16, 64)
        lblSub.Name = "lblSub"
        lblSub.Size = New System.Drawing.Size(143, 13)
        lblSub.TabIndex = 2
        lblSub.Text = "Sub Total . . . . . . . . . . . . ."
        '
        'SubTotalLabel
        '
        Me.SubTotalLabel.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SubTotalLabel.Location = New System.Drawing.Point(170, 64)
        Me.SubTotalLabel.Name = "SubTotalLabel"
        Me.SubTotalLabel.Size = New System.Drawing.Size(122, 17)
        Me.SubTotalLabel.TabIndex = 3
        Me.SubTotalLabel.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblDis
        '
        lblDis.AutoSize = True
        lblDis.Location = New System.Drawing.Point(16, 34)
        lblDis.Name = "lblDis"
        lblDis.Size = New System.Drawing.Size(146, 13)
        lblDis.TabIndex = 0
        lblDis.Text = "Discount . . . . . . . . . . . . . ."
        '
        'DiscountLabel
        '
        Me.DiscountLabel.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DiscountLabel.Location = New System.Drawing.Point(170, 34)
        Me.DiscountLabel.Name = "DiscountLabel"
        Me.DiscountLabel.Size = New System.Drawing.Size(122, 17)
        Me.DiscountLabel.TabIndex = 1
        Me.DiscountLabel.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'CustNameTextBox
        '
        Me.CustNameTextBox.Location = New System.Drawing.Point(151, 110)
        Me.CustNameTextBox.Name = "CustNameTextBox"
        Me.CustNameTextBox.Size = New System.Drawing.Size(200, 21)
        Me.CustNameTextBox.TabIndex = 27
        '
        'InvDtDateTimePicker
        '
        Me.InvDtDateTimePicker.Location = New System.Drawing.Point(151, 50)
        Me.InvDtDateTimePicker.Name = "InvDtDateTimePicker"
        Me.InvDtDateTimePicker.Size = New System.Drawing.Size(200, 21)
        Me.InvDtDateTimePicker.TabIndex = 3
        '
        'CustIdTextBox
        '
        Me.CustIdTextBox.Location = New System.Drawing.Point(151, 81)
        Me.CustIdTextBox.Name = "CustIdTextBox"
        Me.CustIdTextBox.Size = New System.Drawing.Size(100, 21)
        Me.CustIdTextBox.TabIndex = 5
        '
        'PONoTextBox
        '
        Me.PONoTextBox.Location = New System.Drawing.Point(151, 139)
        Me.PONoTextBox.Name = "PONoTextBox"
        Me.PONoTextBox.Size = New System.Drawing.Size(100, 21)
        Me.PONoTextBox.TabIndex = 7
        '
        'AgentIDTextBox
        '
        Me.AgentIDTextBox.Location = New System.Drawing.Point(151, 197)
        Me.AgentIDTextBox.Name = "AgentIDTextBox"
        Me.AgentIDTextBox.Size = New System.Drawing.Size(100, 21)
        Me.AgentIDTextBox.TabIndex = 9
        '
        'PayTermsTextBox
        '
        Me.PayTermsTextBox.Location = New System.Drawing.Point(151, 168)
        Me.PayTermsTextBox.Name = "PayTermsTextBox"
        Me.PayTermsTextBox.Size = New System.Drawing.Size(100, 21)
        Me.PayTermsTextBox.TabIndex = 19
        '
        'lblCustName
        '
        lblCustName.AutoSize = True
        lblCustName.Location = New System.Drawing.Point(6, 113)
        lblCustName.Name = "lblCustName"
        lblCustName.Size = New System.Drawing.Size(139, 13)
        lblCustName.TabIndex = 26
        lblCustName.Text = "Customer Name . . . . . . . ."
        '
        'lblOrdDt
        '
        lblOrdDt.AutoSize = True
        lblOrdDt.Location = New System.Drawing.Point(6, 54)
        lblOrdDt.Name = "lblOrdDt"
        lblOrdDt.Size = New System.Drawing.Size(180, 13)
        lblOrdDt.TabIndex = 2
        lblOrdDt.Text = "PO Date . . . . . . . . . . . . . . . . . . ."
        '
        'lblCustNo
        '
        lblCustNo.AutoSize = True
        lblCustNo.Location = New System.Drawing.Point(6, 84)
        lblCustNo.Name = "lblCustNo"
        lblCustNo.Size = New System.Drawing.Size(142, 13)
        lblCustNo.TabIndex = 4
        lblCustNo.Text = "Customer No . . . . . . . . . . "
        '
        'lblRefNo
        '
        lblRefNo.AutoSize = True
        lblRefNo.Location = New System.Drawing.Point(6, 142)
        lblRefNo.Name = "lblRefNo"
        lblRefNo.Size = New System.Drawing.Size(139, 13)
        lblRefNo.TabIndex = 6
        lblRefNo.Text = "Customer Ref. No. . . . . . ."
        '
        'lblAgent
        '
        lblAgent.AutoSize = True
        lblAgent.Location = New System.Drawing.Point(6, 200)
        lblAgent.Name = "lblAgent"
        lblAgent.Size = New System.Drawing.Size(138, 13)
        lblAgent.TabIndex = 8
        lblAgent.Text = "Sales Person Code . . . . . ."
        '
        'lblPTerms
        '
        lblPTerms.AutoSize = True
        lblPTerms.Location = New System.Drawing.Point(6, 171)
        lblPTerms.Name = "lblPTerms"
        lblPTerms.Size = New System.Drawing.Size(137, 13)
        lblPTerms.TabIndex = 18
        lblPTerms.Text = "Payment Terms . . . . . . . ."
        '
        'lblPoNo
        '
        lblPoNo.AutoSize = True
        lblPoNo.Location = New System.Drawing.Point(6, 21)
        lblPoNo.Name = "lblPoNo"
        lblPoNo.Size = New System.Drawing.Size(177, 13)
        lblPoNo.TabIndex = 36
        lblPoNo.Text = "PO No . . . . . . . . . . . . . . . . . . . ."
        '
        'tpCur
        '
        Me.tpCur.Controls.Add(Me.Button4)
        Me.tpCur.Controls.Add(Me.CurCodeTextBox)
        Me.tpCur.Controls.Add(Me.CurExRateTextBox)
        Me.tpCur.Controls.Add(CurCodeLabel1)
        Me.tpCur.Controls.Add(CurExRateLabel1)
        Me.tpCur.Location = New System.Drawing.Point(4, 22)
        Me.tpCur.Name = "tpCur"
        Me.tpCur.Size = New System.Drawing.Size(723, 231)
        Me.tpCur.TabIndex = 1
        Me.tpCur.Text = "Foreign Trade"
        Me.tpCur.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Enabled = False
        Me.Button4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.Location = New System.Drawing.Point(254, 20)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(24, 23)
        Me.Button4.TabIndex = 30
        Me.Button4.Text = "..."
        Me.Button4.UseVisualStyleBackColor = True
        Me.Button4.Visible = False
        '
        'CurCodeTextBox
        '
        Me.CurCodeTextBox.Location = New System.Drawing.Point(154, 21)
        Me.CurCodeTextBox.Name = "CurCodeTextBox"
        Me.CurCodeTextBox.Size = New System.Drawing.Size(100, 21)
        Me.CurCodeTextBox.TabIndex = 27
        '
        'CurExRateTextBox
        '
        Me.CurExRateTextBox.Location = New System.Drawing.Point(153, 51)
        Me.CurExRateTextBox.Name = "CurExRateTextBox"
        Me.CurExRateTextBox.Size = New System.Drawing.Size(100, 21)
        Me.CurExRateTextBox.TabIndex = 29
        '
        'CurCodeLabel1
        '
        CurCodeLabel1.AutoSize = True
        CurCodeLabel1.Location = New System.Drawing.Point(6, 24)
        CurCodeLabel1.Name = "CurCodeLabel1"
        CurCodeLabel1.Size = New System.Drawing.Size(142, 13)
        CurCodeLabel1.TabIndex = 26
        CurCodeLabel1.Text = "Currency Code . . . . . . . . ."
        '
        'CurExRateLabel1
        '
        CurExRateLabel1.AutoSize = True
        CurExRateLabel1.Location = New System.Drawing.Point(6, 54)
        CurExRateLabel1.Name = "CurExRateLabel1"
        CurExRateLabel1.Size = New System.Drawing.Size(141, 13)
        CurExRateLabel1.TabIndex = 28
        CurExRateLabel1.Text = "Currency Exchange Rate . ."
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnPrint)
        Me.Panel1.Controls.Add(Me.btnVoid)
        Me.Panel1.Controls.Add(Me.btnModify)
        Me.Panel1.Controls.Add(Me.btnCancel)
        Me.Panel1.Controls.Add(Me.btnSave)
        Me.Panel1.Controls.Add(Me.btnNew)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 414)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(731, 35)
        Me.Panel1.TabIndex = 9
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPrint.Location = New System.Drawing.Point(634, 5)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(93, 23)
        Me.btnPrint.TabIndex = 5
        Me.btnPrint.Text = "Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'btnVoid
        '
        Me.btnVoid.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnVoid.Location = New System.Drawing.Point(436, 6)
        Me.btnVoid.Name = "btnVoid"
        Me.btnVoid.Size = New System.Drawing.Size(93, 23)
        Me.btnVoid.TabIndex = 4
        Me.btnVoid.Text = "Delete"
        Me.btnVoid.UseVisualStyleBackColor = True
        '
        'btnModify
        '
        Me.btnModify.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnModify.Location = New System.Drawing.Point(139, 6)
        Me.btnModify.Name = "btnModify"
        Me.btnModify.Size = New System.Drawing.Size(93, 23)
        Me.btnModify.TabIndex = 3
        Me.btnModify.Text = "Modify"
        Me.btnModify.UseVisualStyleBackColor = True
        Me.btnModify.Visible = False
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(535, 6)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(93, 23)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(337, 6)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(93, 23)
        Me.btnSave.TabIndex = 1
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnNew.Location = New System.Drawing.Point(238, 6)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(93, 23)
        Me.btnNew.TabIndex = 0
        Me.btnNew.Text = "New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'dgvOrdItem
        '
        Me.dgvOrdItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvOrdItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ItemNo, Me.VariantName, Me.ItemName, Me.WareHouse, Me.UOM, Me.Quantity, Me.Price, Me.DisPr, Me.DisPer, Me.Amount, Me.BaseUOM, Me.BasePrice, Me.FOC, Me.DeliveredDate, Me.Promotion, Me.Offer, Me.Priority, Me.QtyShip, Me.Explode, Me.Discount})
        Me.dgvOrdItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvOrdItem.Location = New System.Drawing.Point(0, 0)
        Me.dgvOrdItem.MultiSelect = False
        Me.dgvOrdItem.Name = "dgvOrdItem"
        Me.dgvOrdItem.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvOrdItem.Size = New System.Drawing.Size(731, 157)
        Me.dgvOrdItem.TabIndex = 10
        '
        'ItemNo
        '
        Me.ItemNo.HeaderText = "Item No."
        Me.ItemNo.Name = "ItemNo"
        '
        'VariantName
        '
        Me.VariantName.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.[Nothing]
        Me.VariantName.HeaderText = "Variant Name"
        Me.VariantName.Name = "VariantName"
        Me.VariantName.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.VariantName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'ItemName
        '
        Me.ItemName.HeaderText = "Item Name"
        Me.ItemName.Name = "ItemName"
        Me.ItemName.Width = 200
        '
        'WareHouse
        '
        Me.WareHouse.HeaderText = "Warehouse"
        Me.WareHouse.Name = "WareHouse"
        Me.WareHouse.ReadOnly = True
        Me.WareHouse.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'UOM
        '
        Me.UOM.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.[Nothing]
        Me.UOM.HeaderText = "UOM"
        Me.UOM.Items.AddRange(New Object() {"a", "b"})
        Me.UOM.Name = "UOM"
        '
        'Quantity
        '
        Me.Quantity.HeaderText = "Quantity"
        Me.Quantity.Name = "Quantity"
        '
        'Price
        '
        Me.Price.HeaderText = "Price"
        Me.Price.Name = "Price"
        Me.Price.ReadOnly = True
        '
        'DisPr
        '
        Me.DisPr.HeaderText = "Discount Price"
        Me.DisPr.Name = "DisPr"
        '
        'DisPer
        '
        Me.DisPer.HeaderText = "Discount in %"
        Me.DisPer.Name = "DisPer"
        '
        'Amount
        '
        Me.Amount.HeaderText = "Amount"
        Me.Amount.Name = "Amount"
        Me.Amount.ReadOnly = True
        '
        'BaseUOM
        '
        Me.BaseUOM.HeaderText = "BaseUOM"
        Me.BaseUOM.Name = "BaseUOM"
        Me.BaseUOM.Visible = False
        '
        'BasePrice
        '
        Me.BasePrice.HeaderText = "BasePrice"
        Me.BasePrice.Name = "BasePrice"
        Me.BasePrice.Visible = False
        '
        'FOC
        '
        Me.FOC.HeaderText = "FOC"
        Me.FOC.Name = "FOC"
        Me.FOC.Visible = False
        '
        'DeliveredDate
        '
        Me.DeliveredDate.HeaderText = "Delivery Date"
        Me.DeliveredDate.Name = "DeliveredDate"
        Me.DeliveredDate.ReadOnly = True
        '
        'Promotion
        '
        Me.Promotion.HeaderText = "Promotion"
        Me.Promotion.Name = "Promotion"
        Me.Promotion.Visible = False
        '
        'Offer
        '
        Me.Offer.HeaderText = "Offer"
        Me.Offer.Name = "Offer"
        Me.Offer.Visible = False
        '
        'Priority
        '
        Me.Priority.HeaderText = "Priority"
        Me.Priority.Name = "Priority"
        Me.Priority.Visible = False
        '
        'QtyShip
        '
        Me.QtyShip.HeaderText = "Qty. Shipped"
        Me.QtyShip.Name = "QtyShip"
        Me.QtyShip.ReadOnly = True
        Me.QtyShip.Visible = False
        Me.QtyShip.Width = 5
        '
        'Explode
        '
        Me.Explode.HeaderText = "Explode"
        Me.Explode.Name = "Explode"
        Me.Explode.Visible = False
        '
        'Discount
        '
        Me.Discount.HeaderText = "Discount"
        Me.Discount.Name = "Discount"
        Me.Discount.Visible = False
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.dgvOrdItem)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(0, 257)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(731, 157)
        Me.Panel3.TabIndex = 37
        '
        'PO
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(731, 449)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.TabControl2)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "PO"
        Me.Text = "PO"
        Me.TabControl2.ResumeLayout(False)
        Me.tpInvoice.ResumeLayout(False)
        Me.tpInvoice.PerformLayout()
        Me.gbAmount.ResumeLayout(False)
        Me.gbAmount.PerformLayout()
        Me.tpCur.ResumeLayout(False)
        Me.tpCur.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        CType(Me.dgvOrdItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabControl2 As System.Windows.Forms.TabControl
    Friend WithEvents tpInvoice As System.Windows.Forms.TabPage
    Friend WithEvents btnPay As System.Windows.Forms.Button
    Friend WithEvents btnAgent As System.Windows.Forms.Button
    Friend WithEvents btnCustNo As System.Windows.Forms.Button
    Friend WithEvents gbAmount As System.Windows.Forms.GroupBox
    Friend WithEvents TotalAmtLabel As System.Windows.Forms.Label
    Friend WithEvents GSTAmtLabel As System.Windows.Forms.Label
    Friend WithEvents SubTotalLabel As System.Windows.Forms.Label
    Friend WithEvents DiscountLabel As System.Windows.Forms.Label
    Friend WithEvents CustNameTextBox As System.Windows.Forms.TextBox
    Friend WithEvents InvDtDateTimePicker As System.Windows.Forms.DateTimePicker
    Friend WithEvents CustIdTextBox As System.Windows.Forms.TextBox
    Friend WithEvents PONoTextBox As System.Windows.Forms.TextBox
    Friend WithEvents AgentIDTextBox As System.Windows.Forms.TextBox
    Friend WithEvents PayTermsTextBox As System.Windows.Forms.TextBox
    Friend WithEvents tpCur As System.Windows.Forms.TabPage
    Friend WithEvents CurCodeTextBox As System.Windows.Forms.TextBox
    Friend WithEvents CurExRateTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents lblVoid As System.Windows.Forms.Label
    Friend WithEvents btnVoid As System.Windows.Forms.Button
    Friend WithEvents btnModify As System.Windows.Forms.Button
    Friend WithEvents dgvOrdItem As System.Windows.Forms.DataGridView
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents InvNoTextBox As System.Windows.Forms.TextBox
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents ItemNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents VariantName As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ItemName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents WareHouse As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents UOM As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Quantity As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Price As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DisPr As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DisPer As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Amount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BaseUOM As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BasePrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FOC As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DeliveredDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Promotion As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Offer As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Priority As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents QtyShip As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Explode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Discount As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
