<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BalanceStockNAV
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim Label6 As System.Windows.Forms.Label
        Dim Label2 As System.Windows.Forms.Label
        Dim Label3 As System.Windows.Forms.Label
        Dim Label4 As System.Windows.Forms.Label
        Me.pnlButton = New System.Windows.Forms.Panel
        Me.btnView = New System.Windows.Forms.Button
        Me.pnlTop = New System.Windows.Forms.Panel
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.cbCategory = New System.Windows.Forms.ComboBox
        Me.cbSupplier = New System.Windows.Forms.ComboBox
        Me.btnApply = New System.Windows.Forms.Button
        Me.cbCompany = New System.Windows.Forms.ComboBox
        Me.dtpToDate = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.dgvOrdItem = New System.Windows.Forms.DataGridView
        Me.Check = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.CustNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CustName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ShipName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.AgentID = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Barcode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CompanyName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.chkSelAll = New System.Windows.Forms.CheckBox
        Label6 = New System.Windows.Forms.Label
        Label2 = New System.Windows.Forms.Label
        Label3 = New System.Windows.Forms.Label
        Label4 = New System.Windows.Forms.Label
        Me.pnlButton.SuspendLayout()
        Me.pnlTop.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvOrdItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnlButton
        '
        Me.pnlButton.Controls.Add(Me.chkSelAll)
        Me.pnlButton.Controls.Add(Me.btnView)
        Me.pnlButton.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlButton.Location = New System.Drawing.Point(0, 369)
        Me.pnlButton.Name = "pnlButton"
        Me.pnlButton.Size = New System.Drawing.Size(830, 37)
        Me.pnlButton.TabIndex = 132
        '
        'btnView
        '
        Me.btnView.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnView.Location = New System.Drawing.Point(733, 7)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(93, 23)
        Me.btnView.TabIndex = 0
        Me.btnView.Text = "&View"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'pnlTop
        '
        Me.pnlTop.Controls.Add(Me.GroupBox2)
        Me.pnlTop.Controls.Add(Me.GroupBox1)
        Me.pnlTop.Controls.Add(Me.btnApply)
        Me.pnlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlTop.Name = "pnlTop"
        Me.pnlTop.Size = New System.Drawing.Size(830, 125)
        Me.pnlTop.TabIndex = 134
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.dtpToDate)
        Me.GroupBox2.Controls.Add(Me.cbCompany)
        Me.GroupBox2.Controls.Add(Me.dtpDate)
        Me.GroupBox2.Controls.Add(Label6)
        Me.GroupBox2.Controls.Add(Label2)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Location = New System.Drawing.Point(7, 13)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(345, 109)
        Me.GroupBox2.TabIndex = 130
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Date Selection"
        '
        'dtpDate
        '
        Me.dtpDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDate.CustomFormat = "MMMM yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpDate.Location = New System.Drawing.Point(110, 46)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(229, 21)
        Me.dtpDate.TabIndex = 130
        '
        'Label6
        '
        Label6.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Label6.AutoSize = True
        Label6.Location = New System.Drawing.Point(6, 50)
        Label6.Name = "Label6"
        Label6.Size = New System.Drawing.Size(183, 13)
        Label6.TabIndex = 131
        Label6.Text = "Start Date. . . . . . . . . . . . . . . . . . "
        '
        'Label2
        '
        Label2.AutoSize = True
        Label2.Location = New System.Drawing.Point(6, 19)
        Label2.Name = "Label2"
        Label2.Size = New System.Drawing.Size(178, 13)
        Label2.TabIndex = 126
        Label2.Text = "Company. . . . . . . . . . . . . . . . . . "
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cbCategory)
        Me.GroupBox1.Controls.Add(Me.cbSupplier)
        Me.GroupBox1.Controls.Add(Label3)
        Me.GroupBox1.Controls.Add(Label4)
        Me.GroupBox1.Location = New System.Drawing.Point(384, 13)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(345, 109)
        Me.GroupBox1.TabIndex = 128
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Item Selection"
        '
        'cbCategory
        '
        Me.cbCategory.FormattingEnabled = True
        Me.cbCategory.Location = New System.Drawing.Point(108, 16)
        Me.cbCategory.Name = "cbCategory"
        Me.cbCategory.Size = New System.Drawing.Size(229, 21)
        Me.cbCategory.TabIndex = 117
        '
        'cbSupplier
        '
        Me.cbSupplier.FormattingEnabled = True
        Me.cbSupplier.Location = New System.Drawing.Point(107, 47)
        Me.cbSupplier.Name = "cbSupplier"
        Me.cbSupplier.Size = New System.Drawing.Size(230, 21)
        Me.cbSupplier.TabIndex = 119
        '
        'Label3
        '
        Label3.AutoSize = True
        Label3.Location = New System.Drawing.Point(6, 24)
        Label3.Name = "Label3"
        Label3.Size = New System.Drawing.Size(143, 13)
        Label3.TabIndex = 116
        Label3.Text = "Category . . . . . . . . . . . . ."
        '
        'Label4
        '
        Label4.AutoSize = True
        Label4.Location = New System.Drawing.Point(6, 52)
        Label4.Name = "Label4"
        Label4.Size = New System.Drawing.Size(164, 13)
        Label4.TabIndex = 118
        Label4.Text = "Supplier Code . . . . . . . . . . . . ."
        '
        'btnApply
        '
        Me.btnApply.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnApply.Location = New System.Drawing.Point(733, 97)
        Me.btnApply.Name = "btnApply"
        Me.btnApply.Size = New System.Drawing.Size(93, 23)
        Me.btnApply.TabIndex = 1
        Me.btnApply.Text = "&Refresh"
        Me.btnApply.UseVisualStyleBackColor = True
        '
        'cbCompany
        '
        Me.cbCompany.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbCompany.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbCompany.FormattingEnabled = True
        Me.cbCompany.Location = New System.Drawing.Point(110, 16)
        Me.cbCompany.Name = "cbCompany"
        Me.cbCompany.Size = New System.Drawing.Size(229, 21)
        Me.cbCompany.TabIndex = 132
        '
        'dtpToDate
        '
        Me.dtpToDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpToDate.CustomFormat = "MMMM yyyy"
        Me.dtpToDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpToDate.Location = New System.Drawing.Point(110, 77)
        Me.dtpToDate.Name = "dtpToDate"
        Me.dtpToDate.Size = New System.Drawing.Size(229, 21)
        Me.dtpToDate.TabIndex = 133
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 80)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(177, 13)
        Me.Label1.TabIndex = 134
        Me.Label1.Text = "End Date. . . . . . . . . . . . . . . . . . "
        '
        'dgvOrdItem
        '
        Me.dgvOrdItem.AllowUserToAddRows = False
        Me.dgvOrdItem.AllowUserToDeleteRows = False
        Me.dgvOrdItem.AllowUserToOrderColumns = True
        Me.dgvOrdItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvOrdItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Check, Me.CustNo, Me.CustName, Me.ShipName, Me.AgentID, Me.Barcode, Me.CompanyName})
        Me.dgvOrdItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvOrdItem.Location = New System.Drawing.Point(0, 125)
        Me.dgvOrdItem.Name = "dgvOrdItem"
        Me.dgvOrdItem.Size = New System.Drawing.Size(830, 244)
        Me.dgvOrdItem.TabIndex = 135
        '
        'Check
        '
        Me.Check.HeaderText = ""
        Me.Check.Name = "Check"
        Me.Check.Width = 30
        '
        'CustNo
        '
        Me.CustNo.HeaderText = "Item No"
        Me.CustNo.Name = "CustNo"
        Me.CustNo.ReadOnly = True
        Me.CustNo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'CustName
        '
        Me.CustName.HeaderText = "Item Name"
        Me.CustName.Name = "CustName"
        Me.CustName.ReadOnly = True
        Me.CustName.Width = 200
        '
        'ShipName
        '
        Me.ShipName.HeaderText = "Base UOM"
        Me.ShipName.Name = "ShipName"
        Me.ShipName.ReadOnly = True
        '
        'AgentID
        '
        Me.AgentID.HeaderText = "Price"
        Me.AgentID.Name = "AgentID"
        Me.AgentID.ReadOnly = True
        '
        'Barcode
        '
        Me.Barcode.HeaderText = "Barcode"
        Me.Barcode.Name = "Barcode"
        Me.Barcode.ReadOnly = True
        '
        'CompanyName
        '
        Me.CompanyName.HeaderText = "Company"
        Me.CompanyName.Name = "CompanyName"
        Me.CompanyName.ReadOnly = True
        Me.CompanyName.Visible = False
        '
        'chkSelAll
        '
        Me.chkSelAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkSelAll.AutoSize = True
        Me.chkSelAll.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSelAll.Location = New System.Drawing.Point(3, 11)
        Me.chkSelAll.Name = "chkSelAll"
        Me.chkSelAll.Size = New System.Drawing.Size(84, 17)
        Me.chkSelAll.TabIndex = 136
        Me.chkSelAll.Text = "Select ALL"
        Me.chkSelAll.UseVisualStyleBackColor = True
        '
        'BalanceStockNAV
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(830, 406)
        Me.Controls.Add(Me.dgvOrdItem)
        Me.Controls.Add(Me.pnlTop)
        Me.Controls.Add(Me.pnlButton)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "BalanceStockNAV"
        Me.Text = "BalanceStockNAV"
        Me.pnlButton.ResumeLayout(False)
        Me.pnlButton.PerformLayout()
        Me.pnlTop.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgvOrdItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlButton As System.Windows.Forms.Panel
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents pnlTop As System.Windows.Forms.Panel
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cbCategory As System.Windows.Forms.ComboBox
    Friend WithEvents cbSupplier As System.Windows.Forms.ComboBox
    Friend WithEvents btnApply As System.Windows.Forms.Button
    Friend WithEvents cbCompany As System.Windows.Forms.ComboBox
    Friend WithEvents dtpToDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dgvOrdItem As System.Windows.Forms.DataGridView
    Friend WithEvents Check As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents CustNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CustName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ShipName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents AgentID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Barcode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CompanyName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents chkSelAll As System.Windows.Forms.CheckBox
End Class
