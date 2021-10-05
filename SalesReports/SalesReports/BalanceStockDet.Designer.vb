<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BalanceStockDet
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
        Dim Label3 As System.Windows.Forms.Label
        Dim Label6 As System.Windows.Forms.Label
        Me.Check = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.CustNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.dgvOrdItem = New System.Windows.Forms.DataGridView
        Me.CustName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ShipName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.AgentID = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Barcode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CompanyName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label1 = New System.Windows.Forms.Label
        Me.cbCategory = New System.Windows.Forms.ComboBox
        Me.pnlButton = New System.Windows.Forms.Panel
        Me.chkSelAll = New System.Windows.Forms.CheckBox
        Me.btnView = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.dtpToDate = New System.Windows.Forms.DateTimePicker
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.btnApply = New System.Windows.Forms.Button
        Me.pnlTop = New System.Windows.Forms.Panel
        Label3 = New System.Windows.Forms.Label
        Label6 = New System.Windows.Forms.Label
        CType(Me.dgvOrdItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlButton.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.pnlTop.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label3
        '
        Label3.AutoSize = True
        Label3.Location = New System.Drawing.Point(8, 22)
        Label3.Name = "Label3"
        Label3.Size = New System.Drawing.Size(127, 13)
        Label3.TabIndex = 116
        Label3.Text = "Category . . . . . . . . . . . . ."
        '
        'Label6
        '
        Label6.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Label6.AutoSize = True
        Label6.Location = New System.Drawing.Point(8, 50)
        Label6.Name = "Label6"
        Label6.Size = New System.Drawing.Size(163, 13)
        Label6.TabIndex = 131
        Label6.Text = "Start Date. . . . . . . . . . . . . . . . . . "
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
        Me.dgvOrdItem.Size = New System.Drawing.Size(676, 265)
        Me.dgvOrdItem.TabIndex = 138
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
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 80)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(160, 13)
        Me.Label1.TabIndex = 134
        Me.Label1.Text = "End Date. . . . . . . . . . . . . . . . . . "
        '
        'cbCategory
        '
        Me.cbCategory.FormattingEnabled = True
        Me.cbCategory.Location = New System.Drawing.Point(110, 17)
        Me.cbCategory.Name = "cbCategory"
        Me.cbCategory.Size = New System.Drawing.Size(229, 21)
        Me.cbCategory.TabIndex = 117
        '
        'pnlButton
        '
        Me.pnlButton.Controls.Add(Me.chkSelAll)
        Me.pnlButton.Controls.Add(Me.btnView)
        Me.pnlButton.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlButton.Location = New System.Drawing.Point(0, 390)
        Me.pnlButton.Name = "pnlButton"
        Me.pnlButton.Size = New System.Drawing.Size(676, 37)
        Me.pnlButton.TabIndex = 136
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
        'btnView
        '
        Me.btnView.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnView.Location = New System.Drawing.Point(579, 7)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(93, 23)
        Me.btnView.TabIndex = 0
        Me.btnView.Text = "&View"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cbCategory)
        Me.GroupBox2.Controls.Add(Label3)
        Me.GroupBox2.Controls.Add(Me.dtpToDate)
        Me.GroupBox2.Controls.Add(Me.dtpDate)
        Me.GroupBox2.Controls.Add(Label6)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 9)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(345, 109)
        Me.GroupBox2.TabIndex = 130
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Selection"
        '
        'dtpToDate
        '
        Me.dtpToDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpToDate.CustomFormat = "MMMM yyyy"
        Me.dtpToDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpToDate.Location = New System.Drawing.Point(110, 77)
        Me.dtpToDate.Name = "dtpToDate"
        Me.dtpToDate.Size = New System.Drawing.Size(229, 20)
        Me.dtpToDate.TabIndex = 133
        '
        'dtpDate
        '
        Me.dtpDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDate.CustomFormat = "MMMM yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpDate.Location = New System.Drawing.Point(110, 46)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(229, 20)
        Me.dtpDate.TabIndex = 130
        '
        'btnApply
        '
        Me.btnApply.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnApply.Location = New System.Drawing.Point(359, 94)
        Me.btnApply.Name = "btnApply"
        Me.btnApply.Size = New System.Drawing.Size(93, 23)
        Me.btnApply.TabIndex = 1
        Me.btnApply.Text = "&Refresh"
        Me.btnApply.UseVisualStyleBackColor = True
        '
        'pnlTop
        '
        Me.pnlTop.Controls.Add(Me.GroupBox2)
        Me.pnlTop.Controls.Add(Me.btnApply)
        Me.pnlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlTop.Name = "pnlTop"
        Me.pnlTop.Size = New System.Drawing.Size(676, 125)
        Me.pnlTop.TabIndex = 137
        '
        'BalanceStockDet
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(676, 427)
        Me.Controls.Add(Me.dgvOrdItem)
        Me.Controls.Add(Me.pnlButton)
        Me.Controls.Add(Me.pnlTop)
        Me.Name = "BalanceStockDet"
        Me.Text = "Balance Stock Detail"
        CType(Me.dgvOrdItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlButton.ResumeLayout(False)
        Me.pnlButton.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.pnlTop.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Check As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents CustNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgvOrdItem As System.Windows.Forms.DataGridView
    Friend WithEvents CustName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ShipName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents AgentID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Barcode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CompanyName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cbCategory As System.Windows.Forms.ComboBox
    Friend WithEvents pnlButton As System.Windows.Forms.Panel
    Friend WithEvents chkSelAll As System.Windows.Forms.CheckBox
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents dtpToDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnApply As System.Windows.Forms.Button
    Friend WithEvents pnlTop As System.Windows.Forms.Panel
End Class
