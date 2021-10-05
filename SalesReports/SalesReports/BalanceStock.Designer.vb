<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BalanceStock
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
        Me.dgvOrdItem = New System.Windows.Forms.DataGridView
        Me.CustNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CustName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ShipName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.AgentID = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Barcode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CompanyName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnApply = New System.Windows.Forms.Button
        Me.cbCategory = New System.Windows.Forms.ComboBox
        Me.pnlTop = New System.Windows.Forms.Panel
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.btnView = New System.Windows.Forms.Button
        Me.pnlButton = New System.Windows.Forms.Panel
        Me.chkSelAll = New System.Windows.Forms.CheckBox
        Label3 = New System.Windows.Forms.Label
        Label6 = New System.Windows.Forms.Label
        CType(Me.dgvOrdItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlTop.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.pnlButton.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label3
        '
        Label3.AutoSize = True
        Label3.Location = New System.Drawing.Point(7, 23)
        Label3.Name = "Label3"
        Label3.Size = New System.Drawing.Size(143, 13)
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
        Label6.Text = "Period. . . . . . . . . . . . . . . . . . "
        '
        'Check
        '
        Me.Check.HeaderText = ""
        Me.Check.Name = "Check"
        Me.Check.Width = 30
        '
        'dgvOrdItem
        '
        Me.dgvOrdItem.AllowUserToAddRows = False
        Me.dgvOrdItem.AllowUserToDeleteRows = False
        Me.dgvOrdItem.AllowUserToOrderColumns = True
        Me.dgvOrdItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvOrdItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Check, Me.CustNo, Me.CustName, Me.ShipName, Me.AgentID, Me.Barcode, Me.CompanyName})
        Me.dgvOrdItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvOrdItem.Location = New System.Drawing.Point(0, 103)
        Me.dgvOrdItem.Name = "dgvOrdItem"
        Me.dgvOrdItem.Size = New System.Drawing.Size(677, 258)
        Me.dgvOrdItem.TabIndex = 134
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
        'btnApply
        '
        Me.btnApply.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnApply.Location = New System.Drawing.Point(358, 67)
        Me.btnApply.Name = "btnApply"
        Me.btnApply.Size = New System.Drawing.Size(93, 23)
        Me.btnApply.TabIndex = 1
        Me.btnApply.Text = "&Refresh"
        Me.btnApply.UseVisualStyleBackColor = True
        '
        'cbCategory
        '
        Me.cbCategory.FormattingEnabled = True
        Me.cbCategory.Location = New System.Drawing.Point(110, 19)
        Me.cbCategory.Name = "cbCategory"
        Me.cbCategory.Size = New System.Drawing.Size(229, 21)
        Me.cbCategory.TabIndex = 117
        '
        'pnlTop
        '
        Me.pnlTop.Controls.Add(Me.GroupBox2)
        Me.pnlTop.Controls.Add(Me.btnApply)
        Me.pnlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlTop.Name = "pnlTop"
        Me.pnlTop.Size = New System.Drawing.Size(677, 103)
        Me.pnlTop.TabIndex = 133
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cbCategory)
        Me.GroupBox2.Controls.Add(Label3)
        Me.GroupBox2.Controls.Add(Me.dtpDate)
        Me.GroupBox2.Controls.Add(Label6)
        Me.GroupBox2.Location = New System.Drawing.Point(7, 13)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(345, 77)
        Me.GroupBox2.TabIndex = 130
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Selection"
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
        'btnView
        '
        Me.btnView.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnView.Location = New System.Drawing.Point(580, 7)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(93, 23)
        Me.btnView.TabIndex = 0
        Me.btnView.Text = "&View"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'pnlButton
        '
        Me.pnlButton.Controls.Add(Me.chkSelAll)
        Me.pnlButton.Controls.Add(Me.btnView)
        Me.pnlButton.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlButton.Location = New System.Drawing.Point(0, 361)
        Me.pnlButton.Name = "pnlButton"
        Me.pnlButton.Size = New System.Drawing.Size(677, 37)
        Me.pnlButton.TabIndex = 132
        '
        'chkSelAll
        '
        Me.chkSelAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkSelAll.AutoSize = True
        Me.chkSelAll.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSelAll.Location = New System.Drawing.Point(4, 9)
        Me.chkSelAll.Name = "chkSelAll"
        Me.chkSelAll.Size = New System.Drawing.Size(84, 17)
        Me.chkSelAll.TabIndex = 108
        Me.chkSelAll.Text = "Select ALL"
        Me.chkSelAll.UseVisualStyleBackColor = True
        '
        'BalanceStock
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(677, 398)
        Me.Controls.Add(Me.dgvOrdItem)
        Me.Controls.Add(Me.pnlTop)
        Me.Controls.Add(Me.pnlButton)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "BalanceStock"
        Me.Text = "Balance Stock"
        CType(Me.dgvOrdItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlTop.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.pnlButton.ResumeLayout(False)
        Me.pnlButton.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Check As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents dgvOrdItem As System.Windows.Forms.DataGridView
    Friend WithEvents CustNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CustName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ShipName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents AgentID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Barcode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CompanyName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnApply As System.Windows.Forms.Button
    Friend WithEvents cbCategory As System.Windows.Forms.ComboBox
    Friend WithEvents pnlTop As System.Windows.Forms.Panel
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents pnlButton As System.Windows.Forms.Panel
    Friend WithEvents chkSelAll As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
End Class
