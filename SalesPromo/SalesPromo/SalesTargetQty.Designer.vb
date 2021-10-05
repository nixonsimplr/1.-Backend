<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SalesTargetQty
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
        Dim lblSalePersCode As System.Windows.Forms.Label
        Dim Label1 As System.Windows.Forms.Label
        Me.btnSave = New System.Windows.Forms.Button
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.dgvTarget = New System.Windows.Forms.DataGridView
        Me.Item = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ItemNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ItemName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.UOM = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.Qy = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PnlButton = New System.Windows.Forms.Panel
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.cmbMonth = New System.Windows.Forms.ComboBox
        Me.cmbAgent = New System.Windows.Forms.ComboBox
        Me.chkSelAll = New System.Windows.Forms.CheckBox
        lblSalePersCode = New System.Windows.Forms.Label
        Label1 = New System.Windows.Forms.Label
        Me.Panel2.SuspendLayout()
        CType(Me.dgvTarget, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PnlButton.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblSalePersCode
        '
        lblSalePersCode.AutoSize = True
        lblSalePersCode.Location = New System.Drawing.Point(3, 18)
        lblSalePersCode.Name = "lblSalePersCode"
        lblSalePersCode.Size = New System.Drawing.Size(156, 13)
        lblSalePersCode.TabIndex = 67
        lblSalePersCode.Text = "Sales Person Code. . . . . . . . ."
        '
        'Label1
        '
        Label1.AutoSize = True
        Label1.Location = New System.Drawing.Point(3, 49)
        Label1.Name = "Label1"
        Label1.Size = New System.Drawing.Size(162, 13)
        Label1.TabIndex = 68
        Label1.Text = "Select the month. . . . . . . . . . ."
        Label1.Visible = False
        AddHandler Label1.Click, AddressOf Me.Label1_Click_1
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(572, 8)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(93, 23)
        Me.btnSave.TabIndex = 76
        Me.btnSave.Text = "&Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.dgvTarget)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 56)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(675, 340)
        Me.Panel2.TabIndex = 91
        '
        'dgvTarget
        '
        Me.dgvTarget.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvTarget.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Item, Me.ItemNo, Me.ItemName, Me.UOM, Me.Qy})
        Me.dgvTarget.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvTarget.Location = New System.Drawing.Point(0, 0)
        Me.dgvTarget.MultiSelect = False
        Me.dgvTarget.Name = "dgvTarget"
        Me.dgvTarget.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvTarget.Size = New System.Drawing.Size(675, 340)
        Me.dgvTarget.TabIndex = 0
        '
        'Item
        '
        Me.Item.HeaderText = ""
        Me.Item.Name = "Item"
        Me.Item.Width = 30
        '
        'ItemNo
        '
        Me.ItemNo.HeaderText = "ItemNo"
        Me.ItemNo.Name = "ItemNo"
        '
        'ItemName
        '
        Me.ItemName.HeaderText = "Item Name"
        Me.ItemName.Name = "ItemName"
        Me.ItemName.Width = 300
        '
        'UOM
        '
        Me.UOM.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.[Nothing]
        Me.UOM.HeaderText = "UOM"
        Me.UOM.Name = "UOM"
        Me.UOM.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.UOM.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'Qy
        '
        Me.Qy.HeaderText = "Quantity"
        Me.Qy.Name = "Qy"
        '
        'PnlButton
        '
        Me.PnlButton.Controls.Add(Me.btnSave)
        Me.PnlButton.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PnlButton.Location = New System.Drawing.Point(0, 396)
        Me.PnlButton.Name = "PnlButton"
        Me.PnlButton.Size = New System.Drawing.Size(675, 35)
        Me.PnlButton.TabIndex = 89
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.chkSelAll)
        Me.Panel1.Controls.Add(Me.cmbMonth)
        Me.Panel1.Controls.Add(Me.cmbAgent)
        Me.Panel1.Controls.Add(Label1)
        Me.Panel1.Controls.Add(lblSalePersCode)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(675, 56)
        Me.Panel1.TabIndex = 90
        '
        'cmbMonth
        '
        Me.cmbMonth.FormattingEnabled = True
        Me.cmbMonth.Items.AddRange(New Object() {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November" & Global.Microsoft.VisualBasic.ChrW(9), "December"})
        Me.cmbMonth.Location = New System.Drawing.Point(151, 41)
        Me.cmbMonth.Name = "cmbMonth"
        Me.cmbMonth.Size = New System.Drawing.Size(163, 21)
        Me.cmbMonth.TabIndex = 98
        Me.cmbMonth.Visible = False
        '
        'cmbAgent
        '
        Me.cmbAgent.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAgent.FormattingEnabled = True
        Me.cmbAgent.Location = New System.Drawing.Point(151, 12)
        Me.cmbAgent.Name = "cmbAgent"
        Me.cmbAgent.Size = New System.Drawing.Size(289, 21)
        Me.cmbAgent.TabIndex = 97
        '
        'chkSelAll
        '
        Me.chkSelAll.AutoSize = True
        Me.chkSelAll.Location = New System.Drawing.Point(5, 40)
        Me.chkSelAll.Name = "chkSelAll"
        Me.chkSelAll.Size = New System.Drawing.Size(75, 17)
        Me.chkSelAll.TabIndex = 99
        Me.chkSelAll.Text = "Select ALL"
        Me.chkSelAll.UseVisualStyleBackColor = True
        '
        'SalesTargetQty
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(675, 431)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.PnlButton)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "SalesTargetQty"
        Me.Text = "Sales Target in Qty"
        Me.Panel2.ResumeLayout(False)
        CType(Me.dgvTarget, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PnlButton.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents dgvTarget As System.Windows.Forms.DataGridView
    Friend WithEvents PnlButton As System.Windows.Forms.Panel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents cmbMonth As System.Windows.Forms.ComboBox
    Friend WithEvents cmbAgent As System.Windows.Forms.ComboBox
    Friend WithEvents Item As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ItemNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ItemName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents UOM As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Qy As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents chkSelAll As System.Windows.Forms.CheckBox

End Class
