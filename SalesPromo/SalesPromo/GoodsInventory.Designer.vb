<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class GoodsInventory
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
        Me.dgvCustomerPrice = New System.Windows.Forms.DataGridView
        Me.Location = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ItemNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ItemName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.UOM = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.Qy = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PnlButton = New System.Windows.Forms.Panel
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        CType(Me.dgvCustomerPrice, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PnlButton.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvCustomerPrice
        '
        Me.dgvCustomerPrice.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvCustomerPrice.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Location, Me.ItemNo, Me.ItemName, Me.UOM, Me.Qy})
        Me.dgvCustomerPrice.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvCustomerPrice.Location = New System.Drawing.Point(0, 0)
        Me.dgvCustomerPrice.MultiSelect = False
        Me.dgvCustomerPrice.Name = "dgvCustomerPrice"
        Me.dgvCustomerPrice.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvCustomerPrice.Size = New System.Drawing.Size(644, 406)
        Me.dgvCustomerPrice.TabIndex = 2
        '
        'Location
        '
        Me.Location.HeaderText = "Location"
        Me.Location.Name = "Location"
        '
        'ItemNo
        '
        Me.ItemNo.HeaderText = "Item No"
        Me.ItemNo.Name = "ItemNo"
        Me.ItemNo.ReadOnly = True
        '
        'ItemName
        '
        Me.ItemName.HeaderText = "Item Name"
        Me.ItemName.Name = "ItemName"
        Me.ItemName.ReadOnly = True
        Me.ItemName.Width = 200
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
        Me.Qy.HeaderText = "Qty"
        Me.Qy.Name = "Qy"
        '
        'PnlButton
        '
        Me.PnlButton.Controls.Add(Me.btnDelete)
        Me.PnlButton.Controls.Add(Me.btnSave)
        Me.PnlButton.Controls.Add(Me.btnCancel)
        Me.PnlButton.Controls.Add(Me.btnNew)
        Me.PnlButton.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PnlButton.Location = New System.Drawing.Point(0, 371)
        Me.PnlButton.Name = "PnlButton"
        Me.PnlButton.Size = New System.Drawing.Size(644, 35)
        Me.PnlButton.TabIndex = 96
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDelete.Location = New System.Drawing.Point(448, 6)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(93, 23)
        Me.btnDelete.TabIndex = 79
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(349, 6)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(93, 23)
        Me.btnSave.TabIndex = 76
        Me.btnSave.Text = "&Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(547, 6)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(93, 23)
        Me.btnCancel.TabIndex = 77
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnNew.Location = New System.Drawing.Point(250, 6)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(93, 23)
        Me.btnNew.TabIndex = 75
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'GoodsInventory
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(644, 406)
        Me.Controls.Add(Me.PnlButton)
        Me.Controls.Add(Me.dgvCustomerPrice)
        Me.Name = "GoodsInventory"
        Me.Text = "GoodsInventory"
        CType(Me.dgvCustomerPrice, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PnlButton.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvCustomerPrice As System.Windows.Forms.DataGridView
    Friend WithEvents Location As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ItemNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ItemName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents UOM As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Qy As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PnlButton As System.Windows.Forms.Panel
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
End Class
