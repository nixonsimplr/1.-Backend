<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class LoyaltyByItem
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
        Dim lblCustPrice As System.Windows.Forms.Label
        Me.txtCode = New System.Windows.Forms.TextBox
        Me.btnCustNo = New System.Windows.Forms.Button
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.dgvCustomerPrice = New System.Windows.Forms.DataGridView
        Me.ItemNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ItemName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.UOM = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.Qy = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.PnlButton = New System.Windows.Forms.Panel
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        lblCustPrice = New System.Windows.Forms.Label
        Me.Panel2.SuspendLayout()
        CType(Me.dgvCustomerPrice, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.PnlButton.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblCustPrice
        '
        lblCustPrice.AutoSize = True
        lblCustPrice.Location = New System.Drawing.Point(5, 21)
        lblCustPrice.Name = "lblCustPrice"
        lblCustPrice.Size = New System.Drawing.Size(125, 13)
        lblCustPrice.TabIndex = 72
        lblCustPrice.Text = "Member Type. . . . . . . ."
        '
        'txtCode
        '
        Me.txtCode.Location = New System.Drawing.Point(125, 18)
        Me.txtCode.Name = "txtCode"
        Me.txtCode.Size = New System.Drawing.Size(100, 21)
        Me.txtCode.TabIndex = 73
        '
        'btnCustNo
        '
        Me.btnCustNo.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCustNo.Location = New System.Drawing.Point(225, 17)
        Me.btnCustNo.Name = "btnCustNo"
        Me.btnCustNo.Size = New System.Drawing.Size(24, 23)
        Me.btnCustNo.TabIndex = 76
        Me.btnCustNo.Text = "..."
        Me.btnCustNo.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.dgvCustomerPrice)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 55)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(563, 278)
        Me.Panel2.TabIndex = 97
        '
        'dgvCustomerPrice
        '
        Me.dgvCustomerPrice.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvCustomerPrice.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ItemNo, Me.ItemName, Me.UOM, Me.Qy})
        Me.dgvCustomerPrice.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvCustomerPrice.Location = New System.Drawing.Point(0, 0)
        Me.dgvCustomerPrice.MultiSelect = False
        Me.dgvCustomerPrice.Name = "dgvCustomerPrice"
        Me.dgvCustomerPrice.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvCustomerPrice.Size = New System.Drawing.Size(563, 278)
        Me.dgvCustomerPrice.TabIndex = 1
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
        Me.Qy.HeaderText = "Points"
        Me.Qy.Name = "Qy"
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(268, 6)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(93, 23)
        Me.btnSave.TabIndex = 76
        Me.btnSave.Text = "&Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnNew.Location = New System.Drawing.Point(169, 6)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(93, 23)
        Me.btnNew.TabIndex = 75
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.txtCode)
        Me.Panel1.Controls.Add(Me.btnCustNo)
        Me.Panel1.Controls.Add(lblCustPrice)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(563, 55)
        Me.Panel1.TabIndex = 96
        '
        'PnlButton
        '
        Me.PnlButton.Controls.Add(Me.btnDelete)
        Me.PnlButton.Controls.Add(Me.btnSave)
        Me.PnlButton.Controls.Add(Me.btnCancel)
        Me.PnlButton.Controls.Add(Me.btnNew)
        Me.PnlButton.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PnlButton.Location = New System.Drawing.Point(0, 333)
        Me.PnlButton.Name = "PnlButton"
        Me.PnlButton.Size = New System.Drawing.Size(563, 35)
        Me.PnlButton.TabIndex = 95
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDelete.Location = New System.Drawing.Point(367, 6)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(93, 23)
        Me.btnDelete.TabIndex = 79
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(466, 6)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(93, 23)
        Me.btnCancel.TabIndex = 77
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'LoyaltyByItem
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(563, 368)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.PnlButton)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "LoyaltyByItem"
        Me.Text = "Loyalty By Item"
        Me.Panel2.ResumeLayout(False)
        CType(Me.dgvCustomerPrice, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.PnlButton.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents txtCode As System.Windows.Forms.TextBox
    Friend WithEvents btnCustNo As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents dgvCustomerPrice As System.Windows.Forms.DataGridView
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents PnlButton As System.Windows.Forms.Panel
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents ItemNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ItemName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents UOM As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Qy As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
