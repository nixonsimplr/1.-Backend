<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ConsignmentQty
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
        Me.dgvItems = New System.Windows.Forms.DataGridView
        Me.ItemNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ItemName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.UOM = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PnlButton = New System.Windows.Forms.Panel
        Me.btnSave = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.DgvCustName = New System.Windows.Forms.DataGridView
        Me.CustNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CustName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnCust = New System.Windows.Forms.Button
        Me.txtCustName = New System.Windows.Forms.TextBox
        Me.txtCustNo = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        CType(Me.dgvItems, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PnlButton.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.DgvCustName, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvItems
        '
        Me.dgvItems.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvItems.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ItemNo, Me.ItemName, Me.UOM, Me.Qty})
        Me.dgvItems.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItems.Location = New System.Drawing.Point(0, 133)
        Me.dgvItems.Name = "dgvItems"
        Me.dgvItems.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvItems.Size = New System.Drawing.Size(516, 242)
        Me.dgvItems.TabIndex = 100
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
        Me.ItemName.Width = 250
        '
        'UOM
        '
        Me.UOM.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.[Nothing]
        Me.UOM.HeaderText = "UOM"
        Me.UOM.Name = "UOM"
        Me.UOM.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.UOM.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'Qty
        '
        Me.Qty.HeaderText = "Qty"
        Me.Qty.Name = "Qty"
        '
        'PnlButton
        '
        Me.PnlButton.Controls.Add(Me.btnSave)
        Me.PnlButton.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PnlButton.Location = New System.Drawing.Point(0, 340)
        Me.PnlButton.Name = "PnlButton"
        Me.PnlButton.Size = New System.Drawing.Size(516, 35)
        Me.PnlButton.TabIndex = 90
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(413, 8)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(93, 23)
        Me.btnSave.TabIndex = 76
        Me.btnSave.Text = "&Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.DgvCustName)
        Me.Panel1.Controls.Add(Me.btnCust)
        Me.Panel1.Controls.Add(Me.txtCustName)
        Me.Panel1.Controls.Add(Me.txtCustNo)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(516, 133)
        Me.Panel1.TabIndex = 91
        '
        'DgvCustName
        '
        Me.DgvCustName.AllowDrop = True
        Me.DgvCustName.AllowUserToAddRows = False
        Me.DgvCustName.AllowUserToDeleteRows = False
        Me.DgvCustName.AllowUserToResizeColumns = False
        Me.DgvCustName.AllowUserToResizeRows = False
        Me.DgvCustName.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgvCustName.ColumnHeadersVisible = False
        Me.DgvCustName.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CustNo, Me.CustName})
        Me.DgvCustName.Location = New System.Drawing.Point(224, 68)
        Me.DgvCustName.MultiSelect = False
        Me.DgvCustName.Name = "DgvCustName"
        Me.DgvCustName.RowHeadersVisible = False
        Me.DgvCustName.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.DgvCustName.Size = New System.Drawing.Size(220, 65)
        Me.DgvCustName.TabIndex = 39
        Me.DgvCustName.Visible = False
        '
        'CustNo
        '
        Me.CustNo.HeaderText = "CustomerNo"
        Me.CustNo.Name = "CustNo"
        Me.CustNo.Visible = False
        '
        'CustName
        '
        Me.CustName.HeaderText = "Customer Name"
        Me.CustName.Name = "CustName"
        Me.CustName.ReadOnly = True
        Me.CustName.Width = 250
        '
        'btnCust
        '
        Me.btnCust.Location = New System.Drawing.Point(385, 15)
        Me.btnCust.Name = "btnCust"
        Me.btnCust.Size = New System.Drawing.Size(28, 24)
        Me.btnCust.TabIndex = 4
        Me.btnCust.Text = "..."
        Me.btnCust.UseVisualStyleBackColor = True
        '
        'txtCustName
        '
        Me.txtCustName.Location = New System.Drawing.Point(224, 45)
        Me.txtCustName.Name = "txtCustName"
        Me.txtCustName.Size = New System.Drawing.Size(220, 21)
        Me.txtCustName.TabIndex = 3
        '
        'txtCustNo
        '
        Me.txtCustNo.Location = New System.Drawing.Point(224, 16)
        Me.txtCustNo.Name = "txtCustNo"
        Me.txtCustNo.Size = New System.Drawing.Size(158, 21)
        Me.txtCustNo.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 53)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(247, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Customer Name . . . . . . . . . . . . . . . . . . . . . . . "
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(233, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Customer No . . . . . . . . . . . . . . . . . . . . . . . "
        '
        'ConsignmentQty
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(516, 375)
        Me.Controls.Add(Me.dgvItems)
        Me.Controls.Add(Me.PnlButton)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "ConsignmentQty"
        Me.Text = "Consignment Qty"
        CType(Me.dgvItems, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PnlButton.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.DgvCustName, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvItems As System.Windows.Forms.DataGridView
    Friend WithEvents PnlButton As System.Windows.Forms.Panel
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents ItemNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ItemName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents UOM As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents txtCustName As System.Windows.Forms.TextBox
    Friend WithEvents txtCustNo As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnCust As System.Windows.Forms.Button
    Friend WithEvents DgvCustName As System.Windows.Forms.DataGridView
    Friend WithEvents CustNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CustName As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
