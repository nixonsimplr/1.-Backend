<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CommMaster
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
        Me.GoodsChkIn_btnSave = New System.Windows.Forms.Button
        Me.dgvExchange = New System.Windows.Forms.DataGridView
        Me.ItemNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ItemName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Qy = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.Out = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PnlButton = New System.Windows.Forms.Panel
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.btnAllAgent = New System.Windows.Forms.Button
        Me.btnAllItem = New System.Windows.Forms.Button
        Me.btnAllCust = New System.Windows.Forms.Button
        CType(Me.dgvExchange, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PnlButton.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GoodsChkIn_btnSave
        '
        Me.GoodsChkIn_btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GoodsChkIn_btnSave.Location = New System.Drawing.Point(341, 6)
        Me.GoodsChkIn_btnSave.Name = "GoodsChkIn_btnSave"
        Me.GoodsChkIn_btnSave.Size = New System.Drawing.Size(93, 23)
        Me.GoodsChkIn_btnSave.TabIndex = 76
        Me.GoodsChkIn_btnSave.Text = "&Save"
        Me.GoodsChkIn_btnSave.UseVisualStyleBackColor = True
        '
        'dgvExchange
        '
        Me.dgvExchange.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvExchange.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ItemNo, Me.ItemName, Me.Qy, Me.Out})
        Me.dgvExchange.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvExchange.Location = New System.Drawing.Point(0, 60)
        Me.dgvExchange.Name = "dgvExchange"
        Me.dgvExchange.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvExchange.Size = New System.Drawing.Size(446, 207)
        Me.dgvExchange.TabIndex = 94
        '
        'ItemNo
        '
        Me.ItemNo.HeaderText = "Code"
        Me.ItemNo.Name = "ItemNo"
        '
        'ItemName
        '
        Me.ItemName.HeaderText = "Description"
        Me.ItemName.Name = "ItemName"
        '
        'Qy
        '
        Me.Qy.HeaderText = "Commission Type"
        Me.Qy.Items.AddRange(New Object() {"Customer", "Sales Agent", "Product"})
        Me.Qy.Name = "Qy"
        Me.Qy.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Qy.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.Qy.Width = 200
        '
        'Out
        '
        Me.Out.HeaderText = "IN"
        Me.Out.Name = "Out"
        Me.Out.Visible = False
        '
        'PnlButton
        '
        Me.PnlButton.Controls.Add(Me.GoodsChkIn_btnSave)
        Me.PnlButton.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PnlButton.Location = New System.Drawing.Point(0, 267)
        Me.PnlButton.Name = "PnlButton"
        Me.PnlButton.Size = New System.Drawing.Size(446, 35)
        Me.PnlButton.TabIndex = 93
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnAllCust)
        Me.Panel1.Controls.Add(Me.btnAllItem)
        Me.Panel1.Controls.Add(Me.btnAllAgent)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(446, 60)
        Me.Panel1.TabIndex = 95
        '
        'btnAllAgent
        '
        Me.btnAllAgent.Location = New System.Drawing.Point(15, 8)
        Me.btnAllAgent.Name = "btnAllAgent"
        Me.btnAllAgent.Size = New System.Drawing.Size(118, 42)
        Me.btnAllAgent.TabIndex = 0
        Me.btnAllAgent.Text = "Load All Agent"
        Me.btnAllAgent.UseVisualStyleBackColor = True
        '
        'btnAllItem
        '
        Me.btnAllItem.Location = New System.Drawing.Point(161, 8)
        Me.btnAllItem.Name = "btnAllItem"
        Me.btnAllItem.Size = New System.Drawing.Size(118, 42)
        Me.btnAllItem.TabIndex = 1
        Me.btnAllItem.Text = "Load All Item"
        Me.btnAllItem.UseVisualStyleBackColor = True
        '
        'btnAllCust
        '
        Me.btnAllCust.Location = New System.Drawing.Point(307, 8)
        Me.btnAllCust.Name = "btnAllCust"
        Me.btnAllCust.Size = New System.Drawing.Size(118, 42)
        Me.btnAllCust.TabIndex = 2
        Me.btnAllCust.Text = "Load All Customer"
        Me.btnAllCust.UseVisualStyleBackColor = True
        '
        'CommMaster
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(446, 302)
        Me.Controls.Add(Me.dgvExchange)
        Me.Controls.Add(Me.PnlButton)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "CommMaster"
        Me.Text = "Commission Master"
        CType(Me.dgvExchange, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PnlButton.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GoodsChkIn_btnSave As System.Windows.Forms.Button
    Friend WithEvents dgvExchange As System.Windows.Forms.DataGridView
    Friend WithEvents PnlButton As System.Windows.Forms.Panel
    Friend WithEvents ItemNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ItemName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Qy As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Out As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnAllCust As System.Windows.Forms.Button
    Friend WithEvents btnAllItem As System.Windows.Forms.Button
    Friend WithEvents btnAllAgent As System.Windows.Forms.Button
End Class
