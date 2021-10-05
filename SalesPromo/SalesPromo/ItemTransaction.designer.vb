<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ItemTransaction
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
        Me.btnSave = New System.Windows.Forms.Button
        Me.dtpDocumentDate = New System.Windows.Forms.DateTimePicker
        Me.txtDocNo = New System.Windows.Forms.TextBox
        Me.dgvItemTrans = New System.Windows.Forms.DataGridView
        Me.ItemID = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ItemName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BaseUOM = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnCancel = New System.Windows.Forms.Button
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.PnlButton = New System.Windows.Forms.Panel
        Me.btnModify = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnDocNo = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.btnLocation = New System.Windows.Forms.Button
        Me.txtLocation = New System.Windows.Forms.TextBox
        Me.cmbType = New System.Windows.Forms.ComboBox
        Item_lblDocDate = New System.Windows.Forms.Label
        Item_lblDocNo = New System.Windows.Forms.Label
        Item_lblType = New System.Windows.Forms.Label
        Item_lblLocation = New System.Windows.Forms.Label
        CType(Me.dgvItemTrans, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.PnlButton.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Item_lblDocDate
        '
        Item_lblDocDate.AutoSize = True
        Item_lblDocDate.Location = New System.Drawing.Point(7, 42)
        Item_lblDocDate.Name = "Item_lblDocDate"
        Item_lblDocDate.Size = New System.Drawing.Size(144, 13)
        Item_lblDocDate.TabIndex = 64
        Item_lblDocDate.Text = "Document Date . . . . . . . . ."
        '
        'Item_lblDocNo
        '
        Item_lblDocNo.AutoSize = True
        Item_lblDocNo.Location = New System.Drawing.Point(7, 12)
        Item_lblDocNo.Name = "Item_lblDocNo"
        Item_lblDocNo.Size = New System.Drawing.Size(148, 13)
        Item_lblDocNo.TabIndex = 62
        Item_lblDocNo.Text = "Document No . . . . . . . . . . ."
        '
        'Item_lblType
        '
        Item_lblType.AutoSize = True
        Item_lblType.Location = New System.Drawing.Point(7, 73)
        Item_lblType.Name = "Item_lblType"
        Item_lblType.Size = New System.Drawing.Size(140, 13)
        Item_lblType.TabIndex = 65
        Item_lblType.Text = "Type. . . . . . . . . . . . . . . ."
        '
        'Item_lblLocation
        '
        Item_lblLocation.AutoSize = True
        Item_lblLocation.Location = New System.Drawing.Point(6, 102)
        Item_lblLocation.Name = "Item_lblLocation"
        Item_lblLocation.Size = New System.Drawing.Size(145, 13)
        Item_lblLocation.TabIndex = 93
        Item_lblLocation.Text = "Location . . . . . . . . . . . . . ."
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(257, 9)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(93, 23)
        Me.btnSave.TabIndex = 76
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'dtpDocumentDate
        '
        Me.dtpDocumentDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpDocumentDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDocumentDate.Location = New System.Drawing.Point(156, 38)
        Me.dtpDocumentDate.Name = "dtpDocumentDate"
        Me.dtpDocumentDate.Size = New System.Drawing.Size(101, 21)
        Me.dtpDocumentDate.TabIndex = 72
        '
        'txtDocNo
        '
        Me.txtDocNo.Location = New System.Drawing.Point(156, 9)
        Me.txtDocNo.Name = "txtDocNo"
        Me.txtDocNo.Size = New System.Drawing.Size(100, 21)
        Me.txtDocNo.TabIndex = 63
        '
        'dgvItemTrans
        '
        Me.dgvItemTrans.AllowUserToAddRows = False
        Me.dgvItemTrans.AllowUserToDeleteRows = False
        Me.dgvItemTrans.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvItemTrans.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ItemID, Me.ItemName, Me.BaseUOM, Me.Qty})
        Me.dgvItemTrans.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItemTrans.Location = New System.Drawing.Point(0, 0)
        Me.dgvItemTrans.Name = "dgvItemTrans"
        Me.dgvItemTrans.ReadOnly = True
        Me.dgvItemTrans.Size = New System.Drawing.Size(453, 167)
        Me.dgvItemTrans.TabIndex = 74
        '
        'ItemID
        '
        Me.ItemID.HeaderText = "Item No"
        Me.ItemID.Name = "ItemID"
        Me.ItemID.ReadOnly = True
        '
        'ItemName
        '
        Me.ItemName.HeaderText = "Item Name"
        Me.ItemName.Name = "ItemName"
        Me.ItemName.ReadOnly = True
        '
        'BaseUOM
        '
        Me.BaseUOM.HeaderText = "UOM"
        Me.BaseUOM.Name = "BaseUOM"
        Me.BaseUOM.ReadOnly = True
        '
        'Qty
        '
        Me.Qty.HeaderText = "Quantity"
        Me.Qty.Name = "Qty"
        Me.Qty.ReadOnly = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(355, 9)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(93, 23)
        Me.btnCancel.TabIndex = 77
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.dgvItemTrans)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 136)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(453, 167)
        Me.Panel2.TabIndex = 91
        '
        'PnlButton
        '
        Me.PnlButton.Controls.Add(Me.btnSave)
        Me.PnlButton.Controls.Add(Me.btnCancel)
        Me.PnlButton.Controls.Add(Me.btnModify)
        Me.PnlButton.Controls.Add(Me.btnNew)
        Me.PnlButton.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PnlButton.Location = New System.Drawing.Point(0, 303)
        Me.PnlButton.Name = "PnlButton"
        Me.PnlButton.Size = New System.Drawing.Size(453, 38)
        Me.PnlButton.TabIndex = 89
        '
        'btnModify
        '
        Me.btnModify.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnModify.Location = New System.Drawing.Point(159, 9)
        Me.btnModify.Name = "btnModify"
        Me.btnModify.Size = New System.Drawing.Size(93, 23)
        Me.btnModify.TabIndex = 78
        Me.btnModify.Text = "Modify Exchange"
        Me.btnModify.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnNew.Location = New System.Drawing.Point(61, 9)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(93, 23)
        Me.btnNew.TabIndex = 75
        Me.btnNew.Text = "New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnDocNo
        '
        Me.btnDocNo.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDocNo.Location = New System.Drawing.Point(256, 8)
        Me.btnDocNo.Name = "btnDocNo"
        Me.btnDocNo.Size = New System.Drawing.Size(24, 23)
        Me.btnDocNo.TabIndex = 73
        Me.btnDocNo.Text = "..."
        Me.btnDocNo.UseVisualStyleBackColor = True
        Me.btnDocNo.Visible = False
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnLocation)
        Me.Panel1.Controls.Add(Me.txtLocation)
        Me.Panel1.Controls.Add(Item_lblLocation)
        Me.Panel1.Controls.Add(Me.cmbType)
        Me.Panel1.Controls.Add(Me.btnDocNo)
        Me.Panel1.Controls.Add(Me.dtpDocumentDate)
        Me.Panel1.Controls.Add(Me.txtDocNo)
        Me.Panel1.Controls.Add(Item_lblDocDate)
        Me.Panel1.Controls.Add(Item_lblType)
        Me.Panel1.Controls.Add(Item_lblDocNo)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(453, 136)
        Me.Panel1.TabIndex = 90
        '
        'btnLocation
        '
        Me.btnLocation.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnLocation.Location = New System.Drawing.Point(257, 98)
        Me.btnLocation.Name = "btnLocation"
        Me.btnLocation.Size = New System.Drawing.Size(24, 23)
        Me.btnLocation.TabIndex = 95
        Me.btnLocation.Text = "..."
        Me.btnLocation.UseVisualStyleBackColor = True
        Me.btnLocation.Visible = False
        '
        'txtLocation
        '
        Me.txtLocation.Location = New System.Drawing.Point(157, 99)
        Me.txtLocation.Name = "txtLocation"
        Me.txtLocation.Size = New System.Drawing.Size(100, 21)
        Me.txtLocation.TabIndex = 94
        '
        'cmbType
        '
        Me.cmbType.FormattingEnabled = True
        Me.cmbType.Items.AddRange(New Object() {"IN", "OUT", "TRANSFER"})
        Me.cmbType.Location = New System.Drawing.Point(156, 70)
        Me.cmbType.Name = "cmbType"
        Me.cmbType.Size = New System.Drawing.Size(101, 21)
        Me.cmbType.TabIndex = 92
        '
        'ItemTransaction
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(453, 341)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.PnlButton)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "ItemTransaction"
        Me.Text = "Item Transaction"
        CType(Me.dgvItemTrans, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.PnlButton.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents dtpDocumentDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtDocNo As System.Windows.Forms.TextBox
    Friend WithEvents dgvItemTrans As System.Windows.Forms.DataGridView
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents PnlButton As System.Windows.Forms.Panel
    Friend WithEvents btnModify As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnDocNo As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents cmbType As System.Windows.Forms.ComboBox
    Friend WithEvents btnLocation As System.Windows.Forms.Button
    Friend WithEvents txtLocation As System.Windows.Forms.TextBox
    Friend WithEvents ItemID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ItemName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BaseUOM As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Item_lblDocDate As System.Windows.Forms.Label
    Friend WithEvents Item_lblDocNo As System.Windows.Forms.Label
    Friend WithEvents Item_lblType As System.Windows.Forms.Label
    Friend WithEvents Item_lblLocation As System.Windows.Forms.Label
End Class
