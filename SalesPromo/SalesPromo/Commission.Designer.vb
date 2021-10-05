<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Commission
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
        Dim Item_lblLocation As System.Windows.Forms.Label
        Dim Item_lblType As System.Windows.Forms.Label
        Me.PnlButton = New System.Windows.Forms.Panel
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.dgvCommission = New System.Windows.Forms.DataGridView
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.rbPercentage = New System.Windows.Forms.RadioButton
        Me.rbPrice = New System.Windows.Forms.RadioButton
        Me.cmbAgent = New System.Windows.Forms.ComboBox
        Me.btnRefresh = New System.Windows.Forms.Button
        Item_lblLocation = New System.Windows.Forms.Label
        Item_lblType = New System.Windows.Forms.Label
        Me.PnlButton.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.dgvCommission, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Item_lblLocation
        '
        Item_lblLocation.AutoSize = True
        Item_lblLocation.Location = New System.Drawing.Point(3, 18)
        Item_lblLocation.Name = "Item_lblLocation"
        Item_lblLocation.Size = New System.Drawing.Size(162, 13)
        Item_lblLocation.TabIndex = 93
        Item_lblLocation.Text = "Sales Agent . . . . . . . . . . . . . ."
        '
        'Item_lblType
        '
        Item_lblType.AutoSize = True
        Item_lblType.Location = New System.Drawing.Point(3, 48)
        Item_lblType.Name = "Item_lblType"
        Item_lblType.Size = New System.Drawing.Size(112, 13)
        Item_lblType.TabIndex = 65
        Item_lblType.Text = "Type. . . . . . . . . . . ."
        '
        'PnlButton
        '
        Me.PnlButton.Controls.Add(Me.btnSave)
        Me.PnlButton.Controls.Add(Me.btnCancel)
        Me.PnlButton.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PnlButton.Location = New System.Drawing.Point(0, 376)
        Me.PnlButton.Name = "PnlButton"
        Me.PnlButton.Size = New System.Drawing.Size(638, 38)
        Me.PnlButton.TabIndex = 92
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(442, 9)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(93, 23)
        Me.btnSave.TabIndex = 76
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(540, 9)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(93, 23)
        Me.btnCancel.TabIndex = 77
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.dgvCommission)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 78)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(638, 298)
        Me.Panel2.TabIndex = 94
        '
        'dgvCommission
        '
        Me.dgvCommission.AllowUserToAddRows = False
        Me.dgvCommission.AllowUserToDeleteRows = False
        Me.dgvCommission.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvCommission.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvCommission.Location = New System.Drawing.Point(0, 0)
        Me.dgvCommission.Name = "dgvCommission"
        Me.dgvCommission.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.dgvCommission.Size = New System.Drawing.Size(638, 298)
        Me.dgvCommission.TabIndex = 74
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnRefresh)
        Me.Panel1.Controls.Add(Me.rbPercentage)
        Me.Panel1.Controls.Add(Me.rbPrice)
        Me.Panel1.Controls.Add(Me.cmbAgent)
        Me.Panel1.Controls.Add(Item_lblType)
        Me.Panel1.Controls.Add(Item_lblLocation)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(638, 78)
        Me.Panel1.TabIndex = 93
        '
        'rbPercentage
        '
        Me.rbPercentage.AutoSize = True
        Me.rbPercentage.Location = New System.Drawing.Point(182, 48)
        Me.rbPercentage.Name = "rbPercentage"
        Me.rbPercentage.Size = New System.Drawing.Size(80, 17)
        Me.rbPercentage.TabIndex = 95
        Me.rbPercentage.TabStop = True
        Me.rbPercentage.Text = "Percentage"
        Me.rbPercentage.UseVisualStyleBackColor = True
        '
        'rbPrice
        '
        Me.rbPrice.AutoSize = True
        Me.rbPrice.Location = New System.Drawing.Point(114, 48)
        Me.rbPrice.Name = "rbPrice"
        Me.rbPrice.Size = New System.Drawing.Size(48, 17)
        Me.rbPrice.TabIndex = 94
        Me.rbPrice.TabStop = True
        Me.rbPrice.Text = "Price"
        Me.rbPrice.UseVisualStyleBackColor = True
        '
        'cmbAgent
        '
        Me.cmbAgent.FormattingEnabled = True
        Me.cmbAgent.Location = New System.Drawing.Point(113, 15)
        Me.cmbAgent.Name = "cmbAgent"
        Me.cmbAgent.Size = New System.Drawing.Size(139, 21)
        Me.cmbAgent.TabIndex = 92
        '
        'btnRefresh
        '
        Me.btnRefresh.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnRefresh.Location = New System.Drawing.Point(551, 45)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(75, 23)
        Me.btnRefresh.TabIndex = 96
        Me.btnRefresh.Text = "Refresh"
        Me.btnRefresh.UseVisualStyleBackColor = True
        '
        'Commission
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(638, 414)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.PnlButton)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "Commission"
        Me.Text = "Commission"
        Me.PnlButton.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        CType(Me.dgvCommission, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents PnlButton As System.Windows.Forms.Panel
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents dgvCommission As System.Windows.Forms.DataGridView
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents rbPercentage As System.Windows.Forms.RadioButton
    Friend WithEvents rbPrice As System.Windows.Forms.RadioButton
    Friend WithEvents cmbAgent As System.Windows.Forms.ComboBox
    Friend WithEvents btnRefresh As System.Windows.Forms.Button
End Class
