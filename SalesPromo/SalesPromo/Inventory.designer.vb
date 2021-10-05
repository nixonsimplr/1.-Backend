<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Inventory
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
        Dim Inv_lblLocation As System.Windows.Forms.Label
        Dim Label1 As System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.dgvInventory = New System.Windows.Forms.DataGridView
        Me.btnCancel = New System.Windows.Forms.Button
        Me.PnlButton = New System.Windows.Forms.Panel
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnModify = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnLocation = New System.Windows.Forms.Button
        Me.txtLocation = New System.Windows.Forms.TextBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.cmbLocation = New System.Windows.Forms.ComboBox
        Me.txtLocationName = New System.Windows.Forms.TextBox
        Inv_lblLocation = New System.Windows.Forms.Label
        Label1 = New System.Windows.Forms.Label
        Me.Panel2.SuspendLayout()
        CType(Me.dgvInventory, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PnlButton.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Inv_lblLocation
        '
        Inv_lblLocation.AutoSize = True
        Inv_lblLocation.Location = New System.Drawing.Point(5, 15)
        Inv_lblLocation.Name = "Inv_lblLocation"
        Inv_lblLocation.Size = New System.Drawing.Size(170, 13)
        Inv_lblLocation.TabIndex = 93
        Inv_lblLocation.Text = "Location Code. . . . . . . . . . . . . ."
        '
        'Label1
        '
        Label1.AutoSize = True
        Label1.Location = New System.Drawing.Point(5, 46)
        Label1.Name = "Label1"
        Label1.Size = New System.Drawing.Size(172, 13)
        Label1.TabIndex = 123
        Label1.Text = "Location Name. . . . . . . . . . . . . ."
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.dgvInventory)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 72)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(588, 435)
        Me.Panel2.TabIndex = 94
        '
        'dgvInventory
        '
        Me.dgvInventory.AllowUserToAddRows = False
        Me.dgvInventory.AllowUserToDeleteRows = False
        Me.dgvInventory.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvInventory.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvInventory.Location = New System.Drawing.Point(0, 0)
        Me.dgvInventory.Name = "dgvInventory"
        Me.dgvInventory.ReadOnly = True
        Me.dgvInventory.Size = New System.Drawing.Size(588, 435)
        Me.dgvInventory.TabIndex = 74
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(490, 9)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(93, 23)
        Me.btnCancel.TabIndex = 77
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'PnlButton
        '
        Me.PnlButton.Controls.Add(Me.btnSave)
        Me.PnlButton.Controls.Add(Me.btnCancel)
        Me.PnlButton.Controls.Add(Me.btnModify)
        Me.PnlButton.Controls.Add(Me.btnNew)
        Me.PnlButton.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PnlButton.Location = New System.Drawing.Point(0, 507)
        Me.PnlButton.Name = "PnlButton"
        Me.PnlButton.Size = New System.Drawing.Size(588, 38)
        Me.PnlButton.TabIndex = 92
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(392, 9)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(93, 23)
        Me.btnSave.TabIndex = 76
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnModify
        '
        Me.btnModify.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnModify.Location = New System.Drawing.Point(294, 9)
        Me.btnModify.Name = "btnModify"
        Me.btnModify.Size = New System.Drawing.Size(93, 23)
        Me.btnModify.TabIndex = 78
        Me.btnModify.Text = "Modify Exchange"
        Me.btnModify.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnNew.Location = New System.Drawing.Point(196, 9)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(93, 23)
        Me.btnNew.TabIndex = 75
        Me.btnNew.Text = "New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnLocation
        '
        Me.btnLocation.Enabled = False
        Me.btnLocation.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnLocation.Location = New System.Drawing.Point(378, 11)
        Me.btnLocation.Name = "btnLocation"
        Me.btnLocation.Size = New System.Drawing.Size(24, 23)
        Me.btnLocation.TabIndex = 95
        Me.btnLocation.Text = "..."
        Me.btnLocation.UseVisualStyleBackColor = True
        Me.btnLocation.Visible = False
        '
        'txtLocation
        '
        Me.txtLocation.Location = New System.Drawing.Point(278, 12)
        Me.txtLocation.Name = "txtLocation"
        Me.txtLocation.Size = New System.Drawing.Size(100, 21)
        Me.txtLocation.TabIndex = 94
        Me.txtLocation.Visible = False
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.txtLocationName)
        Me.Panel1.Controls.Add(Label1)
        Me.Panel1.Controls.Add(Me.cmbLocation)
        Me.Panel1.Controls.Add(Me.btnLocation)
        Me.Panel1.Controls.Add(Me.txtLocation)
        Me.Panel1.Controls.Add(Inv_lblLocation)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(588, 72)
        Me.Panel1.TabIndex = 93
        '
        'cmbLocation
        '
        Me.cmbLocation.FormattingEnabled = True
        Me.cmbLocation.Location = New System.Drawing.Point(133, 11)
        Me.cmbLocation.Name = "cmbLocation"
        Me.cmbLocation.Size = New System.Drawing.Size(156, 21)
        Me.cmbLocation.TabIndex = 122
        '
        'txtLocationName
        '
        Me.txtLocationName.Location = New System.Drawing.Point(133, 43)
        Me.txtLocationName.Name = "txtLocationName"
        Me.txtLocationName.Size = New System.Drawing.Size(352, 21)
        Me.txtLocationName.TabIndex = 124
        '
        'Inventory
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(588, 545)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.PnlButton)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "Inventory"
        Me.Text = "Inventory"
        Me.Panel2.ResumeLayout(False)
        CType(Me.dgvInventory, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PnlButton.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents dgvInventory As System.Windows.Forms.DataGridView
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents PnlButton As System.Windows.Forms.Panel
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnModify As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnLocation As System.Windows.Forms.Button
    Friend WithEvents txtLocation As System.Windows.Forms.TextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents cmbLocation As System.Windows.Forms.ComboBox
    Friend WithEvents txtLocationName As System.Windows.Forms.TextBox
End Class
