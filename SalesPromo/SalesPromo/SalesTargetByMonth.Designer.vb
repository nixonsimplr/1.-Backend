<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SalesTargetByMonth
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.btnSave = New System.Windows.Forms.Button
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.dgvTarget = New System.Windows.Forms.DataGridView
        Me.Month = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TargetAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.AchievedAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PnlButton = New System.Windows.Forms.Panel
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.cmbAgent = New System.Windows.Forms.ComboBox
        lblSalePersCode = New System.Windows.Forms.Label
        Me.Panel2.SuspendLayout()
        CType(Me.dgvTarget, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PnlButton.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(572, 12)
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
        Me.Panel2.Size = New System.Drawing.Size(675, 375)
        Me.Panel2.TabIndex = 91
        '
        'dgvTarget
        '
        Me.dgvTarget.AllowUserToAddRows = False
        Me.dgvTarget.AllowUserToDeleteRows = False
        Me.dgvTarget.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvTarget.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Month, Me.TargetAmount, Me.AchievedAmount})
        Me.dgvTarget.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvTarget.Location = New System.Drawing.Point(0, 0)
        Me.dgvTarget.MultiSelect = False
        Me.dgvTarget.Name = "dgvTarget"
        Me.dgvTarget.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvTarget.Size = New System.Drawing.Size(675, 375)
        Me.dgvTarget.TabIndex = 0
        '
        'Month
        '
        Me.Month.HeaderText = "Month"
        Me.Month.Name = "Month"
        Me.Month.ReadOnly = True
        Me.Month.Width = 150
        '
        'TargetAmount
        '
        DataGridViewCellStyle1.Format = "N2"
        DataGridViewCellStyle1.NullValue = "0.00"
        Me.TargetAmount.DefaultCellStyle = DataGridViewCellStyle1
        Me.TargetAmount.HeaderText = "Target Amount($)"
        Me.TargetAmount.Name = "TargetAmount"
        Me.TargetAmount.Width = 150
        '
        'AchievedAmount
        '
        Me.AchievedAmount.HeaderText = "Achieved Amount($)"
        Me.AchievedAmount.Name = "AchievedAmount"
        Me.AchievedAmount.ReadOnly = True
        Me.AchievedAmount.Width = 150
        '
        'PnlButton
        '
        Me.PnlButton.Controls.Add(Me.btnSave)
        Me.PnlButton.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PnlButton.Location = New System.Drawing.Point(0, 392)
        Me.PnlButton.Name = "PnlButton"
        Me.PnlButton.Size = New System.Drawing.Size(675, 39)
        Me.PnlButton.TabIndex = 89
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cmbAgent)
        Me.Panel1.Controls.Add(lblSalePersCode)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(675, 56)
        Me.Panel1.TabIndex = 90
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
        'lblSalePersCode
        '
        lblSalePersCode.AutoSize = True
        lblSalePersCode.Location = New System.Drawing.Point(3, 18)
        lblSalePersCode.Name = "lblSalePersCode"
        lblSalePersCode.Size = New System.Drawing.Size(156, 13)
        lblSalePersCode.TabIndex = 67
        lblSalePersCode.Text = "Sales Person Code. . . . . . . . ."
        '
        'SalesTargetByMonth
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(675, 431)
        Me.Controls.Add(Me.PnlButton)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "SalesTargetByMonth"
        Me.Text = "Sales Target By Month"
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
    Friend WithEvents cmbAgent As System.Windows.Forms.ComboBox
    Friend WithEvents Month As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TargetAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents AchievedAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents lblSalePersCode As System.Windows.Forms.Label
End Class
