<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ReasonMaster
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
        Me.dgvExchange = New System.Windows.Forms.DataGridView
        Me.PnlButton = New System.Windows.Forms.Panel
        Me.btnSave = New System.Windows.Forms.Button
        Me.ItemNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ItemName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Qy = New System.Windows.Forms.DataGridViewComboBoxColumn
        CType(Me.dgvExchange, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PnlButton.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvExchange
        '
        Me.dgvExchange.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvExchange.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ItemNo, Me.ItemName, Me.Qy})
        Me.dgvExchange.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvExchange.Location = New System.Drawing.Point(0, 0)
        Me.dgvExchange.Name = "dgvExchange"
        Me.dgvExchange.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvExchange.Size = New System.Drawing.Size(596, 300)
        Me.dgvExchange.TabIndex = 96
        '
        'PnlButton
        '
        Me.PnlButton.Controls.Add(Me.btnSave)
        Me.PnlButton.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PnlButton.Location = New System.Drawing.Point(0, 300)
        Me.PnlButton.Name = "PnlButton"
        Me.PnlButton.Size = New System.Drawing.Size(596, 35)
        Me.PnlButton.TabIndex = 95
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(498, 6)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(93, 23)
        Me.btnSave.TabIndex = 76
        Me.btnSave.Text = "&Save"
        Me.btnSave.UseVisualStyleBackColor = True
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
        Me.ItemName.Width = 250
        '
        'Qy
        '
        Me.Qy.HeaderText = "Reason Type"
        Me.Qy.Name = "Qy"
        Me.Qy.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Qy.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.Qy.Width = 200
        '
        'ReasonMaster
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(596, 335)
        Me.Controls.Add(Me.dgvExchange)
        Me.Controls.Add(Me.PnlButton)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "ReasonMaster"
        Me.Text = "Reason Master"
        CType(Me.dgvExchange, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PnlButton.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvExchange As System.Windows.Forms.DataGridView
    Friend WithEvents PnlButton As System.Windows.Forms.Panel
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents ItemNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ItemName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Qy As System.Windows.Forms.DataGridViewComboBoxColumn
End Class
