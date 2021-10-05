<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Barcodes
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
        Dim lblCustPrice As System.Windows.Forms.Label
        Dim Label1 As System.Windows.Forms.Label
        Dim Label2 As System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.dgvCustomerPrice = New System.Windows.Forms.DataGridView
        Me.Barcode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.cmbUOM = New System.Windows.Forms.ComboBox
        Me.txtCode = New System.Windows.Forms.TextBox
        Me.btnItemNo = New System.Windows.Forms.Button
        Me.txtDescription = New System.Windows.Forms.TextBox
        Me.PnlButton = New System.Windows.Forms.Panel
        Me.btnSave = New System.Windows.Forms.Button
        lblCustPrice = New System.Windows.Forms.Label
        Label1 = New System.Windows.Forms.Label
        Label2 = New System.Windows.Forms.Label
        Me.Panel2.SuspendLayout()
        CType(Me.dgvCustomerPrice, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.PnlButton.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblCustPrice
        '
        lblCustPrice.AutoSize = True
        lblCustPrice.Location = New System.Drawing.Point(5, 15)
        lblCustPrice.Name = "lblCustPrice"
        lblCustPrice.Size = New System.Drawing.Size(120, 13)
        lblCustPrice.TabIndex = 72
        lblCustPrice.Text = "Code. . . . . . . . . . . . ."
        '
        'Label1
        '
        Label1.AutoSize = True
        Label1.Location = New System.Drawing.Point(5, 44)
        Label1.Name = "Label1"
        Label1.Size = New System.Drawing.Size(141, 13)
        Label1.TabIndex = 74
        Label1.Text = "Description. . . . . . . . . . . ."
        '
        'Label2
        '
        Label2.AutoSize = True
        Label2.Location = New System.Drawing.Point(5, 74)
        Label2.Name = "Label2"
        Label2.Size = New System.Drawing.Size(146, 13)
        Label2.TabIndex = 77
        Label2.Text = "UOM. . . . . . . . . . . . . . . . ."
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.dgvCustomerPrice)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 99)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(528, 259)
        Me.Panel2.TabIndex = 97
        '
        'dgvCustomerPrice
        '
        Me.dgvCustomerPrice.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvCustomerPrice.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Barcode})
        Me.dgvCustomerPrice.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvCustomerPrice.Location = New System.Drawing.Point(0, 0)
        Me.dgvCustomerPrice.MultiSelect = False
        Me.dgvCustomerPrice.Name = "dgvCustomerPrice"
        Me.dgvCustomerPrice.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvCustomerPrice.Size = New System.Drawing.Size(528, 259)
        Me.dgvCustomerPrice.TabIndex = 1
        '
        'Barcode
        '
        Me.Barcode.HeaderText = "Barcodes"
        Me.Barcode.Name = "Barcode"
        Me.Barcode.Width = 300
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cmbUOM)
        Me.Panel1.Controls.Add(Label2)
        Me.Panel1.Controls.Add(Me.txtCode)
        Me.Panel1.Controls.Add(Me.btnItemNo)
        Me.Panel1.Controls.Add(Me.txtDescription)
        Me.Panel1.Controls.Add(lblCustPrice)
        Me.Panel1.Controls.Add(Label1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(528, 99)
        Me.Panel1.TabIndex = 96
        '
        'cmbUOM
        '
        Me.cmbUOM.FormattingEnabled = True
        Me.cmbUOM.Location = New System.Drawing.Point(125, 70)
        Me.cmbUOM.Name = "cmbUOM"
        Me.cmbUOM.Size = New System.Drawing.Size(124, 21)
        Me.cmbUOM.TabIndex = 78
        '
        'txtCode
        '
        Me.txtCode.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtCode.Location = New System.Drawing.Point(125, 12)
        Me.txtCode.Name = "txtCode"
        Me.txtCode.Size = New System.Drawing.Size(100, 21)
        Me.txtCode.TabIndex = 73
        '
        'btnItemNo
        '
        Me.btnItemNo.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnItemNo.Location = New System.Drawing.Point(225, 11)
        Me.btnItemNo.Name = "btnItemNo"
        Me.btnItemNo.Size = New System.Drawing.Size(24, 23)
        Me.btnItemNo.TabIndex = 76
        Me.btnItemNo.Text = "..."
        Me.btnItemNo.UseVisualStyleBackColor = True
        '
        'txtDescription
        '
        Me.txtDescription.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtDescription.Location = New System.Drawing.Point(125, 41)
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(255, 21)
        Me.txtDescription.TabIndex = 75
        '
        'PnlButton
        '
        Me.PnlButton.Controls.Add(Me.btnSave)
        Me.PnlButton.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PnlButton.Location = New System.Drawing.Point(0, 358)
        Me.PnlButton.Name = "PnlButton"
        Me.PnlButton.Size = New System.Drawing.Size(528, 35)
        Me.PnlButton.TabIndex = 95
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(423, 6)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(93, 23)
        Me.btnSave.TabIndex = 76
        Me.btnSave.Text = "&Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'Barcodes
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(528, 393)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.PnlButton)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "Barcodes"
        Me.Text = "Barcodes"
        Me.Panel2.ResumeLayout(False)
        CType(Me.dgvCustomerPrice, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.PnlButton.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents dgvCustomerPrice As System.Windows.Forms.DataGridView
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents txtCode As System.Windows.Forms.TextBox
    Friend WithEvents btnItemNo As System.Windows.Forms.Button
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents PnlButton As System.Windows.Forms.Panel
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents cmbUOM As System.Windows.Forms.ComboBox
    Friend WithEvents Barcode As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
