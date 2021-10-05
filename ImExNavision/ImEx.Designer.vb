<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Imex
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
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnImpInvoice = New System.Windows.Forms.Button
        Me.lblExportDate = New System.Windows.Forms.Label
        Me.cmbExport = New System.Windows.Forms.ComboBox
        Me.btnIm = New System.Windows.Forms.Button
        Me.btnEx = New System.Windows.Forms.Button
        Me.dgvStatus = New System.Windows.Forms.DataGridView
        Me.Tables = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Status = New System.Windows.Forms.DataGridViewImageColumn
        Me.Panel1.SuspendLayout()
        CType(Me.dgvStatus, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnDelete)
        Me.Panel1.Controls.Add(Me.btnImpInvoice)
        Me.Panel1.Controls.Add(Me.lblExportDate)
        Me.Panel1.Controls.Add(Me.cmbExport)
        Me.Panel1.Controls.Add(Me.btnIm)
        Me.Panel1.Controls.Add(Me.btnEx)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 461)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(629, 29)
        Me.Panel1.TabIndex = 11
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDelete.Location = New System.Drawing.Point(3, 3)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(137, 23)
        Me.btnDelete.TabIndex = 5
        Me.btnDelete.Text = "Delete Tables in Navision"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnImpInvoice
        '
        Me.btnImpInvoice.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnImpInvoice.Location = New System.Drawing.Point(497, 3)
        Me.btnImpInvoice.Name = "btnImpInvoice"
        Me.btnImpInvoice.Size = New System.Drawing.Size(130, 23)
        Me.btnImpInvoice.TabIndex = 4
        Me.btnImpInvoice.Text = "Import Invoice"
        Me.btnImpInvoice.UseVisualStyleBackColor = True
        '
        'lblExportDate
        '
        Me.lblExportDate.AutoSize = True
        Me.lblExportDate.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblExportDate.Location = New System.Drawing.Point(3, 8)
        Me.lblExportDate.Name = "lblExportDate"
        Me.lblExportDate.Size = New System.Drawing.Size(92, 13)
        Me.lblExportDate.TabIndex = 3
        Me.lblExportDate.Text = "Export Date . . ."
        Me.lblExportDate.Visible = False
        '
        'cmbExport
        '
        Me.cmbExport.FormattingEnabled = True
        Me.cmbExport.Location = New System.Drawing.Point(96, 5)
        Me.cmbExport.Name = "cmbExport"
        Me.cmbExport.Size = New System.Drawing.Size(53, 21)
        Me.cmbExport.TabIndex = 2
        Me.cmbExport.Visible = False
        '
        'btnIm
        '
        Me.btnIm.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnIm.Location = New System.Drawing.Point(361, 3)
        Me.btnIm.Name = "btnIm"
        Me.btnIm.Size = New System.Drawing.Size(130, 23)
        Me.btnIm.TabIndex = 1
        Me.btnIm.Text = "Import All"
        Me.btnIm.UseVisualStyleBackColor = True
        '
        'btnEx
        '
        Me.btnEx.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnEx.Location = New System.Drawing.Point(223, 3)
        Me.btnEx.Name = "btnEx"
        Me.btnEx.Size = New System.Drawing.Size(132, 23)
        Me.btnEx.TabIndex = 0
        Me.btnEx.Text = "Export to Navision"
        Me.btnEx.UseVisualStyleBackColor = True
        '
        'dgvStatus
        '
        Me.dgvStatus.AllowUserToAddRows = False
        Me.dgvStatus.AllowUserToDeleteRows = False
        Me.dgvStatus.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvStatus.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvStatus.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Tables, Me.Status})
        Me.dgvStatus.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvStatus.Location = New System.Drawing.Point(0, 0)
        Me.dgvStatus.MultiSelect = False
        Me.dgvStatus.Name = "dgvStatus"
        Me.dgvStatus.ReadOnly = True
        Me.dgvStatus.Size = New System.Drawing.Size(629, 461)
        Me.dgvStatus.TabIndex = 14
        '
        'Tables
        '
        Me.Tables.HeaderText = "Action"
        Me.Tables.Name = "Tables"
        Me.Tables.ReadOnly = True
        '
        'Status
        '
        Me.Status.HeaderText = "Status"
        Me.Status.Name = "Status"
        Me.Status.ReadOnly = True
        '
        'Imex
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(629, 490)
        Me.Controls.Add(Me.dgvStatus)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "Imex"
        Me.Text = "Import / Export"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.dgvStatus, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnIm As System.Windows.Forms.Button
    Friend WithEvents btnEx As System.Windows.Forms.Button
    Friend WithEvents dgvStatus As System.Windows.Forms.DataGridView
    Friend WithEvents cmbExport As System.Windows.Forms.ComboBox
    Friend WithEvents lblExportDate As System.Windows.Forms.Label
    Friend WithEvents Tables As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Status As System.Windows.Forms.DataGridViewImageColumn
    Friend WithEvents btnImpInvoice As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button

End Class
