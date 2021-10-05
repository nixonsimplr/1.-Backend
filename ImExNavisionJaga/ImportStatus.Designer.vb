<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ImportStatus
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
        Dim Label3 As System.Windows.Forms.Label
        Me.btnIm = New System.Windows.Forms.Button
        Me.dgvStatus = New System.Windows.Forms.DataGridView
        Me.Tables = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Status = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.chkPayment = New System.Windows.Forms.CheckBox
        Me.chkOrders = New System.Windows.Forms.CheckBox
        Me.chkSelAll = New System.Windows.Forms.CheckBox
        Me.cmbExport = New System.Windows.Forms.ComboBox
        Me.cmbAgent = New System.Windows.Forms.ComboBox
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnEx = New System.Windows.Forms.Button
        Me.lblExportDate = New System.Windows.Forms.Label
        Label3 = New System.Windows.Forms.Label
        CType(Me.dgvStatus, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label3
        '
        Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Label3.AutoSize = True
        Label3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Label3.Location = New System.Drawing.Point(3, -25)
        Label3.Name = "Label3"
        Label3.Size = New System.Drawing.Size(228, 13)
        Label3.TabIndex = 101
        Label3.Text = "Select the Salesman to Export . . . . . . . . "
        Label3.Visible = False
        '
        'btnIm
        '
        Me.btnIm.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnIm.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnIm.Location = New System.Drawing.Point(448, 4)
        Me.btnIm.Name = "btnIm"
        Me.btnIm.Size = New System.Drawing.Size(201, 42)
        Me.btnIm.TabIndex = 1
        Me.btnIm.Text = "Import from SAP"
        Me.btnIm.UseVisualStyleBackColor = True
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
        Me.dgvStatus.Size = New System.Drawing.Size(652, 366)
        Me.dgvStatus.TabIndex = 16
        '
        'Tables
        '
        Me.Tables.HeaderText = "Tables"
        Me.Tables.Name = "Tables"
        '
        'Status
        '
        Me.Status.HeaderText = "Status"
        Me.Status.Name = "Status"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Controls.Add(Me.chkSelAll)
        Me.Panel1.Controls.Add(Me.cmbExport)
        Me.Panel1.Controls.Add(Me.cmbAgent)
        Me.Panel1.Controls.Add(Label3)
        Me.Panel1.Controls.Add(Me.btnDelete)
        Me.Panel1.Controls.Add(Me.btnIm)
        Me.Panel1.Controls.Add(Me.btnEx)
        Me.Panel1.Controls.Add(Me.lblExportDate)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 366)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(652, 49)
        Me.Panel1.TabIndex = 15
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.chkPayment)
        Me.GroupBox1.Controls.Add(Me.chkOrders)
        Me.GroupBox1.Location = New System.Drawing.Point(13, 58)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(198, 70)
        Me.GroupBox1.TabIndex = 108
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Export to Navision"
        Me.GroupBox1.Visible = False
        '
        'chkPayment
        '
        Me.chkPayment.AutoSize = True
        Me.chkPayment.Location = New System.Drawing.Point(14, 44)
        Me.chkPayment.Name = "chkPayment"
        Me.chkPayment.Size = New System.Drawing.Size(73, 17)
        Me.chkPayment.TabIndex = 1
        Me.chkPayment.Text = "Payments"
        Me.chkPayment.UseVisualStyleBackColor = True
        '
        'chkOrders
        '
        Me.chkOrders.AutoSize = True
        Me.chkOrders.Location = New System.Drawing.Point(14, 21)
        Me.chkOrders.Name = "chkOrders"
        Me.chkOrders.Size = New System.Drawing.Size(59, 17)
        Me.chkOrders.TabIndex = 0
        Me.chkOrders.Text = "Orders"
        Me.chkOrders.UseVisualStyleBackColor = True
        '
        'chkSelAll
        '
        Me.chkSelAll.AutoSize = True
        Me.chkSelAll.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSelAll.Location = New System.Drawing.Point(10, 17)
        Me.chkSelAll.Name = "chkSelAll"
        Me.chkSelAll.Size = New System.Drawing.Size(84, 17)
        Me.chkSelAll.TabIndex = 107
        Me.chkSelAll.Text = "Select ALL"
        Me.chkSelAll.UseVisualStyleBackColor = True
        '
        'cmbExport
        '
        Me.cmbExport.FormattingEnabled = True
        Me.cmbExport.Location = New System.Drawing.Point(207, 31)
        Me.cmbExport.Name = "cmbExport"
        Me.cmbExport.Size = New System.Drawing.Size(235, 21)
        Me.cmbExport.TabIndex = 105
        Me.cmbExport.Visible = False
        '
        'cmbAgent
        '
        Me.cmbAgent.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmbAgent.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAgent.FormattingEnabled = True
        Me.cmbAgent.Location = New System.Drawing.Point(207, -30)
        Me.cmbAgent.Name = "cmbAgent"
        Me.cmbAgent.Size = New System.Drawing.Size(235, 21)
        Me.cmbAgent.TabIndex = 102
        Me.cmbAgent.Visible = False
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDelete.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete.Location = New System.Drawing.Point(2, 4)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(231, 42)
        Me.btnDelete.TabIndex = 18
        Me.btnDelete.Text = "Delete Tables in Navision"
        Me.btnDelete.UseVisualStyleBackColor = True
        Me.btnDelete.Visible = False
        '
        'btnEx
        '
        Me.btnEx.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnEx.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnEx.Location = New System.Drawing.Point(274, 4)
        Me.btnEx.Name = "btnEx"
        Me.btnEx.Size = New System.Drawing.Size(168, 42)
        Me.btnEx.TabIndex = 17
        Me.btnEx.Text = "Export to Navision"
        Me.btnEx.UseVisualStyleBackColor = True
        Me.btnEx.Visible = False
        '
        'lblExportDate
        '
        Me.lblExportDate.AutoSize = True
        Me.lblExportDate.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblExportDate.Location = New System.Drawing.Point(5, 34)
        Me.lblExportDate.Name = "lblExportDate"
        Me.lblExportDate.Size = New System.Drawing.Size(206, 13)
        Me.lblExportDate.TabIndex = 106
        Me.lblExportDate.Text = "Export Date . . . . . . . . . . . . . . . . . . . . . ."
        Me.lblExportDate.Visible = False
        '
        'ImportStatus
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(652, 415)
        Me.Controls.Add(Me.dgvStatus)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "ImportStatus"
        Me.Text = "Import Status"
        CType(Me.dgvStatus, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnIm As System.Windows.Forms.Button
    Friend WithEvents dgvStatus As System.Windows.Forms.DataGridView
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Tables As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Status As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnEx As System.Windows.Forms.Button
    Friend WithEvents cmbAgent As System.Windows.Forms.ComboBox
    Friend WithEvents cmbExport As System.Windows.Forms.ComboBox
    Friend WithEvents lblExportDate As System.Windows.Forms.Label
    Friend WithEvents chkSelAll As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents chkPayment As System.Windows.Forms.CheckBox
    Friend WithEvents chkOrders As System.Windows.Forms.CheckBox
End Class
