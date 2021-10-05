<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ConsignmentOrder
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
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.btnCrOrder = New System.Windows.Forms.Button
        Me.pnlTop = New System.Windows.Forms.Panel
        Me.chkAllCheque = New System.Windows.Forms.CheckBox
        Me.chkSelAll = New System.Windows.Forms.CheckBox
        Me.dtpToDate = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.cmbAgent = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpFromDate = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.dgvPDC = New System.Windows.Forms.DataGridView
        Me.Check = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ChqNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ChqDt = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CustName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Amount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Panel1.SuspendLayout()
        Me.pnlTop.SuspendLayout()
        CType(Me.dgvPDC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnCrOrder)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 378)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(671, 37)
        Me.Panel1.TabIndex = 16
        '
        'btnCrOrder
        '
        Me.btnCrOrder.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCrOrder.Location = New System.Drawing.Point(574, 7)
        Me.btnCrOrder.Name = "btnCrOrder"
        Me.btnCrOrder.Size = New System.Drawing.Size(93, 23)
        Me.btnCrOrder.TabIndex = 0
        Me.btnCrOrder.Text = "&Create Order"
        Me.btnCrOrder.UseVisualStyleBackColor = True
        '
        'pnlTop
        '
        Me.pnlTop.Controls.Add(Me.chkAllCheque)
        Me.pnlTop.Controls.Add(Me.chkSelAll)
        Me.pnlTop.Controls.Add(Me.dtpToDate)
        Me.pnlTop.Controls.Add(Me.Label3)
        Me.pnlTop.Controls.Add(Me.cmbAgent)
        Me.pnlTop.Controls.Add(Me.Label2)
        Me.pnlTop.Controls.Add(Me.dtpFromDate)
        Me.pnlTop.Controls.Add(Me.Label1)
        Me.pnlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlTop.Name = "pnlTop"
        Me.pnlTop.Size = New System.Drawing.Size(671, 152)
        Me.pnlTop.TabIndex = 22
        '
        'chkAllCheque
        '
        Me.chkAllCheque.AutoSize = True
        Me.chkAllCheque.Location = New System.Drawing.Point(7, 12)
        Me.chkAllCheque.Name = "chkAllCheque"
        Me.chkAllCheque.Size = New System.Drawing.Size(102, 17)
        Me.chkAllCheque.TabIndex = 24
        Me.chkAllCheque.Text = "Show All Orders"
        Me.chkAllCheque.UseVisualStyleBackColor = True
        '
        'chkSelAll
        '
        Me.chkSelAll.AutoSize = True
        Me.chkSelAll.Location = New System.Drawing.Point(7, 124)
        Me.chkSelAll.Name = "chkSelAll"
        Me.chkSelAll.Size = New System.Drawing.Size(75, 17)
        Me.chkSelAll.TabIndex = 23
        Me.chkSelAll.Text = "Select ALL"
        Me.chkSelAll.UseVisualStyleBackColor = True
        '
        'dtpToDate
        '
        Me.dtpToDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpToDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpToDate.Location = New System.Drawing.Point(191, 59)
        Me.dtpToDate.Name = "dtpToDate"
        Me.dtpToDate.Size = New System.Drawing.Size(103, 21)
        Me.dtpToDate.TabIndex = 21
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 68)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(203, 13)
        Me.Label3.TabIndex = 22
        Me.Label3.Text = "To Date. . . . . . . . . . . . . . . . . . . . . . ."
        '
        'cmbAgent
        '
        Me.cmbAgent.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAgent.FormattingEnabled = True
        Me.cmbAgent.Location = New System.Drawing.Point(191, 85)
        Me.cmbAgent.Name = "cmbAgent"
        Me.cmbAgent.Size = New System.Drawing.Size(292, 21)
        Me.cmbAgent.TabIndex = 20
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 93)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(320, 13)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "Select the Consignment Customer. . . . . . . . . . . . . . . . . . . . . ."
        '
        'dtpFromDate
        '
        Me.dtpFromDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpFromDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFromDate.Location = New System.Drawing.Point(191, 32)
        Me.dtpFromDate.Name = "dtpFromDate"
        Me.dtpFromDate.Size = New System.Drawing.Size(103, 21)
        Me.dtpFromDate.TabIndex = 17
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 41)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(250, 13)
        Me.Label1.TabIndex = 18
        Me.Label1.Text = "From Date. . . . . . . . . . . . . . . . . . . . . . . . . . . ."
        '
        'dgvPDC
        '
        Me.dgvPDC.AllowUserToDeleteRows = False
        Me.dgvPDC.AllowUserToOrderColumns = True
        Me.dgvPDC.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvPDC.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Check, Me.ChqNo, Me.ChqDt, Me.CustName, Me.Amount})
        Me.dgvPDC.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvPDC.Location = New System.Drawing.Point(0, 152)
        Me.dgvPDC.Name = "dgvPDC"
        Me.dgvPDC.Size = New System.Drawing.Size(671, 226)
        Me.dgvPDC.TabIndex = 23
        '
        'Check
        '
        Me.Check.HeaderText = ""
        Me.Check.Name = "Check"
        Me.Check.Width = 30
        '
        'ChqNo
        '
        Me.ChqNo.HeaderText = "Order No"
        Me.ChqNo.Name = "ChqNo"
        Me.ChqNo.ReadOnly = True
        '
        'ChqDt
        '
        Me.ChqDt.HeaderText = "Order Date"
        Me.ChqDt.Name = "ChqDt"
        Me.ChqDt.ReadOnly = True
        Me.ChqDt.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'CustName
        '
        Me.CustName.HeaderText = "Sales Agent"
        Me.CustName.Name = "CustName"
        Me.CustName.ReadOnly = True
        Me.CustName.Width = 200
        '
        'Amount
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomRight
        DataGridViewCellStyle2.NullValue = Nothing
        Me.Amount.DefaultCellStyle = DataGridViewCellStyle2
        Me.Amount.HeaderText = "Amount"
        Me.Amount.Name = "Amount"
        Me.Amount.ReadOnly = True
        '
        'ConsignmentOrder
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(671, 415)
        Me.Controls.Add(Me.dgvPDC)
        Me.Controls.Add(Me.pnlTop)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "ConsignmentOrder"
        Me.Text = "Consignment Order"
        Me.Panel1.ResumeLayout(False)
        Me.pnlTop.ResumeLayout(False)
        Me.pnlTop.PerformLayout()
        CType(Me.dgvPDC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnCrOrder As System.Windows.Forms.Button
    Friend WithEvents pnlTop As System.Windows.Forms.Panel
    Friend WithEvents chkAllCheque As System.Windows.Forms.CheckBox
    Friend WithEvents chkSelAll As System.Windows.Forms.CheckBox
    Friend WithEvents dtpToDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmbAgent As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpFromDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dgvPDC As System.Windows.Forms.DataGridView
    Friend WithEvents Check As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ChqNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ChqDt As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CustName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Amount As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
