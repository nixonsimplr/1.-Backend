<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PostDatedCheque
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
        Me.btnApprove = New System.Windows.Forms.Button
        Me.dgvPDC = New System.Windows.Forms.DataGridView
        Me.Check = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ChqDt = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ChqNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BankName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.RcptNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.RcptDt = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CustNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CustName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Amount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.btnDecline = New System.Windows.Forms.Button
        Me.cmbAgent = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpFromDate = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.pnlTop = New System.Windows.Forms.Panel
        Me.chkAllCheque = New System.Windows.Forms.CheckBox
        Me.chkSelAll = New System.Windows.Forms.CheckBox
        Me.dtpToDate = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.pnlGrid = New System.Windows.Forms.Panel
        CType(Me.dgvPDC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.pnlTop.SuspendLayout()
        Me.pnlGrid.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnApprove
        '
        Me.btnApprove.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnApprove.Location = New System.Drawing.Point(628, 7)
        Me.btnApprove.Name = "btnApprove"
        Me.btnApprove.Size = New System.Drawing.Size(93, 23)
        Me.btnApprove.TabIndex = 0
        Me.btnApprove.Text = "&Approve"
        Me.btnApprove.UseVisualStyleBackColor = True
        '
        'dgvPDC
        '
        Me.dgvPDC.AllowUserToDeleteRows = False
        Me.dgvPDC.AllowUserToOrderColumns = True
        Me.dgvPDC.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvPDC.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Check, Me.ChqDt, Me.ChqNo, Me.BankName, Me.RcptNo, Me.RcptDt, Me.CustNo, Me.CustName, Me.Amount})
        Me.dgvPDC.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvPDC.Location = New System.Drawing.Point(0, 0)
        Me.dgvPDC.Name = "dgvPDC"
        Me.dgvPDC.Size = New System.Drawing.Size(725, 278)
        Me.dgvPDC.TabIndex = 16
        '
        'Check
        '
        Me.Check.HeaderText = ""
        Me.Check.Name = "Check"
        Me.Check.Width = 30
        '
        'ChqDt
        '
        Me.ChqDt.HeaderText = "Cheque Date"
        Me.ChqDt.Name = "ChqDt"
        Me.ChqDt.ReadOnly = True
        Me.ChqDt.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'ChqNo
        '
        Me.ChqNo.HeaderText = "Cheque No"
        Me.ChqNo.Name = "ChqNo"
        Me.ChqNo.ReadOnly = True
        '
        'BankName
        '
        Me.BankName.HeaderText = "Bank Name"
        Me.BankName.Name = "BankName"
        Me.BankName.ReadOnly = True
        Me.BankName.Width = 150
        '
        'RcptNo
        '
        Me.RcptNo.HeaderText = "Receipt No"
        Me.RcptNo.Name = "RcptNo"
        Me.RcptNo.ReadOnly = True
        '
        'RcptDt
        '
        Me.RcptDt.HeaderText = "Receipt Date"
        Me.RcptDt.Name = "RcptDt"
        Me.RcptDt.ReadOnly = True
        '
        'CustNo
        '
        Me.CustNo.HeaderText = "Cust No."
        Me.CustNo.Name = "CustNo"
        Me.CustNo.ReadOnly = True
        '
        'CustName
        '
        Me.CustName.HeaderText = "Customer Name"
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
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnDecline)
        Me.Panel1.Controls.Add(Me.btnApprove)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 430)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(725, 37)
        Me.Panel1.TabIndex = 15
        '
        'btnDecline
        '
        Me.btnDecline.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDecline.Location = New System.Drawing.Point(529, 7)
        Me.btnDecline.Name = "btnDecline"
        Me.btnDecline.Size = New System.Drawing.Size(93, 23)
        Me.btnDecline.TabIndex = 1
        Me.btnDecline.Text = "&Decline"
        Me.btnDecline.UseVisualStyleBackColor = True
        '
        'cmbAgent
        '
        Me.cmbAgent.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAgent.FormattingEnabled = True
        Me.cmbAgent.Location = New System.Drawing.Point(114, 92)
        Me.cmbAgent.Name = "cmbAgent"
        Me.cmbAgent.Size = New System.Drawing.Size(135, 21)
        Me.cmbAgent.TabIndex = 20
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(5, 95)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(201, 13)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "Agent ID. . . . . . . . . . . . . . . . . . . . . ."
        '
        'dtpFromDate
        '
        Me.dtpFromDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpFromDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFromDate.Location = New System.Drawing.Point(114, 39)
        Me.dtpFromDate.Name = "dtpFromDate"
        Me.dtpFromDate.Size = New System.Drawing.Size(135, 21)
        Me.dtpFromDate.TabIndex = 17
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 41)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(124, 13)
        Me.Label1.TabIndex = 18
        Me.Label1.Text = "From Date. . . . . . . . . ."
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
        Me.pnlTop.Size = New System.Drawing.Size(725, 152)
        Me.pnlTop.TabIndex = 21
        '
        'chkAllCheque
        '
        Me.chkAllCheque.AutoSize = True
        Me.chkAllCheque.Location = New System.Drawing.Point(7, 12)
        Me.chkAllCheque.Name = "chkAllCheque"
        Me.chkAllCheque.Size = New System.Drawing.Size(111, 17)
        Me.chkAllCheque.TabIndex = 24
        Me.chkAllCheque.Text = "Show All Cheques"
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
        Me.dtpToDate.Location = New System.Drawing.Point(114, 66)
        Me.dtpToDate.Name = "dtpToDate"
        Me.dtpToDate.Size = New System.Drawing.Size(135, 21)
        Me.dtpToDate.TabIndex = 21
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 68)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(112, 13)
        Me.Label3.TabIndex = 22
        Me.Label3.Text = "To Date. . . . . . . . . ."
        '
        'pnlGrid
        '
        Me.pnlGrid.Controls.Add(Me.dgvPDC)
        Me.pnlGrid.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlGrid.Location = New System.Drawing.Point(0, 152)
        Me.pnlGrid.Name = "pnlGrid"
        Me.pnlGrid.Size = New System.Drawing.Size(725, 278)
        Me.pnlGrid.TabIndex = 22
        '
        'PostDatedCheque
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(725, 467)
        Me.Controls.Add(Me.pnlGrid)
        Me.Controls.Add(Me.pnlTop)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "PostDatedCheque"
        Me.Text = "Post Dated Cheque"
        CType(Me.dgvPDC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.pnlTop.ResumeLayout(False)
        Me.pnlTop.PerformLayout()
        Me.pnlGrid.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnApprove As System.Windows.Forms.Button
    Friend WithEvents dgvPDC As System.Windows.Forms.DataGridView
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnDecline As System.Windows.Forms.Button
    Friend WithEvents cmbAgent As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpFromDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents pnlTop As System.Windows.Forms.Panel
    Friend WithEvents pnlGrid As System.Windows.Forms.Panel
    Friend WithEvents dtpToDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents chkSelAll As System.Windows.Forms.CheckBox
    Friend WithEvents Check As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ChqDt As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ChqNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BankName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RcptNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RcptDt As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CustNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CustName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Amount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents chkAllCheque As System.Windows.Forms.CheckBox
End Class
