<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PaymPostDated
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
        Me.lblAgentNo = New System.Windows.Forms.Label
        Me.lbldate = New System.Windows.Forms.Label
        Me.lblPayMeth = New System.Windows.Forms.Label
        Me.cmbAgent = New System.Windows.Forms.ComboBox
        Me.cmbTerms = New System.Windows.Forms.ComboBox
        Me.dtpToDate = New System.Windows.Forms.DateTimePicker
        Me.btnPrint = New System.Windows.Forms.Button
        Me.dtpFromDate = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.rbChqDt = New System.Windows.Forms.RadioButton
        Me.rbCust = New System.Windows.Forms.RadioButton
        Me.SuspendLayout()
        '
        'lblAgentNo
        '
        Me.lblAgentNo.AutoSize = True
        Me.lblAgentNo.Location = New System.Drawing.Point(10, 71)
        Me.lblAgentNo.Name = "lblAgentNo"
        Me.lblAgentNo.Size = New System.Drawing.Size(183, 13)
        Me.lblAgentNo.TabIndex = 109
        Me.lblAgentNo.Text = "Salesman Name . . . . . . . . . . . . . . "
        '
        'lbldate
        '
        Me.lbldate.AutoSize = True
        Me.lbldate.Location = New System.Drawing.Point(10, 43)
        Me.lbldate.Name = "lbldate"
        Me.lbldate.Size = New System.Drawing.Size(171, 13)
        Me.lbldate.TabIndex = 108
        Me.lbldate.Text = "To Date. . . . . . . . . . . . . . . . . . "
        '
        'lblPayMeth
        '
        Me.lblPayMeth.AutoSize = True
        Me.lblPayMeth.Location = New System.Drawing.Point(10, 128)
        Me.lblPayMeth.Name = "lblPayMeth"
        Me.lblPayMeth.Size = New System.Drawing.Size(147, 13)
        Me.lblPayMeth.TabIndex = 106
        Me.lblPayMeth.Text = "Payment Method . . . . . . . . "
        Me.lblPayMeth.Visible = False
        '
        'cmbAgent
        '
        Me.cmbAgent.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAgent.FormattingEnabled = True
        Me.cmbAgent.Location = New System.Drawing.Point(154, 68)
        Me.cmbAgent.Name = "cmbAgent"
        Me.cmbAgent.Size = New System.Drawing.Size(230, 21)
        Me.cmbAgent.TabIndex = 110
        '
        'cmbTerms
        '
        Me.cmbTerms.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTerms.FormattingEnabled = True
        Me.cmbTerms.Items.AddRange(New Object() {"All", "Cash", "Cheque"})
        Me.cmbTerms.Location = New System.Drawing.Point(154, 125)
        Me.cmbTerms.Name = "cmbTerms"
        Me.cmbTerms.Size = New System.Drawing.Size(230, 21)
        Me.cmbTerms.TabIndex = 107
        Me.cmbTerms.Visible = False
        '
        'dtpToDate
        '
        Me.dtpToDate.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.dtpToDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpToDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpToDate.Location = New System.Drawing.Point(154, 39)
        Me.dtpToDate.Name = "dtpToDate"
        Me.dtpToDate.Size = New System.Drawing.Size(230, 21)
        Me.dtpToDate.TabIndex = 105
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(155, 124)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(75, 23)
        Me.btnPrint.TabIndex = 104
        Me.btnPrint.Text = "View"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'dtpFromDate
        '
        Me.dtpFromDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpFromDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFromDate.Location = New System.Drawing.Point(153, 12)
        Me.dtpFromDate.Name = "dtpFromDate"
        Me.dtpFromDate.Size = New System.Drawing.Size(230, 21)
        Me.dtpFromDate.TabIndex = 111
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(183, 13)
        Me.Label1.TabIndex = 112
        Me.Label1.Text = "From Date. . . . . . . . . . . . . . . . . . "
        '
        'rbChqDt
        '
        Me.rbChqDt.AutoSize = True
        Me.rbChqDt.Checked = True
        Me.rbChqDt.Location = New System.Drawing.Point(115, 100)
        Me.rbChqDt.Name = "rbChqDt"
        Me.rbChqDt.Size = New System.Drawing.Size(70, 17)
        Me.rbChqDt.TabIndex = 113
        Me.rbChqDt.TabStop = True
        Me.rbChqDt.Text = "Chq Date"
        Me.rbChqDt.UseVisualStyleBackColor = True
        '
        'rbCust
        '
        Me.rbCust.AutoSize = True
        Me.rbCust.Location = New System.Drawing.Point(208, 100)
        Me.rbCust.Name = "rbCust"
        Me.rbCust.Size = New System.Drawing.Size(71, 17)
        Me.rbCust.TabIndex = 114
        Me.rbCust.Text = "Customer"
        Me.rbCust.UseVisualStyleBackColor = True
        '
        'PaymPostDated
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(395, 154)
        Me.Controls.Add(Me.rbCust)
        Me.Controls.Add(Me.rbChqDt)
        Me.Controls.Add(Me.dtpFromDate)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.cmbAgent)
        Me.Controls.Add(Me.lblAgentNo)
        Me.Controls.Add(Me.cmbTerms)
        Me.Controls.Add(Me.lblPayMeth)
        Me.Controls.Add(Me.dtpToDate)
        Me.Controls.Add(Me.lbldate)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "PaymPostDated"
        Me.Text = "Post Dated Cheque Report"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmbAgent As System.Windows.Forms.ComboBox
    Friend WithEvents cmbTerms As System.Windows.Forms.ComboBox
    Friend WithEvents dtpToDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents lblAgentNo As System.Windows.Forms.Label
    Friend WithEvents lbldate As System.Windows.Forms.Label
    Friend WithEvents lblPayMeth As System.Windows.Forms.Label
    Friend WithEvents dtpFromDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents rbChqDt As System.Windows.Forms.RadioButton
    Friend WithEvents rbCust As System.Windows.Forms.RadioButton
End Class
