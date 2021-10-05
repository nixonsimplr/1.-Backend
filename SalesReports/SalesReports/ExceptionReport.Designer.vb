<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ExceptionReport
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
        Me.dtpFromDate = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnPrint = New System.Windows.Forms.Button
        Me.cmbAgent = New System.Windows.Forms.ComboBox
        Me.lblAgentNo = New System.Windows.Forms.Label
        Me.dtpToDate = New System.Windows.Forms.DateTimePicker
        Me.lbldate = New System.Windows.Forms.Label
        Me.cmbType = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'dtpFromDate
        '
        Me.dtpFromDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpFromDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFromDate.Location = New System.Drawing.Point(133, 10)
        Me.dtpFromDate.Name = "dtpFromDate"
        Me.dtpFromDate.Size = New System.Drawing.Size(221, 21)
        Me.dtpFromDate.TabIndex = 120
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(5, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(183, 13)
        Me.Label1.TabIndex = 121
        Me.Label1.Text = "From Date. . . . . . . . . . . . . . . . . . "
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(143, 127)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(75, 23)
        Me.btnPrint.TabIndex = 113
        Me.btnPrint.Text = "View"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'cmbAgent
        '
        Me.cmbAgent.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAgent.FormattingEnabled = True
        Me.cmbAgent.Location = New System.Drawing.Point(134, 64)
        Me.cmbAgent.Name = "cmbAgent"
        Me.cmbAgent.Size = New System.Drawing.Size(221, 21)
        Me.cmbAgent.TabIndex = 119
        '
        'lblAgentNo
        '
        Me.lblAgentNo.AutoSize = True
        Me.lblAgentNo.Location = New System.Drawing.Point(6, 67)
        Me.lblAgentNo.Name = "lblAgentNo"
        Me.lblAgentNo.Size = New System.Drawing.Size(183, 13)
        Me.lblAgentNo.TabIndex = 118
        Me.lblAgentNo.Text = "Salesman Name . . . . . . . . . . . . . . "
        '
        'dtpToDate
        '
        Me.dtpToDate.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.dtpToDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpToDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpToDate.Location = New System.Drawing.Point(134, 37)
        Me.dtpToDate.Name = "dtpToDate"
        Me.dtpToDate.Size = New System.Drawing.Size(221, 21)
        Me.dtpToDate.TabIndex = 114
        '
        'lbldate
        '
        Me.lbldate.AutoSize = True
        Me.lbldate.Location = New System.Drawing.Point(6, 41)
        Me.lbldate.Name = "lbldate"
        Me.lbldate.Size = New System.Drawing.Size(171, 13)
        Me.lbldate.TabIndex = 117
        Me.lbldate.Text = "To Date. . . . . . . . . . . . . . . . . . "
        '
        'cmbType
        '
        Me.cmbType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbType.FormattingEnabled = True
        Me.cmbType.Location = New System.Drawing.Point(133, 91)
        Me.cmbType.Name = "cmbType"
        Me.cmbType.Size = New System.Drawing.Size(221, 21)
        Me.cmbType.TabIndex = 123
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(5, 94)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(182, 13)
        Me.Label2.TabIndex = 122
        Me.Label2.Text = "Exception Type . . . . . . . . . . . . . . "
        '
        'ExceptionReport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(360, 163)
        Me.Controls.Add(Me.cmbType)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dtpFromDate)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.cmbAgent)
        Me.Controls.Add(Me.lblAgentNo)
        Me.Controls.Add(Me.dtpToDate)
        Me.Controls.Add(Me.lbldate)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "ExceptionReport"
        Me.Text = "Exception Report By Salesman"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dtpFromDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents cmbAgent As System.Windows.Forms.ComboBox
    Friend WithEvents lblAgentNo As System.Windows.Forms.Label
    Friend WithEvents dtpToDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lbldate As System.Windows.Forms.Label
    Friend WithEvents cmbType As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
End Class
