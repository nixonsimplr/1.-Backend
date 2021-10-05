<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PayColInterval
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
        Dim Label3 As System.Windows.Forms.Label
        Dim Label2 As System.Windows.Forms.Label
        Dim Label1 As System.Windows.Forms.Label
        Dim Label20 As System.Windows.Forms.Label
        Me.dtpStartDate = New System.Windows.Forms.DateTimePicker
        Me.cmbAgent = New System.Windows.Forms.ComboBox
        Me.cmbPayMethod = New System.Windows.Forms.ComboBox
        Me.dtpEndDate = New System.Windows.Forms.DateTimePicker
        Me.btnPrint = New System.Windows.Forms.Button
        Label3 = New System.Windows.Forms.Label
        Label2 = New System.Windows.Forms.Label
        Label1 = New System.Windows.Forms.Label
        Label20 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Label3
        '
        Label3.AutoSize = True
        Label3.Location = New System.Drawing.Point(8, 11)
        Label3.Name = "Label3"
        Label3.Size = New System.Drawing.Size(158, 13)
        Label3.TabIndex = 123
        Label3.Text = "From Date . . . . . . . . . . . . . . "
        '
        'Label2
        '
        Label2.AutoSize = True
        Label2.Location = New System.Drawing.Point(8, 89)
        Label2.Name = "Label2"
        Label2.Size = New System.Drawing.Size(153, 13)
        Label2.TabIndex = 120
        Label2.Text = "Agent No . . . . . . . . . . . . . . "
        '
        'Label1
        '
        Label1.AutoSize = True
        Label1.Location = New System.Drawing.Point(8, 37)
        Label1.Name = "Label1"
        Label1.Size = New System.Drawing.Size(153, 13)
        Label1.TabIndex = 119
        Label1.Text = "To Date . . . . . . . . . . . . . . . "
        '
        'Label20
        '
        Label20.AutoSize = True
        Label20.Location = New System.Drawing.Point(8, 62)
        Label20.Name = "Label20"
        Label20.Size = New System.Drawing.Size(158, 13)
        Label20.TabIndex = 117
        Label20.Text = "Payment Method . . . . . . . . . ."
        '
        'dtpStartDate
        '
        Me.dtpStartDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpStartDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpStartDate.Location = New System.Drawing.Point(152, 7)
        Me.dtpStartDate.Name = "dtpStartDate"
        Me.dtpStartDate.Size = New System.Drawing.Size(153, 21)
        Me.dtpStartDate.TabIndex = 122
        '
        'cmbAgent
        '
        Me.cmbAgent.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAgent.FormattingEnabled = True
        Me.cmbAgent.Location = New System.Drawing.Point(152, 86)
        Me.cmbAgent.Name = "cmbAgent"
        Me.cmbAgent.Size = New System.Drawing.Size(153, 21)
        Me.cmbAgent.TabIndex = 121
        '
        'cmbPayMethod
        '
        Me.cmbPayMethod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPayMethod.FormattingEnabled = True
        Me.cmbPayMethod.Items.AddRange(New Object() {"All", "Cash", "Cheque"})
        Me.cmbPayMethod.Location = New System.Drawing.Point(152, 59)
        Me.cmbPayMethod.Name = "cmbPayMethod"
        Me.cmbPayMethod.Size = New System.Drawing.Size(153, 21)
        Me.cmbPayMethod.TabIndex = 118
        '
        'dtpEndDate
        '
        Me.dtpEndDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpEndDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpEndDate.Location = New System.Drawing.Point(152, 33)
        Me.dtpEndDate.Name = "dtpEndDate"
        Me.dtpEndDate.Size = New System.Drawing.Size(153, 21)
        Me.dtpEndDate.TabIndex = 116
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(120, 133)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(75, 23)
        Me.btnPrint.TabIndex = 115
        Me.btnPrint.Text = "View"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'PayColInterval
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(315, 162)
        Me.Controls.Add(Me.dtpStartDate)
        Me.Controls.Add(Me.cmbAgent)
        Me.Controls.Add(Label2)
        Me.Controls.Add(Me.cmbPayMethod)
        Me.Controls.Add(Label20)
        Me.Controls.Add(Me.dtpEndDate)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Label3)
        Me.Controls.Add(Label1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "PayColInterval"
        Me.Text = "Payment Collection Interval"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dtpStartDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents cmbAgent As System.Windows.Forms.ComboBox
    Friend WithEvents cmbPayMethod As System.Windows.Forms.ComboBox
    Friend WithEvents dtpEndDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnPrint As System.Windows.Forms.Button
End Class
