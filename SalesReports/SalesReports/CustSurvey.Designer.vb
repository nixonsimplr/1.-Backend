<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CustSurvey
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
        Dim Label2 As System.Windows.Forms.Label
        Dim Label1 As System.Windows.Forms.Label
        Me.dtpToDate = New System.Windows.Forms.DateTimePicker
        Me.cmbAgent = New System.Windows.Forms.ComboBox
        Me.dtpFromDate = New System.Windows.Forms.DateTimePicker
        Me.btnPrint = New System.Windows.Forms.Button
        Label3 = New System.Windows.Forms.Label
        Label2 = New System.Windows.Forms.Label
        Label1 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'dtpToDate
        '
        Me.dtpToDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpToDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpToDate.Location = New System.Drawing.Point(123, 38)
        Me.dtpToDate.Name = "dtpToDate"
        Me.dtpToDate.Size = New System.Drawing.Size(153, 20)
        Me.dtpToDate.TabIndex = 123
        '
        'Label3
        '
        Label3.AutoSize = True
        Label3.Location = New System.Drawing.Point(11, 45)
        Label3.Name = "Label3"
        Label3.Size = New System.Drawing.Size(154, 13)
        Label3.TabIndex = 124
        Label3.Text = "To Date. . . . . . . . . . . . . . . . . . "
        '
        'cmbAgent
        '
        Me.cmbAgent.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAgent.FormattingEnabled = True
        Me.cmbAgent.Location = New System.Drawing.Point(123, 64)
        Me.cmbAgent.Name = "cmbAgent"
        Me.cmbAgent.Size = New System.Drawing.Size(153, 21)
        Me.cmbAgent.TabIndex = 122
        '
        'Label2
        '
        Label2.AutoSize = True
        Label2.Location = New System.Drawing.Point(11, 72)
        Label2.Name = "Label2"
        Label2.Size = New System.Drawing.Size(139, 13)
        Label2.TabIndex = 121
        Label2.Text = "Agent No . . . . . . . . . . . . . . "
        '
        'dtpFromDate
        '
        Me.dtpFromDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpFromDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFromDate.Location = New System.Drawing.Point(123, 12)
        Me.dtpFromDate.Name = "dtpFromDate"
        Me.dtpFromDate.Size = New System.Drawing.Size(153, 20)
        Me.dtpFromDate.TabIndex = 119
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(109, 100)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(75, 23)
        Me.btnPrint.TabIndex = 118
        Me.btnPrint.Text = "Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Label1.AutoSize = True
        Label1.Location = New System.Drawing.Point(11, 19)
        Label1.Name = "Label1"
        Label1.Size = New System.Drawing.Size(164, 13)
        Label1.TabIndex = 120
        Label1.Text = "From Date. . . . . . . . . . . . . . . . . . "
        '
        'CustSurvey
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(296, 136)
        Me.Controls.Add(Me.dtpToDate)
        Me.Controls.Add(Label3)
        Me.Controls.Add(Me.cmbAgent)
        Me.Controls.Add(Label2)
        Me.Controls.Add(Me.dtpFromDate)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Label1)
        Me.Name = "CustSurvey"
        Me.Text = "Customer Survey"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dtpToDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents cmbAgent As System.Windows.Forms.ComboBox
    Friend WithEvents dtpFromDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnPrint As System.Windows.Forms.Button
End Class
