<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Delivery
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
        Me.cbDelivery = New System.Windows.Forms.ComboBox
        Me.dtpDelivery = New System.Windows.Forms.DateTimePicker
        Me.btnPrint = New System.Windows.Forms.Button
        Label3 = New System.Windows.Forms.Label
        Label2 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Label3
        '
        Label3.AutoSize = True
        Label3.Location = New System.Drawing.Point(12, 16)
        Label3.Name = "Label3"
        Label3.Size = New System.Drawing.Size(163, 13)
        Label3.TabIndex = 48
        Label3.Text = "Delivery Date . . . . . . . . . . . . ."
        '
        'Label2
        '
        Label2.AutoSize = True
        Label2.Location = New System.Drawing.Point(12, 45)
        Label2.Name = "Label2"
        Label2.Size = New System.Drawing.Size(163, 13)
        Label2.TabIndex = 47
        Label2.Text = "Delivery Area . . . . . . . . . . . . ."
        '
        'cbDelivery
        '
        Me.cbDelivery.FormattingEnabled = True
        Me.cbDelivery.Location = New System.Drawing.Point(170, 42)
        Me.cbDelivery.Name = "cbDelivery"
        Me.cbDelivery.Size = New System.Drawing.Size(124, 21)
        Me.cbDelivery.TabIndex = 50
        '
        'dtpDelivery
        '
        Me.dtpDelivery.CustomFormat = "dd/MM/yyyy"
        Me.dtpDelivery.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDelivery.Location = New System.Drawing.Point(170, 12)
        Me.dtpDelivery.Name = "dtpDelivery"
        Me.dtpDelivery.Size = New System.Drawing.Size(124, 21)
        Me.dtpDelivery.TabIndex = 49
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(119, 80)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(75, 23)
        Me.btnPrint.TabIndex = 105
        Me.btnPrint.Text = "View"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'Delivery
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(314, 119)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.cbDelivery)
        Me.Controls.Add(Me.dtpDelivery)
        Me.Controls.Add(Label3)
        Me.Controls.Add(Label2)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "Delivery"
        Me.Text = "Delivery"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cbDelivery As System.Windows.Forms.ComboBox
    Friend WithEvents dtpDelivery As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnPrint As System.Windows.Forms.Button
End Class
