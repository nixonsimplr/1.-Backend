<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class VoidInvoice
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
        Dim Label2 As System.Windows.Forms.Label
        Me.cmbAgent = New System.Windows.Forms.ComboBox
        Me.btnPrint = New System.Windows.Forms.Button
        Label2 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Label2
        '
        Label2.AutoSize = True
        Label2.Location = New System.Drawing.Point(10, 15)
        Label2.Name = "Label2"
        Label2.Size = New System.Drawing.Size(153, 13)
        Label2.TabIndex = 94
        Label2.Text = "Agent No . . . . . . . . . . . . . . "
        '
        'cmbAgent
        '
        Me.cmbAgent.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAgent.FormattingEnabled = True
        Me.cmbAgent.Location = New System.Drawing.Point(154, 12)
        Me.cmbAgent.Name = "cmbAgent"
        Me.cmbAgent.Size = New System.Drawing.Size(153, 21)
        Me.cmbAgent.TabIndex = 95
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(122, 54)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(75, 23)
        Me.btnPrint.TabIndex = 93
        Me.btnPrint.Text = "View"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'VoidInvoice
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(319, 89)
        Me.Controls.Add(Me.cmbAgent)
        Me.Controls.Add(Label2)
        Me.Controls.Add(Me.btnPrint)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "VoidInvoice"
        Me.Text = "Void Invoice"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmbAgent As System.Windows.Forms.ComboBox
    Friend WithEvents btnPrint As System.Windows.Forms.Button
End Class
