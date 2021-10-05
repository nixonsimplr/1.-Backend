<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DailySalesCust
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
        Dim Label1 As System.Windows.Forms.Label
        Me.lblItem = New System.Windows.Forms.Label
        Me.rbItem = New System.Windows.Forms.RadioButton
        Me.rbCustomer = New System.Windows.Forms.RadioButton
        Me.cmbCustomer = New System.Windows.Forms.ComboBox
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.btnPrint = New System.Windows.Forms.Button
        Label1 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Label1
        '
        Label1.AutoSize = True
        Label1.Location = New System.Drawing.Point(8, 16)
        Label1.Name = "Label1"
        Label1.Size = New System.Drawing.Size(156, 13)
        Label1.TabIndex = 122
        Label1.Text = "Date. . . . . . . . . . . . . . . . . . "
        '
        'lblItem
        '
        Me.lblItem.AutoSize = True
        Me.lblItem.Location = New System.Drawing.Point(8, 49)
        Me.lblItem.Name = "lblItem"
        Me.lblItem.Size = New System.Drawing.Size(0, 13)
        Me.lblItem.TabIndex = 126
        '
        'rbItem
        '
        Me.rbItem.AutoSize = True
        Me.rbItem.Location = New System.Drawing.Point(165, 77)
        Me.rbItem.Name = "rbItem"
        Me.rbItem.Size = New System.Drawing.Size(62, 17)
        Me.rbItem.TabIndex = 125
        Me.rbItem.TabStop = True
        Me.rbItem.Text = "By Item"
        Me.rbItem.UseVisualStyleBackColor = True
        '
        'rbCustomer
        '
        Me.rbCustomer.AutoSize = True
        Me.rbCustomer.Location = New System.Drawing.Point(57, 77)
        Me.rbCustomer.Name = "rbCustomer"
        Me.rbCustomer.Size = New System.Drawing.Size(86, 17)
        Me.rbCustomer.TabIndex = 124
        Me.rbCustomer.TabStop = True
        Me.rbCustomer.Text = "By Customer"
        Me.rbCustomer.UseVisualStyleBackColor = True
        '
        'cmbCustomer
        '
        Me.cmbCustomer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCustomer.FormattingEnabled = True
        Me.cmbCustomer.Location = New System.Drawing.Point(152, 41)
        Me.cmbCustomer.Name = "cmbCustomer"
        Me.cmbCustomer.Size = New System.Drawing.Size(153, 21)
        Me.cmbCustomer.TabIndex = 123
        '
        'dtpDate
        '
        Me.dtpDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(152, 12)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(153, 21)
        Me.dtpDate.TabIndex = 121
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(110, 109)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(75, 23)
        Me.btnPrint.TabIndex = 120
        Me.btnPrint.Text = "View"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'DailySalesCust
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(320, 138)
        Me.Controls.Add(Me.lblItem)
        Me.Controls.Add(Me.rbItem)
        Me.Controls.Add(Me.rbCustomer)
        Me.Controls.Add(Me.cmbCustomer)
        Me.Controls.Add(Me.dtpDate)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Label1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "DailySalesCust"
        Me.Text = "Daily Sales By Customer and Item"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblItem As System.Windows.Forms.Label
    Friend WithEvents rbItem As System.Windows.Forms.RadioButton
    Friend WithEvents rbCustomer As System.Windows.Forms.RadioButton
    Friend WithEvents cmbCustomer As System.Windows.Forms.ComboBox
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnPrint As System.Windows.Forms.Button
End Class
