<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class GRNReport
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
        Dim Label1 As System.Windows.Forms.Label
        Dim Label2 As System.Windows.Forms.Label
        Dim Label4 As System.Windows.Forms.Label
        Me.dtpToDate = New System.Windows.Forms.DateTimePicker
        Me.dtpFromDate = New System.Windows.Forms.DateTimePicker
        Me.cmbAgent = New System.Windows.Forms.ComboBox
        Me.btnPrint = New System.Windows.Forms.Button
        Me.cmbItem = New System.Windows.Forms.ComboBox
        Me.rdoSum = New System.Windows.Forms.RadioButton
        Me.RdoDet = New System.Windows.Forms.RadioButton
        Label3 = New System.Windows.Forms.Label
        Label1 = New System.Windows.Forms.Label
        Label2 = New System.Windows.Forms.Label
        Label4 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Label3
        '
        Label3.AutoSize = True
        Label3.Location = New System.Drawing.Point(9, 40)
        Label3.Name = "Label3"
        Label3.Size = New System.Drawing.Size(171, 13)
        Label3.TabIndex = 130
        Label3.Text = "To Date. . . . . . . . . . . . . . . . . . "
        AddHandler Label3.Click, AddressOf Me.Label3_Click
        '
        'Label1
        '
        Label1.AutoSize = True
        Label1.Location = New System.Drawing.Point(9, 14)
        Label1.Name = "Label1"
        Label1.Size = New System.Drawing.Size(183, 13)
        Label1.TabIndex = 128
        Label1.Text = "From Date. . . . . . . . . . . . . . . . . . "
        AddHandler Label1.Click, AddressOf Me.Label1_Click
        '
        'Label2
        '
        Label2.AutoSize = True
        Label2.Location = New System.Drawing.Point(10, 98)
        Label2.Name = "Label2"
        Label2.Size = New System.Drawing.Size(153, 13)
        Label2.TabIndex = 125
        Label2.Text = "Agent No . . . . . . . . . . . . . . "
        '
        'Label4
        '
        Label4.AutoSize = True
        Label4.Location = New System.Drawing.Point(9, 70)
        Label4.Name = "Label4"
        Label4.Size = New System.Drawing.Size(172, 13)
        Label4.TabIndex = 131
        Label4.Text = "Item Code . . . . . . . . . . . . . . . . "
        '
        'dtpToDate
        '
        Me.dtpToDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpToDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpToDate.Location = New System.Drawing.Point(153, 40)
        Me.dtpToDate.Name = "dtpToDate"
        Me.dtpToDate.Size = New System.Drawing.Size(153, 21)
        Me.dtpToDate.TabIndex = 129
        '
        'dtpFromDate
        '
        Me.dtpFromDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpFromDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFromDate.Location = New System.Drawing.Point(153, 14)
        Me.dtpFromDate.Name = "dtpFromDate"
        Me.dtpFromDate.Size = New System.Drawing.Size(153, 21)
        Me.dtpFromDate.TabIndex = 127
        '
        'cmbAgent
        '
        Me.cmbAgent.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAgent.FormattingEnabled = True
        Me.cmbAgent.Location = New System.Drawing.Point(152, 94)
        Me.cmbAgent.Name = "cmbAgent"
        Me.cmbAgent.Size = New System.Drawing.Size(220, 21)
        Me.cmbAgent.TabIndex = 126
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(152, 152)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(75, 23)
        Me.btnPrint.TabIndex = 122
        Me.btnPrint.Text = "View"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'cmbItem
        '
        Me.cmbItem.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbItem.FormattingEnabled = True
        Me.cmbItem.Location = New System.Drawing.Point(152, 67)
        Me.cmbItem.Name = "cmbItem"
        Me.cmbItem.Size = New System.Drawing.Size(220, 21)
        Me.cmbItem.TabIndex = 132
        '
        'rdoSum
        '
        Me.rdoSum.AutoSize = True
        Me.rdoSum.Checked = True
        Me.rdoSum.Location = New System.Drawing.Point(75, 128)
        Me.rdoSum.Name = "rdoSum"
        Me.rdoSum.Size = New System.Drawing.Size(69, 17)
        Me.rdoSum.TabIndex = 133
        Me.rdoSum.TabStop = True
        Me.rdoSum.Text = "Summary"
        Me.rdoSum.UseVisualStyleBackColor = True
        '
        'RdoDet
        '
        Me.RdoDet.AutoSize = True
        Me.RdoDet.Location = New System.Drawing.Point(227, 128)
        Me.RdoDet.Name = "RdoDet"
        Me.RdoDet.Size = New System.Drawing.Size(52, 17)
        Me.RdoDet.TabIndex = 134
        Me.RdoDet.Text = "Detail"
        Me.RdoDet.UseVisualStyleBackColor = True
        '
        'GRNReport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(389, 187)
        Me.Controls.Add(Me.RdoDet)
        Me.Controls.Add(Me.rdoSum)
        Me.Controls.Add(Me.cmbItem)
        Me.Controls.Add(Label4)
        Me.Controls.Add(Me.dtpToDate)
        Me.Controls.Add(Label3)
        Me.Controls.Add(Me.dtpFromDate)
        Me.Controls.Add(Label1)
        Me.Controls.Add(Me.cmbAgent)
        Me.Controls.Add(Label2)
        Me.Controls.Add(Me.btnPrint)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "GRNReport"
        Me.Text = "GRN Report"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dtpToDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpFromDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents cmbAgent As System.Windows.Forms.ComboBox
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents cmbItem As System.Windows.Forms.ComboBox
    Friend WithEvents rdoSum As System.Windows.Forms.RadioButton
    Friend WithEvents RdoDet As System.Windows.Forms.RadioButton
End Class
