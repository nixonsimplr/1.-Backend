<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DailySalesInv
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
        Dim Label20 As System.Windows.Forms.Label
        Dim Label1 As System.Windows.Forms.Label
        Dim Label3 As System.Windows.Forms.Label
        Dim Label4 As System.Windows.Forms.Label
        Me.cmbAgent = New System.Windows.Forms.ComboBox
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.btnPrint = New System.Windows.Forms.Button
        Me.cmbCust = New System.Windows.Forms.ComboBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.rbAmt = New System.Windows.Forms.RadioButton
        Me.rbQty = New System.Windows.Forms.RadioButton
        Label2 = New System.Windows.Forms.Label
        Label20 = New System.Windows.Forms.Label
        Label1 = New System.Windows.Forms.Label
        Label3 = New System.Windows.Forms.Label
        Label4 = New System.Windows.Forms.Label
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label2
        '
        Label2.AutoSize = True
        Label2.Location = New System.Drawing.Point(11, 45)
        Label2.Name = "Label2"
        Label2.Size = New System.Drawing.Size(153, 13)
        Label2.TabIndex = 102
        Label2.Text = "Agent No . . . . . . . . . . . . . . "
        '
        'Label20
        '
        Label20.AutoSize = True
        Label20.Location = New System.Drawing.Point(11, 43)
        Label20.Name = "Label20"
        Label20.Size = New System.Drawing.Size(155, 13)
        Label20.TabIndex = 99
        Label20.Text = "Terms. . . . . . . . . . . . . . . . . "
        Label20.Visible = False
        '
        'Label1
        '
        Label1.AutoSize = True
        Label1.Location = New System.Drawing.Point(11, 14)
        Label1.Name = "Label1"
        Label1.Size = New System.Drawing.Size(156, 13)
        Label1.TabIndex = 101
        Label1.Text = "Date. . . . . . . . . . . . . . . . . . "
        '
        'Label3
        '
        Label3.AutoSize = True
        Label3.Location = New System.Drawing.Point(11, 75)
        Label3.Name = "Label3"
        Label3.Size = New System.Drawing.Size(154, 13)
        Label3.TabIndex = 104
        Label3.Text = "Customer . . . . . . . . . . . . . . "
        '
        'Label4
        '
        Label4.AutoSize = True
        Label4.Location = New System.Drawing.Point(13, 105)
        Label4.Name = "Label4"
        Label4.Size = New System.Drawing.Size(134, 13)
        Label4.TabIndex = 106
        Label4.Text = "Value . . . . . . . . . . . . . . "
        '
        'cmbAgent
        '
        Me.cmbAgent.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbAgent.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAgent.FormattingEnabled = True
        Me.cmbAgent.Location = New System.Drawing.Point(155, 40)
        Me.cmbAgent.Name = "cmbAgent"
        Me.cmbAgent.Size = New System.Drawing.Size(257, 21)
        Me.cmbAgent.TabIndex = 103
        '
        'dtpDate
        '
        Me.dtpDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(155, 10)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(153, 21)
        Me.dtpDate.TabIndex = 98
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.btnPrint.Location = New System.Drawing.Point(169, 133)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(75, 23)
        Me.btnPrint.TabIndex = 97
        Me.btnPrint.Text = "View"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'cmbCust
        '
        Me.cmbCust.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbCust.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCust.FormattingEnabled = True
        Me.cmbCust.Location = New System.Drawing.Point(155, 70)
        Me.cmbCust.Name = "cmbCust"
        Me.cmbCust.Size = New System.Drawing.Size(257, 21)
        Me.cmbCust.TabIndex = 105
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.rbAmt)
        Me.Panel1.Controls.Add(Me.rbQty)
        Me.Panel1.Location = New System.Drawing.Point(155, 98)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(158, 29)
        Me.Panel1.TabIndex = 107
        '
        'rbAmt
        '
        Me.rbAmt.AutoSize = True
        Me.rbAmt.Location = New System.Drawing.Point(77, 4)
        Me.rbAmt.Name = "rbAmt"
        Me.rbAmt.Size = New System.Drawing.Size(62, 17)
        Me.rbAmt.TabIndex = 1
        Me.rbAmt.Text = "Amount"
        Me.rbAmt.UseVisualStyleBackColor = True
        '
        'rbQty
        '
        Me.rbQty.AutoSize = True
        Me.rbQty.Checked = True
        Me.rbQty.Location = New System.Drawing.Point(4, 4)
        Me.rbQty.Name = "rbQty"
        Me.rbQty.Size = New System.Drawing.Size(67, 17)
        Me.rbQty.TabIndex = 0
        Me.rbQty.TabStop = True
        Me.rbQty.Text = "Quantity"
        Me.rbQty.UseVisualStyleBackColor = True
        '
        'DailySalesInv
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(429, 168)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Label4)
        Me.Controls.Add(Me.cmbCust)
        Me.Controls.Add(Label3)
        Me.Controls.Add(Me.cmbAgent)
        Me.Controls.Add(Label2)
        Me.Controls.Add(Label20)
        Me.Controls.Add(Me.dtpDate)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Label1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "DailySalesInv"
        Me.Text = "Daily Sales"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmbAgent As System.Windows.Forms.ComboBox
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents cmbCust As System.Windows.Forms.ComboBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents rbQty As System.Windows.Forms.RadioButton
    Friend WithEvents rbAmt As System.Windows.Forms.RadioButton
End Class
