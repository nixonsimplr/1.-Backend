<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPayment
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
        Me.Cust_Paym_gbCCard = New System.Windows.Forms.GroupBox
        Me.dtpExpiryDate = New System.Windows.Forms.DateTimePicker
        Me.txtCBankName = New System.Windows.Forms.TextBox
        Me.txtCCardNo = New System.Windows.Forms.TextBox
        Me.gbCheque = New System.Windows.Forms.GroupBox
        Me.dtpChqDate = New System.Windows.Forms.DateTimePicker
        Me.txtBankName = New System.Windows.Forms.TextBox
        Me.txtChqNo = New System.Windows.Forms.TextBox
        Me.Cust_Paym_btnCancel = New System.Windows.Forms.Button
        Me.Cust_Paym_btnOK = New System.Windows.Forms.Button
        Me.txtAmt = New System.Windows.Forms.TextBox
        Me.cmbPaymethod = New System.Windows.Forms.ComboBox
        Label4 = New System.Windows.Forms.Label
        Label3 = New System.Windows.Forms.Label
        ExtDocNoLabel = New System.Windows.Forms.Label
        Cust_Paym_lblBankName = New System.Windows.Forms.Label
        Cust_Paym_lblExpDat = New System.Windows.Forms.Label
        Cust_Paym_lblCCNo = New System.Windows.Forms.Label
        Cust_Paym_lblPayMhd = New System.Windows.Forms.Label
        Cust_Paym_lblAmount = New System.Windows.Forms.Label
        Me.Cust_Paym_gbCCard.SuspendLayout()
        Me.gbCheque.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label4
        '
        Label4.AutoSize = True
        Label4.Location = New System.Drawing.Point(6, 84)
        Label4.Name = "Label4"
        Label4.Size = New System.Drawing.Size(134, 13)
        Label4.TabIndex = 41
        Label4.Text = "Bank Name. . . . . . . . . . ."
        '
        'Label3
        '
        Label3.AutoSize = True
        Label3.Location = New System.Drawing.Point(6, 57)
        Label3.Name = "Label3"
        Label3.Size = New System.Drawing.Size(144, 13)
        Label3.TabIndex = 39
        Label3.Text = "Cheque Date. . . . . . . . . . ."
        '
        'ExtDocNoLabel
        '
        ExtDocNoLabel.AutoSize = True
        ExtDocNoLabel.Location = New System.Drawing.Point(6, 28)
        ExtDocNoLabel.Name = "ExtDocNoLabel"
        ExtDocNoLabel.Size = New System.Drawing.Size(134, 13)
        ExtDocNoLabel.TabIndex = 37
        ExtDocNoLabel.Text = "Cheque No. . . . . . . . . . ."
        '
        'Cust_Paym_gbCCard
        '
        Me.Cust_Paym_gbCCard.Controls.Add(Me.dtpExpiryDate)
        Me.Cust_Paym_gbCCard.Controls.Add(Me.txtCBankName)
        Me.Cust_Paym_gbCCard.Controls.Add(Cust_Paym_lblBankName)
        Me.Cust_Paym_gbCCard.Controls.Add(Cust_Paym_lblExpDat)
        Me.Cust_Paym_gbCCard.Controls.Add(Me.txtCCardNo)
        Me.Cust_Paym_gbCCard.Controls.Add(Cust_Paym_lblCCNo)
        Me.Cust_Paym_gbCCard.Location = New System.Drawing.Point(21, 90)
        Me.Cust_Paym_gbCCard.Name = "Cust_Paym_gbCCard"
        Me.Cust_Paym_gbCCard.Size = New System.Drawing.Size(243, 115)
        Me.Cust_Paym_gbCCard.TabIndex = 45
        Me.Cust_Paym_gbCCard.TabStop = False
        Me.Cust_Paym_gbCCard.Text = "Credit Card"
        Me.Cust_Paym_gbCCard.Visible = False
        '
        'dtpExpiryDate
        '
        Me.dtpExpiryDate.CustomFormat = "MM/yyyy"
        Me.dtpExpiryDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpExpiryDate.Location = New System.Drawing.Point(106, 50)
        Me.dtpExpiryDate.Name = "dtpExpiryDate"
        Me.dtpExpiryDate.Size = New System.Drawing.Size(129, 21)
        Me.dtpExpiryDate.TabIndex = 48
        '
        'txtCBankName
        '
        Me.txtCBankName.Location = New System.Drawing.Point(106, 79)
        Me.txtCBankName.Name = "txtCBankName"
        Me.txtCBankName.Size = New System.Drawing.Size(129, 21)
        Me.txtCBankName.TabIndex = 40
        '
        'Cust_Paym_lblBankName
        '
        Cust_Paym_lblBankName.AutoSize = True
        Cust_Paym_lblBankName.Location = New System.Drawing.Point(8, 82)
        Cust_Paym_lblBankName.Name = "Cust_Paym_lblBankName"
        Cust_Paym_lblBankName.Size = New System.Drawing.Size(134, 13)
        Cust_Paym_lblBankName.TabIndex = 47
        Cust_Paym_lblBankName.Text = "Bank Name. . . . . . . . . . ."
        '
        'Cust_Paym_lblExpDat
        '
        Cust_Paym_lblExpDat.AutoSize = True
        Cust_Paym_lblExpDat.Location = New System.Drawing.Point(8, 55)
        Cust_Paym_lblExpDat.Name = "Cust_Paym_lblExpDat"
        Cust_Paym_lblExpDat.Size = New System.Drawing.Size(137, 13)
        Cust_Paym_lblExpDat.TabIndex = 45
        Cust_Paym_lblExpDat.Text = "Expiry Date. . . . . . . . . . ."
        '
        'txtCCardNo
        '
        Me.txtCCardNo.Location = New System.Drawing.Point(106, 22)
        Me.txtCCardNo.Name = "txtCCardNo"
        Me.txtCCardNo.Size = New System.Drawing.Size(129, 21)
        Me.txtCCardNo.TabIndex = 43
        '
        'Cust_Paym_lblCCNo
        '
        Cust_Paym_lblCCNo.AutoSize = True
        Cust_Paym_lblCCNo.Location = New System.Drawing.Point(8, 26)
        Cust_Paym_lblCCNo.Name = "Cust_Paym_lblCCNo"
        Cust_Paym_lblCCNo.Size = New System.Drawing.Size(152, 13)
        Cust_Paym_lblCCNo.TabIndex = 44
        Cust_Paym_lblCCNo.Text = "Credit Card No. . . . . . . . . . ."
        '
        'gbCheque
        '
        Me.gbCheque.Controls.Add(Me.dtpChqDate)
        Me.gbCheque.Controls.Add(Me.txtBankName)
        Me.gbCheque.Controls.Add(Label4)
        Me.gbCheque.Controls.Add(Label3)
        Me.gbCheque.Controls.Add(Me.txtChqNo)
        Me.gbCheque.Controls.Add(ExtDocNoLabel)
        Me.gbCheque.Location = New System.Drawing.Point(25, 97)
        Me.gbCheque.Name = "gbCheque"
        Me.gbCheque.Size = New System.Drawing.Size(243, 115)
        Me.gbCheque.TabIndex = 46
        Me.gbCheque.TabStop = False
        Me.gbCheque.Text = "Cheque"
        Me.gbCheque.Visible = False
        '
        'dtpChqDate
        '
        Me.dtpChqDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpChqDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpChqDate.Location = New System.Drawing.Point(104, 52)
        Me.dtpChqDate.Name = "dtpChqDate"
        Me.dtpChqDate.Size = New System.Drawing.Size(129, 21)
        Me.dtpChqDate.TabIndex = 42
        '
        'txtBankName
        '
        Me.txtBankName.Location = New System.Drawing.Point(104, 81)
        Me.txtBankName.Name = "txtBankName"
        Me.txtBankName.Size = New System.Drawing.Size(129, 21)
        Me.txtBankName.TabIndex = 46
        '
        'txtChqNo
        '
        Me.txtChqNo.Location = New System.Drawing.Point(104, 24)
        Me.txtChqNo.Name = "txtChqNo"
        Me.txtChqNo.Size = New System.Drawing.Size(129, 21)
        Me.txtChqNo.TabIndex = 36
        '
        'Cust_Paym_btnCancel
        '
        Me.Cust_Paym_btnCancel.Location = New System.Drawing.Point(201, 221)
        Me.Cust_Paym_btnCancel.Name = "Cust_Paym_btnCancel"
        Me.Cust_Paym_btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.Cust_Paym_btnCancel.TabIndex = 47
        Me.Cust_Paym_btnCancel.Text = "Cancel"
        Me.Cust_Paym_btnCancel.UseVisualStyleBackColor = True
        '
        'Cust_Paym_btnOK
        '
        Me.Cust_Paym_btnOK.Location = New System.Drawing.Point(120, 221)
        Me.Cust_Paym_btnOK.Name = "Cust_Paym_btnOK"
        Me.Cust_Paym_btnOK.Size = New System.Drawing.Size(75, 23)
        Me.Cust_Paym_btnOK.TabIndex = 48
        Me.Cust_Paym_btnOK.Text = "OK"
        Me.Cust_Paym_btnOK.UseVisualStyleBackColor = True
        '
        'txtAmt
        '
        Me.txtAmt.Location = New System.Drawing.Point(153, 22)
        Me.txtAmt.Name = "txtAmt"
        Me.txtAmt.Size = New System.Drawing.Size(119, 21)
        Me.txtAmt.TabIndex = 50
        '
        'cmbPaymethod
        '
        Me.cmbPaymethod.FormattingEnabled = True
        Me.cmbPaymethod.Location = New System.Drawing.Point(153, 49)
        Me.cmbPaymethod.Name = "cmbPaymethod"
        Me.cmbPaymethod.Size = New System.Drawing.Size(121, 21)
        Me.cmbPaymethod.TabIndex = 51
        '
        'Cust_Paym_lblPayMhd
        '
        Cust_Paym_lblPayMhd.AutoSize = True
        Cust_Paym_lblPayMhd.Location = New System.Drawing.Point(8, 51)
        Cust_Paym_lblPayMhd.Name = "Cust_Paym_lblPayMhd"
        Cust_Paym_lblPayMhd.Size = New System.Drawing.Size(137, 13)
        Cust_Paym_lblPayMhd.TabIndex = 42
        Cust_Paym_lblPayMhd.Text = "Payment Method . . . . . . ."
        '
        'Cust_Paym_lblAmount
        '
        Cust_Paym_lblAmount.AutoSize = True
        Cust_Paym_lblAmount.Location = New System.Drawing.Point(8, 25)
        Cust_Paym_lblAmount.Name = "Cust_Paym_lblAmount"
        Cust_Paym_lblAmount.Size = New System.Drawing.Size(135, 13)
        Cust_Paym_lblAmount.TabIndex = 49
        Cust_Paym_lblAmount.Text = "Amount . . . . . . . . . . . . ."
        '
        'frmPayment
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(292, 256)
        Me.Controls.Add(Me.cmbPaymethod)
        Me.Controls.Add(Me.txtAmt)
        Me.Controls.Add(Cust_Paym_lblAmount)
        Me.Controls.Add(Me.Cust_Paym_btnOK)
        Me.Controls.Add(Me.Cust_Paym_btnCancel)
        Me.Controls.Add(Me.Cust_Paym_gbCCard)
        Me.Controls.Add(Cust_Paym_lblPayMhd)
        Me.Controls.Add(Me.gbCheque)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "frmPayment"
        Me.Text = "Payment"
        Me.Cust_Paym_gbCCard.ResumeLayout(False)
        Me.Cust_Paym_gbCCard.PerformLayout()
        Me.gbCheque.ResumeLayout(False)
        Me.gbCheque.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Cust_Paym_gbCCard As System.Windows.Forms.GroupBox
    Friend WithEvents dtpExpiryDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtCBankName As System.Windows.Forms.TextBox
    Friend WithEvents txtCCardNo As System.Windows.Forms.TextBox
    Friend WithEvents gbCheque As System.Windows.Forms.GroupBox
    Friend WithEvents dtpChqDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtBankName As System.Windows.Forms.TextBox
    Friend WithEvents txtChqNo As System.Windows.Forms.TextBox
    Friend WithEvents Cust_Paym_btnCancel As System.Windows.Forms.Button
    Friend WithEvents Cust_Paym_btnOK As System.Windows.Forms.Button
    Friend WithEvents txtAmt As System.Windows.Forms.TextBox
    Friend WithEvents cmbPaymethod As System.Windows.Forms.ComboBox
    Friend WithEvents Cust_Paym_lblBankName As System.Windows.Forms.Label
    Friend WithEvents Cust_Paym_lblExpDat As System.Windows.Forms.Label
    Friend WithEvents Cust_Paym_lblCCNo As System.Windows.Forms.Label
    Friend WithEvents Cust_Paym_lblPayMhd As System.Windows.Forms.Label
    Friend WithEvents Cust_Paym_lblAmount As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ExtDocNoLabel As System.Windows.Forms.Label
End Class
