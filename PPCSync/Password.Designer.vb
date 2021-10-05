<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Password
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
        Dim lblPassword As System.Windows.Forms.Label
        Me.SServ_Cngpass_btnOK = New System.Windows.Forms.Button
        Me.txtPassword = New System.Windows.Forms.TextBox
        Me.SServ_Cngpass_btnCancel = New System.Windows.Forms.Button
        Me.lblError = New System.Windows.Forms.Label
        lblPassword = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'lblPassword
        '
        lblPassword.AutoSize = True
        lblPassword.Location = New System.Drawing.Point(12, 15)
        lblPassword.Name = "lblPassword"
        lblPassword.Size = New System.Drawing.Size(151, 13)
        lblPassword.TabIndex = 111
        lblPassword.Text = "Password . . . . . . . . . . . . . ."
        '
        'SServ_Cngpass_btnOK
        '
        Me.SServ_Cngpass_btnOK.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.SServ_Cngpass_btnOK.Location = New System.Drawing.Point(52, 66)
        Me.SServ_Cngpass_btnOK.Name = "SServ_Cngpass_btnOK"
        Me.SServ_Cngpass_btnOK.Size = New System.Drawing.Size(75, 23)
        Me.SServ_Cngpass_btnOK.TabIndex = 109
        Me.SServ_Cngpass_btnOK.Text = "OK"
        Me.SServ_Cngpass_btnOK.UseVisualStyleBackColor = True
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(155, 12)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(63)
        Me.txtPassword.Size = New System.Drawing.Size(136, 21)
        Me.txtPassword.TabIndex = 110
        '
        'SServ_Cngpass_btnCancel
        '
        Me.SServ_Cngpass_btnCancel.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.SServ_Cngpass_btnCancel.Location = New System.Drawing.Point(194, 66)
        Me.SServ_Cngpass_btnCancel.Name = "SServ_Cngpass_btnCancel"
        Me.SServ_Cngpass_btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.SServ_Cngpass_btnCancel.TabIndex = 112
        Me.SServ_Cngpass_btnCancel.Text = "Cancel"
        Me.SServ_Cngpass_btnCancel.UseVisualStyleBackColor = True
        '
        'lblError
        '
        Me.lblError.AutoSize = True
        Me.lblError.ForeColor = System.Drawing.Color.Red
        Me.lblError.Location = New System.Drawing.Point(49, 42)
        Me.lblError.Name = "lblError"
        Me.lblError.Size = New System.Drawing.Size(198, 13)
        Me.lblError.TabIndex = 113
        Me.lblError.Text = "Incorrect Password !!! Please Try Again"
        Me.lblError.Visible = False
        '
        'Password
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(311, 101)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblError)
        Me.Controls.Add(Me.SServ_Cngpass_btnCancel)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.SServ_Cngpass_btnOK)
        Me.Controls.Add(lblPassword)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Password"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Enter Password"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents SServ_Cngpass_btnOK As System.Windows.Forms.Button
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents SServ_Cngpass_btnCancel As System.Windows.Forms.Button
    Friend WithEvents lblError As System.Windows.Forms.Label

End Class
