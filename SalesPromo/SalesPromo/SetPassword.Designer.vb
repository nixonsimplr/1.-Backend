<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SetPassword
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
        Me.cmbModName = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtPassword = New System.Windows.Forms.TextBox
        Me.btnSave = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'cmbModName
        '
        Me.cmbModName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbModName.FormattingEnabled = True
        Me.cmbModName.Items.AddRange(New Object() {"Order", "Payment-MODIFY", "Payment-VOID", "Synchronize", "PDC"})
        Me.cmbModName.Location = New System.Drawing.Point(153, 16)
        Me.cmbModName.Name = "cmbModName"
        Me.cmbModName.Size = New System.Drawing.Size(191, 21)
        Me.cmbModName.TabIndex = 22
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 19)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(222, 13)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "Module Name. . . . . . . . . . . . . . . . . . . . . ."
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 54)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(204, 13)
        Me.Label1.TabIndex = 23
        Me.Label1.Text = "Password. . . . . . . . . . . . . . . . . . . . . ."
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(153, 51)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.Size = New System.Drawing.Size(191, 21)
        Me.txtPassword.TabIndex = 24
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(139, 96)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 25
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'SetPassword
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(353, 130)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmbModName)
        Me.Controls.Add(Me.Label2)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "SetPassword"
        Me.Text = "Set Password"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmbModName As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents btnSave As System.Windows.Forms.Button
End Class
