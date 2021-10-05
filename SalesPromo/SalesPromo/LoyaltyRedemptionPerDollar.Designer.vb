<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class LoyaltyRedemptionPerDollar
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
        Dim lblCustPrice As System.Windows.Forms.Label
        Dim Label1 As System.Windows.Forms.Label
        Me.txtCode = New System.Windows.Forms.TextBox
        Me.btnCustNo = New System.Windows.Forms.Button
        Me.txtPoints = New System.Windows.Forms.TextBox
        Me.btnSave = New System.Windows.Forms.Button
        lblCustPrice = New System.Windows.Forms.Label
        Label1 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'lblCustPrice
        '
        lblCustPrice.AutoSize = True
        lblCustPrice.Location = New System.Drawing.Point(8, 15)
        lblCustPrice.Name = "lblCustPrice"
        lblCustPrice.Size = New System.Drawing.Size(160, 13)
        lblCustPrice.TabIndex = 77
        lblCustPrice.Text = "Member Type. . . . . . . . . . . . ."
        '
        'Label1
        '
        Label1.AutoSize = True
        Label1.Location = New System.Drawing.Point(8, 46)
        Label1.Name = "Label1"
        Label1.Size = New System.Drawing.Size(141, 13)
        Label1.TabIndex = 80
        Label1.Text = "Points (1$=). . . . . . . . . . ."
        '
        'txtCode
        '
        Me.txtCode.Location = New System.Drawing.Point(128, 12)
        Me.txtCode.Name = "txtCode"
        Me.txtCode.Size = New System.Drawing.Size(100, 21)
        Me.txtCode.TabIndex = 78
        '
        'btnCustNo
        '
        Me.btnCustNo.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCustNo.Location = New System.Drawing.Point(228, 11)
        Me.btnCustNo.Name = "btnCustNo"
        Me.btnCustNo.Size = New System.Drawing.Size(24, 23)
        Me.btnCustNo.TabIndex = 79
        Me.btnCustNo.Text = "..."
        Me.btnCustNo.UseVisualStyleBackColor = True
        '
        'txtPoints
        '
        Me.txtPoints.Location = New System.Drawing.Point(128, 43)
        Me.txtPoints.Name = "txtPoints"
        Me.txtPoints.Size = New System.Drawing.Size(100, 21)
        Me.txtPoints.TabIndex = 81
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(89, 85)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(93, 23)
        Me.btnSave.TabIndex = 82
        Me.btnSave.Text = "&Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'LoyaltyRedemptionPerDollar
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(263, 124)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.txtPoints)
        Me.Controls.Add(Label1)
        Me.Controls.Add(Me.txtCode)
        Me.Controls.Add(Me.btnCustNo)
        Me.Controls.Add(lblCustPrice)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "LoyaltyRedemptionPerDollar"
        Me.Text = "Loyalty Redemption Per Dollar"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtCode As System.Windows.Forms.TextBox
    Friend WithEvents btnCustNo As System.Windows.Forms.Button
    Friend WithEvents txtPoints As System.Windows.Forms.TextBox
    Friend WithEvents btnSave As System.Windows.Forms.Button
End Class
