<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UserSettings
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
        Me.components = New System.ComponentModel.Container
        Dim Label1 As System.Windows.Forms.Label
        Dim Label2 As System.Windows.Forms.Label

        Dim Label3 As System.Windows.Forms.Label
        Me.pnlLogin = New System.Windows.Forms.Panel
        Me.txtSalesTarget = New System.Windows.Forms.TextBox
        Me.cmbCommCode = New System.Windows.Forms.ComboBox
        Me.chkActive = New System.Windows.Forms.CheckBox
        Me.txtUserCode = New System.Windows.Forms.TextBox
        Me.txtSName = New System.Windows.Forms.TextBox
        Me.cmbUserType = New System.Windows.Forms.ComboBox
        Me.txtPassword = New System.Windows.Forms.TextBox
        Me.txtUserName = New System.Windows.Forms.TextBox
        Me.txtRePassword = New System.Windows.Forms.TextBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.SalesDataSet1 = New SalesPromo.SalesDataSet1
        Me.SalesAgentBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.SalesAgentTableAdapter = New SalesPromo.SalesDataSet1TableAdapters.SalesAgentTableAdapter
        Me.txtCurTarget = New System.Windows.Forms.TextBox
        Label1 = New System.Windows.Forms.Label
        Label2 = New System.Windows.Forms.Label
        UserSet_lblUserCode = New System.Windows.Forms.Label
        UserSet_lblPassword = New System.Windows.Forms.Label
        UserSet_lblRePassword = New System.Windows.Forms.Label
        UserSet_lblUserType = New System.Windows.Forms.Label
        UserSet_lblAccessLevel = New System.Windows.Forms.Label
        UserSet_lblUserName = New System.Windows.Forms.Label
        Label3 = New System.Windows.Forms.Label
        Me.pnlLogin.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.SalesDataSet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SalesAgentBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Label1.AutoSize = True
        Label1.Location = New System.Drawing.Point(11, 197)
        Label1.Name = "Label1"
        Label1.Size = New System.Drawing.Size(174, 13)
        Label1.TabIndex = 153
        Label1.Text = "Commission Code . . . . . . . . . . . ."
        '
        'Label2
        '
        Label2.AutoSize = True
        Label2.Location = New System.Drawing.Point(12, 230)
        Label2.Name = "Label2"
        Label2.Size = New System.Drawing.Size(198, 13)
        Label2.TabIndex = 155
        Label2.Text = "Actual Sales Target . . . . . . . . . . . . . ."
        '
        'pnlLogin
        '
        Me.pnlLogin.Controls.Add(Me.txtCurTarget)
        Me.pnlLogin.Controls.Add(Label3)
        Me.pnlLogin.Controls.Add(Me.txtSalesTarget)
        Me.pnlLogin.Controls.Add(Me.cmbCommCode)
        Me.pnlLogin.Controls.Add(Me.chkActive)
        Me.pnlLogin.Controls.Add(Me.txtUserCode)
        Me.pnlLogin.Controls.Add(Me.txtSName)
        Me.pnlLogin.Controls.Add(Me.cmbUserType)
        Me.pnlLogin.Controls.Add(Me.txtPassword)
        Me.pnlLogin.Controls.Add(Me.txtUserName)
        Me.pnlLogin.Controls.Add(Me.txtRePassword)
        Me.pnlLogin.Controls.Add(Label2)
        Me.pnlLogin.Controls.Add(Label1)
        Me.pnlLogin.Controls.Add(UserSet_lblUserCode)
        Me.pnlLogin.Controls.Add(UserSet_lblPassword)
        Me.pnlLogin.Controls.Add(UserSet_lblRePassword)
        Me.pnlLogin.Controls.Add(UserSet_lblUserType)
        Me.pnlLogin.Controls.Add(UserSet_lblAccessLevel)
        Me.pnlLogin.Controls.Add(UserSet_lblUserName)
        Me.pnlLogin.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlLogin.Location = New System.Drawing.Point(0, 0)
        Me.pnlLogin.Name = "pnlLogin"
        Me.pnlLogin.Size = New System.Drawing.Size(371, 315)
        Me.pnlLogin.TabIndex = 0
        '
        'txtSalesTarget
        '
        Me.txtSalesTarget.Location = New System.Drawing.Point(160, 222)
        Me.txtSalesTarget.Name = "txtSalesTarget"
        Me.txtSalesTarget.Size = New System.Drawing.Size(165, 21)
        Me.txtSalesTarget.TabIndex = 156
        '
        'cmbCommCode
        '
        Me.cmbCommCode.FormattingEnabled = True
        Me.cmbCommCode.Location = New System.Drawing.Point(161, 192)
        Me.cmbCommCode.Name = "cmbCommCode"
        Me.cmbCommCode.Size = New System.Drawing.Size(164, 21)
        Me.cmbCommCode.TabIndex = 154
        '
        'chkActive
        '
        Me.chkActive.AutoSize = True
        Me.chkActive.Location = New System.Drawing.Point(12, 291)
        Me.chkActive.Name = "chkActive"
        Me.chkActive.Size = New System.Drawing.Size(56, 17)
        Me.chkActive.TabIndex = 6
        Me.chkActive.Text = "Active"
        Me.chkActive.UseVisualStyleBackColor = True
        '
        'txtUserCode
        '
        Me.txtUserCode.Location = New System.Drawing.Point(160, 12)
        Me.txtUserCode.Name = "txtUserCode"
        Me.txtUserCode.Size = New System.Drawing.Size(165, 21)
        Me.txtUserCode.TabIndex = 0
        '
        'txtSName
        '
        Me.txtSName.Location = New System.Drawing.Point(160, 71)
        Me.txtSName.Name = "txtSName"
        Me.txtSName.Size = New System.Drawing.Size(165, 21)
        Me.txtSName.TabIndex = 4
        '
        'cmbUserType
        '
        Me.cmbUserType.FormattingEnabled = True
        Me.cmbUserType.Items.AddRange(New Object() {"MDT User", "Server User"})
        Me.cmbUserType.Location = New System.Drawing.Point(160, 160)
        Me.cmbUserType.Name = "cmbUserType"
        Me.cmbUserType.Size = New System.Drawing.Size(165, 21)
        Me.cmbUserType.TabIndex = 5
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(160, 99)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(165, 21)
        Me.txtPassword.TabIndex = 2
        '
        'txtUserName
        '
        Me.txtUserName.Location = New System.Drawing.Point(160, 42)
        Me.txtUserName.Name = "txtUserName"
        Me.txtUserName.Size = New System.Drawing.Size(165, 21)
        Me.txtUserName.TabIndex = 1
        '
        'txtRePassword
        '
        Me.txtRePassword.Location = New System.Drawing.Point(160, 129)
        Me.txtRePassword.Name = "txtRePassword"
        Me.txtRePassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtRePassword.Size = New System.Drawing.Size(165, 21)
        Me.txtRePassword.TabIndex = 3
        '
        'UserSet_lblUserCode
        '
        UserSet_lblUserCode.AutoSize = True
        UserSet_lblUserCode.Location = New System.Drawing.Point(11, 15)
        UserSet_lblUserCode.Name = "UserSet_lblUserCode"
        UserSet_lblUserCode.Size = New System.Drawing.Size(148, 13)
        UserSet_lblUserCode.TabIndex = 111
        UserSet_lblUserCode.Text = "User Code . . . . . . . . . . . . ."
        '
        'UserSet_lblPassword
        '
        UserSet_lblPassword.AutoSize = True
        UserSet_lblPassword.Location = New System.Drawing.Point(11, 104)
        UserSet_lblPassword.Name = "UserSet_lblPassword"
        UserSet_lblPassword.Size = New System.Drawing.Size(151, 13)
        UserSet_lblPassword.TabIndex = 103
        UserSet_lblPassword.Text = "Password . . . . . . . . . . . . . ."
        '
        'UserSet_lblRePassword
        '
        UserSet_lblRePassword.AutoSize = True
        UserSet_lblRePassword.Location = New System.Drawing.Point(11, 132)
        UserSet_lblRePassword.Name = "UserSet_lblRePassword"
        UserSet_lblRePassword.Size = New System.Drawing.Size(148, 13)
        UserSet_lblRePassword.TabIndex = 104
        UserSet_lblRePassword.Text = "Re-enter Password . . . . . . ."
        '
        'UserSet_lblUserType
        '
        UserSet_lblUserType.AutoSize = True
        UserSet_lblUserType.Location = New System.Drawing.Point(11, 164)
        UserSet_lblUserType.Name = "UserSet_lblUserType"
        UserSet_lblUserType.Size = New System.Drawing.Size(147, 13)
        UserSet_lblUserType.TabIndex = 106
        UserSet_lblUserType.Text = "User Type . . . . . . . . . . . . ."
        '
        'UserSet_lblAccessLevel
        '
        UserSet_lblAccessLevel.AutoSize = True
        UserSet_lblAccessLevel.Location = New System.Drawing.Point(11, 75)
        UserSet_lblAccessLevel.Name = "UserSet_lblAccessLevel"
        UserSet_lblAccessLevel.Size = New System.Drawing.Size(151, 13)
        UserSet_lblAccessLevel.TabIndex = 107
        UserSet_lblAccessLevel.Text = "User ID  . . . . . . . . . . . . . . ."
        '
        'UserSet_lblUserName
        '
        UserSet_lblUserName.AutoSize = True
        UserSet_lblUserName.Location = New System.Drawing.Point(11, 45)
        UserSet_lblUserName.Name = "UserSet_lblUserName"
        UserSet_lblUserName.Size = New System.Drawing.Size(150, 13)
        UserSet_lblUserName.TabIndex = 101
        UserSet_lblUserName.Text = "User Name . . . . . . . . . . . . ."
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnDelete)
        Me.Panel1.Controls.Add(Me.btnNew)
        Me.Panel1.Controls.Add(Me.btnCancel)
        Me.Panel1.Controls.Add(Me.btnSave)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 314)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(371, 42)
        Me.Panel1.TabIndex = 1
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDelete.Location = New System.Drawing.Point(229, 7)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(62, 23)
        Me.btnDelete.TabIndex = 10
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnNew.Location = New System.Drawing.Point(93, 7)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(62, 23)
        Me.btnNew.TabIndex = 9
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(297, 7)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(62, 23)
        Me.btnCancel.TabIndex = 8
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(161, 7)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(62, 23)
        Me.btnSave.TabIndex = 7
        Me.btnSave.Text = "&Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'SalesDataSet1
        '
        Me.SalesDataSet1.DataSetName = "SalesDataSet1"
        Me.SalesDataSet1.EnforceConstraints = False
        Me.SalesDataSet1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'SalesAgentBindingSource
        '
        Me.SalesAgentBindingSource.DataMember = "SalesAgent"
        Me.SalesAgentBindingSource.DataSource = Me.SalesDataSet1
        '
        'SalesAgentTableAdapter
        '
        Me.SalesAgentTableAdapter.ClearBeforeFill = True
        '
        'txtCurTarget
        '
        Me.txtCurTarget.Location = New System.Drawing.Point(161, 253)
        Me.txtCurTarget.Name = "txtCurTarget"
        Me.txtCurTarget.Size = New System.Drawing.Size(165, 21)
        Me.txtCurTarget.TabIndex = 158
        '
        'Label3
        '
        Label3.AutoSize = True
        Label3.Location = New System.Drawing.Point(13, 261)
        Label3.Name = "Label3"
        Label3.Size = New System.Drawing.Size(212, 13)
        Label3.TabIndex = 157
        Label3.Text = "Achieved Sales Target . . . . . . . . . . . . . ."
        '
        'UserSettings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(371, 356)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.pnlLogin)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "UserSettings"
        Me.Text = "User Administration"
        Me.pnlLogin.ResumeLayout(False)
        Me.pnlLogin.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        CType(Me.SalesDataSet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SalesAgentBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlLogin As System.Windows.Forms.Panel
    Friend WithEvents txtSName As System.Windows.Forms.TextBox
    Friend WithEvents cmbUserType As System.Windows.Forms.ComboBox
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents txtUserName As System.Windows.Forms.TextBox
    Friend WithEvents txtRePassword As System.Windows.Forms.TextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents txtUserCode As System.Windows.Forms.TextBox
    Friend WithEvents chkActive As System.Windows.Forms.CheckBox
    Friend WithEvents SalesDataSet1 As SalesPromo.SalesDataSet1
    Friend WithEvents SalesAgentBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents SalesAgentTableAdapter As SalesPromo.SalesDataSet1TableAdapters.SalesAgentTableAdapter
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents cmbCommCode As System.Windows.Forms.ComboBox
    Friend WithEvents txtSalesTarget As System.Windows.Forms.TextBox
    Friend WithEvents txtCurTarget As System.Windows.Forms.TextBox
    Friend WithEvents UserSet_lblUserCode As System.Windows.Forms.Label
    Friend WithEvents UserSet_lblPassword As System.Windows.Forms.Label
    Friend WithEvents UserSet_lblRePassword As System.Windows.Forms.Label
    Friend WithEvents UserSet_lblUserType As System.Windows.Forms.Label
    Friend WithEvents UserSet_lblAccessLevel As System.Windows.Forms.Label
    Friend WithEvents UserSet_lblUserName As System.Windows.Forms.Label
End Class
