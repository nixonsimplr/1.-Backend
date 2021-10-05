<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SystemSettings
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
        Dim Label2 As System.Windows.Forms.Label
        Dim Label3 As System.Windows.Forms.Label
        Dim Label4 As System.Windows.Forms.Label
        Dim Label5 As System.Windows.Forms.Label
        Dim Label6 As System.Windows.Forms.Label
        Dim Label7 As System.Windows.Forms.Label
        Dim Label8 As System.Windows.Forms.Label
        Dim Label9 As System.Windows.Forms.Label
        Dim Label10 As System.Windows.Forms.Label
        Dim Label11 As System.Windows.Forms.Label
        Dim Label12 As System.Windows.Forms.Label
        
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.pnlLogin = New System.Windows.Forms.Panel
        Me.txtDefCategory = New System.Windows.Forms.TextBox
        Me.txtExchangePer = New System.Windows.Forms.TextBox
        Me.txtBadLocation = New System.Windows.Forms.TextBox
        Me.txtPrinterName = New System.Windows.Forms.TextBox
        Me.btnDefTerms = New System.Windows.Forms.Button
        Me.txtDefItemTerms = New System.Windows.Forms.TextBox
        Me.txtDefPrGroup = New System.Windows.Forms.TextBox
        Me.btnDefPrGroup = New System.Windows.Forms.Button
        Me.txtMaxDiscount = New System.Windows.Forms.TextBox
        Me.txtAutoMember = New System.Windows.Forms.TextBox
        Me.chkLGST = New System.Windows.Forms.CheckBox
        Me.cmbSalesType = New System.Windows.Forms.ComboBox
        Me.txtPrGroup = New System.Windows.Forms.TextBox
        Me.btnPrGroup = New System.Windows.Forms.Button
        Me.btnMemItemNo = New System.Windows.Forms.Button
        Me.txtSchAmt = New System.Windows.Forms.TextBox
        Me.cmbHisUnit = New System.Windows.Forms.ComboBox
        Me.txtBankAC = New System.Windows.Forms.TextBox
        Me.txtHisNo = New System.Windows.Forms.TextBox
        Me.txtTail4 = New System.Windows.Forms.TextBox
        Me.txtTail3 = New System.Windows.Forms.TextBox
        Me.txtTail2 = New System.Windows.Forms.TextBox
        Me.txtTail1 = New System.Windows.Forms.TextBox
        Me.txtHeader4 = New System.Windows.Forms.TextBox
        Me.txtHeader3 = New System.Windows.Forms.TextBox
        Me.txtHeader2 = New System.Windows.Forms.TextBox
        Me.txtGST = New System.Windows.Forms.TextBox
        Me.cmbMDT = New System.Windows.Forms.ComboBox
        Me.txtComName = New System.Windows.Forms.TextBox
        Me.txtMaxInv = New System.Windows.Forms.TextBox
        Me.txtHeader1 = New System.Windows.Forms.TextBox
        Label1 = New System.Windows.Forms.Label
        Label2 = New System.Windows.Forms.Label
        Label3 = New System.Windows.Forms.Label
        Label4 = New System.Windows.Forms.Label
        Label5 = New System.Windows.Forms.Label
        Label6 = New System.Windows.Forms.Label
        Label7 = New System.Windows.Forms.Label
        Label8 = New System.Windows.Forms.Label
        Label9 = New System.Windows.Forms.Label
        Label10 = New System.Windows.Forms.Label
        Label11 = New System.Windows.Forms.Label
        Label12 = New System.Windows.Forms.Label
        SysSet_lblMemItemNo = New System.Windows.Forms.Label
        SysSet_lblBankAcnt = New System.Windows.Forms.Label
        SysSet_lblHistNo = New System.Windows.Forms.Label
        SysSet_lblHistUnit = New System.Windows.Forms.Label
        SysSet_lblFooter4 = New System.Windows.Forms.Label
        SysSet_lblFooter3 = New System.Windows.Forms.Label
        SysSet_lblFooter2 = New System.Windows.Forms.Label
        SysSet_lblFooter1 = New System.Windows.Forms.Label
        SysSet_lblHeader4 = New System.Windows.Forms.Label
        SysSet_lblHeader3 = New System.Windows.Forms.Label
        SysSet_lblHeader2 = New System.Windows.Forms.Label
        SysSet_lblMax = New System.Windows.Forms.Label
        SysSet_lblCompName = New System.Windows.Forms.Label
        SysSet_lblHeader1 = New System.Windows.Forms.Label
        SysSet_lblGST = New System.Windows.Forms.Label
        SysSet_lblMdtNo = New System.Windows.Forms.Label
        Me.Panel1.SuspendLayout()
        Me.pnlLogin.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Label1.AutoSize = True
        Label1.Location = New System.Drawing.Point(7, 446)
        Label1.Name = "Label1"
        Label1.Size = New System.Drawing.Size(229, 13)
        Label1.TabIndex = 154
        Label1.Text = "Trade Deal Price Group . . . . . . . . . . . . . . . ."
        '
        'Label2
        '
        Label2.AutoSize = True
        Label2.Location = New System.Drawing.Point(7, 505)
        Label2.Name = "Label2"
        Label2.Size = New System.Drawing.Size(171, 13)
        Label2.TabIndex = 158
        Label2.Text = "Sales Type . . . . . . . . . . . . . . . ."
        AddHandler Label2.Click, AddressOf Me.Label2_Click
        '
        'Label3
        '
        Label3.AutoSize = True
        Label3.Location = New System.Drawing.Point(356, 637)
        Label3.Name = "Label3"
        Label3.Size = New System.Drawing.Size(170, 13)
        Label3.TabIndex = 160
        Label3.Text = "GST Included for Loyalty  . . . . . ."
        Label3.Visible = False
        '
        'Label4
        '
        Label4.AutoSize = True
        Label4.Location = New System.Drawing.Point(7, 674)
        Label4.Name = "Label4"
        Label4.Size = New System.Drawing.Size(183, 13)
        Label4.TabIndex = 163
        Label4.Text = "Auto Member . . . . . . . . . . . . . . . ."
        '
        'Label5
        '
        Label5.AutoSize = True
        Label5.Location = New System.Drawing.Point(326, 678)
        Label5.Name = "Label5"
        Label5.Size = New System.Drawing.Size(102, 13)
        Label5.TabIndex = 164
        Label5.Text = "(Minimum Purchase)"
        '
        'Label6
        '
        Label6.AutoSize = True
        Label6.Location = New System.Drawing.Point(7, 564)
        Label6.Name = "Label6"
        Label6.Size = New System.Drawing.Size(232, 13)
        Label6.TabIndex = 166
        Label6.Text = "Last Import Date / Time . . . . . . . . . . . . . . . ."
        '
        'Label7
        '
        Label7.AutoSize = True
        Label7.Location = New System.Drawing.Point(7, 475)
        Label7.Name = "Label7"
        Label7.Size = New System.Drawing.Size(212, 13)
        Label7.TabIndex = 167
        Label7.Text = "Default Price Group . . . . . . . . . . . . . . . ."
        '
        'Label8
        '
        Label8.AutoSize = True
        Label8.Location = New System.Drawing.Point(7, 535)
        Label8.Name = "Label8"
        Label8.Size = New System.Drawing.Size(211, 13)
        Label8.TabIndex = 171
        Label8.Text = "Default Item Terms . . . . . . . . . . . . . . . ."
        '
        'Label9
        '
        Label9.AutoSize = True
        Label9.Location = New System.Drawing.Point(7, 620)
        Label9.Name = "Label9"
        Label9.Size = New System.Drawing.Size(181, 13)
        Label9.TabIndex = 173
        Label9.Text = "Printer Name . . . . . . . . . . . . . . . ."
        '
        'Label10
        '
        Label10.AutoSize = True
        Label10.Location = New System.Drawing.Point(7, 363)
        Label10.Name = "Label10"
        Label10.Size = New System.Drawing.Size(180, 13)
        Label10.TabIndex = 176
        Label10.Text = "Bad Location . . . . . . . . . . . . . . . ."
        '
        'Label11
        '
        Label11.AutoSize = True
        Label11.Location = New System.Drawing.Point(7, 391)
        Label11.Name = "Label11"
        Label11.Size = New System.Drawing.Size(224, 13)
        Label11.TabIndex = 178
        Label11.Text = "Exchange Percentage . . . . . . . . . . . . . . . ."
        '
        'Label12
        '
        Label12.AutoSize = True
        Label12.Location = New System.Drawing.Point(7, 418)
        Label12.Name = "Label12"
        Label12.Size = New System.Drawing.Size(202, 13)
        Label12.TabIndex = 180
        Label12.Text = "Default Category . . . . . . . . . . . . . . . ."
        AddHandler Label12.Click, AddressOf Me.Label12_Click
        '
        'SysSet_lblMemItemNo
        '
        SysSet_lblMemItemNo.AutoSize = True
        SysSet_lblMemItemNo.Location = New System.Drawing.Point(7, 647)
        SysSet_lblMemItemNo.Name = "SysSet_lblMemItemNo"
        SysSet_lblMemItemNo.Size = New System.Drawing.Size(163, 13)
        SysSet_lblMemItemNo.TabIndex = 135
        SysSet_lblMemItemNo.Text = "Admin Charges . . . . . . . . . . . ."
        '
        'SysSet_lblBankAcnt
        '
        SysSet_lblBankAcnt.AutoSize = True
        SysSet_lblBankAcnt.Location = New System.Drawing.Point(7, 593)
        SysSet_lblBankAcnt.Name = "SysSet_lblBankAcnt"
        SysSet_lblBankAcnt.Size = New System.Drawing.Size(184, 13)
        SysSet_lblBankAcnt.TabIndex = 131
        SysSet_lblBankAcnt.Text = "Bank Account . . . . . . . . . . . . . . . ."
        '
        'SysSet_lblHistNo
        '
        SysSet_lblHistNo.AutoSize = True
        SysSet_lblHistNo.Location = New System.Drawing.Point(7, 336)
        SysSet_lblHistNo.Name = "SysSet_lblHistNo"
        SysSet_lblHistNo.Size = New System.Drawing.Size(169, 13)
        SysSet_lblHistNo.TabIndex = 129
        SysSet_lblHistNo.Text = "History No . . . . . . . . . . . . . . . ."
        '
        'SysSet_lblHistUnit
        '
        SysSet_lblHistUnit.AutoSize = True
        SysSet_lblHistUnit.Location = New System.Drawing.Point(7, 309)
        SysSet_lblHistUnit.Name = "SysSet_lblHistUnit"
        SysSet_lblHistUnit.Size = New System.Drawing.Size(175, 13)
        SysSet_lblHistUnit.TabIndex = 127
        SysSet_lblHistUnit.Text = "History Unit . . . . . . . . . . . . . . . ."
        '
        'SysSet_lblFooter4
        '
        SysSet_lblFooter4.AutoSize = True
        SysSet_lblFooter4.Location = New System.Drawing.Point(7, 282)
        SysSet_lblFooter4.Name = "SysSet_lblFooter4"
        SysSet_lblFooter4.Size = New System.Drawing.Size(171, 13)
        SysSet_lblFooter4.TabIndex = 125
        SysSet_lblFooter4.Text = "Footer4 . . . . . . . . . . . . . . . . . ."
        '
        'SysSet_lblFooter3
        '
        SysSet_lblFooter3.AutoSize = True
        SysSet_lblFooter3.Location = New System.Drawing.Point(7, 256)
        SysSet_lblFooter3.Name = "SysSet_lblFooter3"
        SysSet_lblFooter3.Size = New System.Drawing.Size(178, 13)
        SysSet_lblFooter3.TabIndex = 123
        SysSet_lblFooter3.Text = "Footer3 . . . . . . . . . . . . . . . . . . ."
        '
        'SysSet_lblFooter2
        '
        SysSet_lblFooter2.AutoSize = True
        SysSet_lblFooter2.Location = New System.Drawing.Point(7, 229)
        SysSet_lblFooter2.Name = "SysSet_lblFooter2"
        SysSet_lblFooter2.Size = New System.Drawing.Size(185, 13)
        SysSet_lblFooter2.TabIndex = 121
        SysSet_lblFooter2.Text = "Footer2 . . . . . . . . . . . . . . . . . . . ."
        '
        'SysSet_lblFooter1
        '
        SysSet_lblFooter1.AutoSize = True
        SysSet_lblFooter1.Location = New System.Drawing.Point(7, 202)
        SysSet_lblFooter1.Name = "SysSet_lblFooter1"
        SysSet_lblFooter1.Size = New System.Drawing.Size(171, 13)
        SysSet_lblFooter1.TabIndex = 119
        SysSet_lblFooter1.Text = "Footer1 . . . . . . . . . . . . . . . . . ."
        '
        'SysSet_lblHeader4
        '
        SysSet_lblHeader4.AutoSize = True
        SysSet_lblHeader4.Location = New System.Drawing.Point(7, 175)
        SysSet_lblHeader4.Name = "SysSet_lblHeader4"
        SysSet_lblHeader4.Size = New System.Drawing.Size(160, 13)
        SysSet_lblHeader4.TabIndex = 117
        SysSet_lblHeader4.Text = "Header4 . . . . . . . . . . . . . . . ."
        '
        'SysSet_lblHeader3
        '
        SysSet_lblHeader3.AutoSize = True
        SysSet_lblHeader3.Location = New System.Drawing.Point(7, 148)
        SysSet_lblHeader3.Name = "SysSet_lblHeader3"
        SysSet_lblHeader3.Size = New System.Drawing.Size(160, 13)
        SysSet_lblHeader3.TabIndex = 115
        SysSet_lblHeader3.Text = "Header3 . . . . . . . . . . . . . . . ."
        '
        'SysSet_lblHeader2
        '
        SysSet_lblHeader2.AutoSize = True
        SysSet_lblHeader2.Location = New System.Drawing.Point(7, 121)
        SysSet_lblHeader2.Name = "SysSet_lblHeader2"
        SysSet_lblHeader2.Size = New System.Drawing.Size(160, 13)
        SysSet_lblHeader2.TabIndex = 113
        SysSet_lblHeader2.Text = "Header2 . . . . . . . . . . . . . . . ."
        '
        'SysSet_lblMax
        '
        SysSet_lblMax.AutoSize = True
        SysSet_lblMax.Location = New System.Drawing.Point(7, 41)
        SysSet_lblMax.Name = "SysSet_lblMax"
        SysSet_lblMax.Size = New System.Drawing.Size(174, 13)
        SysSet_lblMax.TabIndex = 103
        SysSet_lblMax.Text = "Max. Invoice Print . . . . . . . . . . . "
        '
        'SysSet_lblCompName
        '
        SysSet_lblCompName.AutoSize = True
        SysSet_lblCompName.Location = New System.Drawing.Point(7, 67)
        SysSet_lblCompName.Name = "SysSet_lblCompName"
        SysSet_lblCompName.Size = New System.Drawing.Size(166, 13)
        SysSet_lblCompName.TabIndex = 104
        SysSet_lblCompName.Text = "Company Name . . . . . . . . . . . ."
        '
        'SysSet_lblHeader1
        '
        SysSet_lblHeader1.AutoSize = True
        SysSet_lblHeader1.Location = New System.Drawing.Point(7, 94)
        SysSet_lblHeader1.Name = "SysSet_lblHeader1"
        SysSet_lblHeader1.Size = New System.Drawing.Size(160, 13)
        SysSet_lblHeader1.TabIndex = 107
        SysSet_lblHeader1.Text = "Header1 . . . . . . . . . . . . . . . ."
        '
        'SysSet_lblGST
        '
        SysSet_lblGST.AutoSize = True
        SysSet_lblGST.Location = New System.Drawing.Point(7, 19)
        SysSet_lblGST.Name = "SysSet_lblGST"
        SysSet_lblGST.Size = New System.Drawing.Size(166, 13)
        SysSet_lblGST.TabIndex = 101
        SysSet_lblGST.Text = "GST . . . . . . . . . . . . . . . . . . . ."
        '
        'SysSet_lblMdtNo
        '
        SysSet_lblMdtNo.AutoSize = True
        SysSet_lblMdtNo.Location = New System.Drawing.Point(295, 34)
        SysSet_lblMdtNo.Name = "SysSet_lblMdtNo"
        SysSet_lblMdtNo.Size = New System.Drawing.Size(163, 13)
        SysSet_lblMdtNo.TabIndex = 111
        SysSet_lblMdtNo.Text = "MDT No . . . . . . . . . . . . . . . . ."
        SysSet_lblMdtNo.Visible = False
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnCancel)
        Me.Panel1.Controls.Add(Me.btnSave)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 706)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(575, 39)
        Me.Panel1.TabIndex = 3
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(501, 8)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(62, 23)
        Me.btnCancel.TabIndex = 8
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(433, 8)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(62, 23)
        Me.btnSave.TabIndex = 7
        Me.btnSave.Text = "&Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'pnlLogin
        '
        Me.pnlLogin.Controls.Add(Me.txtDefCategory)
        Me.pnlLogin.Controls.Add(Label12)
        Me.pnlLogin.Controls.Add(Me.txtExchangePer)
        Me.pnlLogin.Controls.Add(Label11)
        Me.pnlLogin.Controls.Add(Me.txtBadLocation)
        Me.pnlLogin.Controls.Add(Label10)
        Me.pnlLogin.Controls.Add(Me.txtPrinterName)
        Me.pnlLogin.Controls.Add(Label9)
        Me.pnlLogin.Controls.Add(Me.btnDefTerms)
        Me.pnlLogin.Controls.Add(Me.txtDefItemTerms)
        Me.pnlLogin.Controls.Add(Label8)
        Me.pnlLogin.Controls.Add(Me.txtDefPrGroup)
        Me.pnlLogin.Controls.Add(Me.btnDefPrGroup)
        Me.pnlLogin.Controls.Add(Label7)
        Me.pnlLogin.Controls.Add(Me.txtMaxDiscount)
        Me.pnlLogin.Controls.Add(Label6)
        Me.pnlLogin.Controls.Add(Label5)
        Me.pnlLogin.Controls.Add(Me.txtAutoMember)
        Me.pnlLogin.Controls.Add(Label4)
        Me.pnlLogin.Controls.Add(Me.chkLGST)
        Me.pnlLogin.Controls.Add(Label3)
        Me.pnlLogin.Controls.Add(Me.cmbSalesType)
        Me.pnlLogin.Controls.Add(Label2)
        Me.pnlLogin.Controls.Add(Me.txtPrGroup)
        Me.pnlLogin.Controls.Add(Me.btnPrGroup)
        Me.pnlLogin.Controls.Add(Label1)
        Me.pnlLogin.Controls.Add(Me.btnMemItemNo)
        Me.pnlLogin.Controls.Add(Me.txtSchAmt)
        Me.pnlLogin.Controls.Add(SysSet_lblMemItemNo)
        Me.pnlLogin.Controls.Add(Me.cmbHisUnit)
        Me.pnlLogin.Controls.Add(Me.txtBankAC)
        Me.pnlLogin.Controls.Add(SysSet_lblBankAcnt)
        Me.pnlLogin.Controls.Add(Me.txtHisNo)
        Me.pnlLogin.Controls.Add(SysSet_lblHistNo)
        Me.pnlLogin.Controls.Add(SysSet_lblHistUnit)
        Me.pnlLogin.Controls.Add(Me.txtTail4)
        Me.pnlLogin.Controls.Add(SysSet_lblFooter4)
        Me.pnlLogin.Controls.Add(Me.txtTail3)
        Me.pnlLogin.Controls.Add(SysSet_lblFooter3)
        Me.pnlLogin.Controls.Add(Me.txtTail2)
        Me.pnlLogin.Controls.Add(SysSet_lblFooter2)
        Me.pnlLogin.Controls.Add(Me.txtTail1)
        Me.pnlLogin.Controls.Add(SysSet_lblFooter1)
        Me.pnlLogin.Controls.Add(Me.txtHeader4)
        Me.pnlLogin.Controls.Add(SysSet_lblHeader4)
        Me.pnlLogin.Controls.Add(Me.txtHeader3)
        Me.pnlLogin.Controls.Add(SysSet_lblHeader3)
        Me.pnlLogin.Controls.Add(Me.txtHeader2)
        Me.pnlLogin.Controls.Add(SysSet_lblHeader2)
        Me.pnlLogin.Controls.Add(Me.txtGST)
        Me.pnlLogin.Controls.Add(Me.cmbMDT)
        Me.pnlLogin.Controls.Add(Me.txtComName)
        Me.pnlLogin.Controls.Add(Me.txtMaxInv)
        Me.pnlLogin.Controls.Add(SysSet_lblMax)
        Me.pnlLogin.Controls.Add(SysSet_lblCompName)
        Me.pnlLogin.Controls.Add(Me.txtHeader1)
        Me.pnlLogin.Controls.Add(SysSet_lblHeader1)
        Me.pnlLogin.Controls.Add(SysSet_lblGST)
        Me.pnlLogin.Controls.Add(SysSet_lblMdtNo)
        Me.pnlLogin.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlLogin.Location = New System.Drawing.Point(0, 0)
        Me.pnlLogin.Name = "pnlLogin"
        Me.pnlLogin.Size = New System.Drawing.Size(575, 706)
        Me.pnlLogin.TabIndex = 2
        '
        'txtDefCategory
        '
        Me.txtDefCategory.Location = New System.Drawing.Point(155, 418)
        Me.txtDefCategory.Name = "txtDefCategory"
        Me.txtDefCategory.Size = New System.Drawing.Size(165, 21)
        Me.txtDefCategory.TabIndex = 179
        '
        'txtExchangePer
        '
        Me.txtExchangePer.Location = New System.Drawing.Point(155, 391)
        Me.txtExchangePer.Name = "txtExchangePer"
        Me.txtExchangePer.Size = New System.Drawing.Size(165, 21)
        Me.txtExchangePer.TabIndex = 177
        '
        'txtBadLocation
        '
        Me.txtBadLocation.Location = New System.Drawing.Point(155, 363)
        Me.txtBadLocation.Name = "txtBadLocation"
        Me.txtBadLocation.Size = New System.Drawing.Size(165, 21)
        Me.txtBadLocation.TabIndex = 175
        '
        'txtPrinterName
        '
        Me.txtPrinterName.Location = New System.Drawing.Point(156, 620)
        Me.txtPrinterName.Name = "txtPrinterName"
        Me.txtPrinterName.Size = New System.Drawing.Size(165, 21)
        Me.txtPrinterName.TabIndex = 174
        '
        'btnDefTerms
        '
        Me.btnDefTerms.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDefTerms.Location = New System.Drawing.Point(327, 535)
        Me.btnDefTerms.Name = "btnDefTerms"
        Me.btnDefTerms.Size = New System.Drawing.Size(24, 23)
        Me.btnDefTerms.TabIndex = 172
        Me.btnDefTerms.Text = "..."
        Me.btnDefTerms.UseVisualStyleBackColor = True
        '
        'txtDefItemTerms
        '
        Me.txtDefItemTerms.Location = New System.Drawing.Point(156, 535)
        Me.txtDefItemTerms.Name = "txtDefItemTerms"
        Me.txtDefItemTerms.Size = New System.Drawing.Size(165, 21)
        Me.txtDefItemTerms.TabIndex = 170
        '
        'txtDefPrGroup
        '
        Me.txtDefPrGroup.Location = New System.Drawing.Point(155, 475)
        Me.txtDefPrGroup.Name = "txtDefPrGroup"
        Me.txtDefPrGroup.Size = New System.Drawing.Size(165, 21)
        Me.txtDefPrGroup.TabIndex = 169
        '
        'btnDefPrGroup
        '
        Me.btnDefPrGroup.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDefPrGroup.Location = New System.Drawing.Point(326, 475)
        Me.btnDefPrGroup.Name = "btnDefPrGroup"
        Me.btnDefPrGroup.Size = New System.Drawing.Size(24, 23)
        Me.btnDefPrGroup.TabIndex = 168
        Me.btnDefPrGroup.Text = "..."
        Me.btnDefPrGroup.UseVisualStyleBackColor = True
        '
        'txtMaxDiscount
        '
        Me.txtMaxDiscount.Location = New System.Drawing.Point(156, 564)
        Me.txtMaxDiscount.Name = "txtMaxDiscount"
        Me.txtMaxDiscount.Size = New System.Drawing.Size(165, 21)
        Me.txtMaxDiscount.TabIndex = 165
        '
        'txtAutoMember
        '
        Me.txtAutoMember.Location = New System.Drawing.Point(155, 674)
        Me.txtAutoMember.Name = "txtAutoMember"
        Me.txtAutoMember.Size = New System.Drawing.Size(165, 21)
        Me.txtAutoMember.TabIndex = 162
        '
        'chkLGST
        '
        Me.chkLGST.AutoSize = True
        Me.chkLGST.Location = New System.Drawing.Point(514, 636)
        Me.chkLGST.Name = "chkLGST"
        Me.chkLGST.Size = New System.Drawing.Size(15, 14)
        Me.chkLGST.TabIndex = 161
        Me.chkLGST.UseVisualStyleBackColor = True
        Me.chkLGST.Visible = False
        '
        'cmbSalesType
        '
        Me.cmbSalesType.FormattingEnabled = True
        Me.cmbSalesType.Location = New System.Drawing.Point(156, 505)
        Me.cmbSalesType.Name = "cmbSalesType"
        Me.cmbSalesType.Size = New System.Drawing.Size(165, 21)
        Me.cmbSalesType.TabIndex = 159
        '
        'txtPrGroup
        '
        Me.txtPrGroup.Location = New System.Drawing.Point(155, 446)
        Me.txtPrGroup.Name = "txtPrGroup"
        Me.txtPrGroup.Size = New System.Drawing.Size(165, 21)
        Me.txtPrGroup.TabIndex = 157
        '
        'btnPrGroup
        '
        Me.btnPrGroup.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrGroup.Location = New System.Drawing.Point(326, 446)
        Me.btnPrGroup.Name = "btnPrGroup"
        Me.btnPrGroup.Size = New System.Drawing.Size(24, 23)
        Me.btnPrGroup.TabIndex = 156
        Me.btnPrGroup.Text = "..."
        Me.btnPrGroup.UseVisualStyleBackColor = True
        '
        'btnMemItemNo
        '
        Me.btnMemItemNo.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnMemItemNo.Location = New System.Drawing.Point(326, 647)
        Me.btnMemItemNo.Name = "btnMemItemNo"
        Me.btnMemItemNo.Size = New System.Drawing.Size(24, 23)
        Me.btnMemItemNo.TabIndex = 152
        Me.btnMemItemNo.Text = "..."
        Me.btnMemItemNo.UseVisualStyleBackColor = True
        Me.btnMemItemNo.Visible = False
        '
        'txtSchAmt
        '
        Me.txtSchAmt.Location = New System.Drawing.Point(155, 647)
        Me.txtSchAmt.Name = "txtSchAmt"
        Me.txtSchAmt.Size = New System.Drawing.Size(165, 21)
        Me.txtSchAmt.TabIndex = 134
        '
        'cmbHisUnit
        '
        Me.cmbHisUnit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbHisUnit.FormattingEnabled = True
        Me.cmbHisUnit.Items.AddRange(New Object() {"Month"})
        Me.cmbHisUnit.Location = New System.Drawing.Point(155, 309)
        Me.cmbHisUnit.Name = "cmbHisUnit"
        Me.cmbHisUnit.Size = New System.Drawing.Size(165, 21)
        Me.cmbHisUnit.TabIndex = 133
        '
        'txtBankAC
        '
        Me.txtBankAC.Location = New System.Drawing.Point(156, 593)
        Me.txtBankAC.Name = "txtBankAC"
        Me.txtBankAC.Size = New System.Drawing.Size(165, 21)
        Me.txtBankAC.TabIndex = 130
        '
        'txtHisNo
        '
        Me.txtHisNo.Location = New System.Drawing.Point(155, 336)
        Me.txtHisNo.Name = "txtHisNo"
        Me.txtHisNo.Size = New System.Drawing.Size(165, 21)
        Me.txtHisNo.TabIndex = 128
        '
        'txtTail4
        '
        Me.txtTail4.Location = New System.Drawing.Point(155, 282)
        Me.txtTail4.Name = "txtTail4"
        Me.txtTail4.Size = New System.Drawing.Size(165, 21)
        Me.txtTail4.TabIndex = 124
        '
        'txtTail3
        '
        Me.txtTail3.Location = New System.Drawing.Point(155, 255)
        Me.txtTail3.Name = "txtTail3"
        Me.txtTail3.Size = New System.Drawing.Size(165, 21)
        Me.txtTail3.TabIndex = 122
        '
        'txtTail2
        '
        Me.txtTail2.Location = New System.Drawing.Point(155, 228)
        Me.txtTail2.Name = "txtTail2"
        Me.txtTail2.Size = New System.Drawing.Size(165, 21)
        Me.txtTail2.TabIndex = 120
        '
        'txtTail1
        '
        Me.txtTail1.Location = New System.Drawing.Point(155, 201)
        Me.txtTail1.Name = "txtTail1"
        Me.txtTail1.Size = New System.Drawing.Size(165, 21)
        Me.txtTail1.TabIndex = 118
        '
        'txtHeader4
        '
        Me.txtHeader4.Location = New System.Drawing.Point(155, 174)
        Me.txtHeader4.Name = "txtHeader4"
        Me.txtHeader4.Size = New System.Drawing.Size(165, 21)
        Me.txtHeader4.TabIndex = 116
        '
        'txtHeader3
        '
        Me.txtHeader3.Location = New System.Drawing.Point(155, 147)
        Me.txtHeader3.Name = "txtHeader3"
        Me.txtHeader3.Size = New System.Drawing.Size(165, 21)
        Me.txtHeader3.TabIndex = 114
        '
        'txtHeader2
        '
        Me.txtHeader2.Location = New System.Drawing.Point(155, 120)
        Me.txtHeader2.Name = "txtHeader2"
        Me.txtHeader2.Size = New System.Drawing.Size(165, 21)
        Me.txtHeader2.TabIndex = 112
        '
        'txtGST
        '
        Me.txtGST.Location = New System.Drawing.Point(155, 14)
        Me.txtGST.Name = "txtGST"
        Me.txtGST.Size = New System.Drawing.Size(96, 21)
        Me.txtGST.TabIndex = 0
        '
        'cmbMDT
        '
        Me.cmbMDT.FormattingEnabled = True
        Me.cmbMDT.Location = New System.Drawing.Point(449, 26)
        Me.cmbMDT.Name = "cmbMDT"
        Me.cmbMDT.Size = New System.Drawing.Size(96, 21)
        Me.cmbMDT.TabIndex = 5
        Me.cmbMDT.Visible = False
        '
        'txtComName
        '
        Me.txtComName.Location = New System.Drawing.Point(155, 66)
        Me.txtComName.Name = "txtComName"
        Me.txtComName.Size = New System.Drawing.Size(346, 21)
        Me.txtComName.TabIndex = 2
        '
        'txtMaxInv
        '
        Me.txtMaxInv.Location = New System.Drawing.Point(155, 40)
        Me.txtMaxInv.Name = "txtMaxInv"
        Me.txtMaxInv.Size = New System.Drawing.Size(96, 21)
        Me.txtMaxInv.TabIndex = 1
        '
        'txtHeader1
        '
        Me.txtHeader1.Location = New System.Drawing.Point(155, 93)
        Me.txtHeader1.Name = "txtHeader1"
        Me.txtHeader1.Size = New System.Drawing.Size(165, 21)
        Me.txtHeader1.TabIndex = 3
        '
        'SystemSettings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(591, 743)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.pnlLogin)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "SystemSettings"
        Me.Text = "System Settings"
        Me.Panel1.ResumeLayout(False)
        Me.pnlLogin.ResumeLayout(False)
        Me.pnlLogin.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents pnlLogin As System.Windows.Forms.Panel
    Friend WithEvents txtGST As System.Windows.Forms.TextBox
    Friend WithEvents txtComName As System.Windows.Forms.TextBox
    Friend WithEvents txtMaxInv As System.Windows.Forms.TextBox
    Friend WithEvents txtHeader1 As System.Windows.Forms.TextBox
    Friend WithEvents cmbMDT As System.Windows.Forms.ComboBox
    Friend WithEvents txtTail3 As System.Windows.Forms.TextBox
    Friend WithEvents txtTail2 As System.Windows.Forms.TextBox
    Friend WithEvents txtTail1 As System.Windows.Forms.TextBox
    Friend WithEvents txtHeader4 As System.Windows.Forms.TextBox
    Friend WithEvents txtHeader3 As System.Windows.Forms.TextBox
    Friend WithEvents txtHeader2 As System.Windows.Forms.TextBox
    Friend WithEvents txtBankAC As System.Windows.Forms.TextBox
    Friend WithEvents txtHisNo As System.Windows.Forms.TextBox
    Friend WithEvents cmbHisUnit As System.Windows.Forms.ComboBox
    Friend WithEvents txtTail4 As System.Windows.Forms.TextBox
    Friend WithEvents txtSchAmt As System.Windows.Forms.TextBox
    Friend WithEvents btnMemItemNo As System.Windows.Forms.Button
    Friend WithEvents txtPrGroup As System.Windows.Forms.TextBox
    Friend WithEvents btnPrGroup As System.Windows.Forms.Button
    Friend WithEvents cmbSalesType As System.Windows.Forms.ComboBox
    Friend WithEvents chkLGST As System.Windows.Forms.CheckBox
    Friend WithEvents txtAutoMember As System.Windows.Forms.TextBox
    Friend WithEvents txtMaxDiscount As System.Windows.Forms.TextBox

    ''''
    'Friend WithEvents SysSet_lblMemItemNo As System.Windows.Forms.Label
    'Friend WithEvents SysSet_lblBankAcnt As System.Windows.Forms.Label
    'Friend WithEvents SysSet_lblHistNo As System.Windows.Forms.Label
    'Friend WithEvents SysSet_lblHistUnit As System.Windows.Forms.Label
    'Friend WithEvents SysSet_lblFooter4 As System.Windows.Forms.Label
    'Friend WithEvents SysSet_lblFooter3 As System.Windows.Forms.Label
    'Friend WithEvents SysSet_lblFooter2 As System.Windows.Forms.Label
    'Friend WithEvents SysSet_lblFooter1 As System.Windows.Forms.Label
    'Friend WithEvents SysSet_lblHeader4 As System.Windows.Forms.Label
    'Friend WithEvents SysSet_lblHeader3 As System.Windows.Forms.Label
    'Friend WithEvents SysSet_lblHeader2 As System.Windows.Forms.Label
    'Friend WithEvents SysSet_lblMax As System.Windows.Forms.Label
    'Friend WithEvents SysSet_lblCompName As System.Windows.Forms.Label
    'Friend WithEvents SysSet_lblHeader1 As System.Windows.Forms.Label
    'Friend WithEvents SysSet_lblGST As System.Windows.Forms.Label
    'Friend WithEvents SysSet_lblMdtNo As System.Windows.Forms.Label
    ''''
    Friend WithEvents txtDefPrGroup As System.Windows.Forms.TextBox
    Friend WithEvents btnDefPrGroup As System.Windows.Forms.Button
    Friend WithEvents txtDefItemTerms As System.Windows.Forms.TextBox
    Friend WithEvents btnDefTerms As System.Windows.Forms.Button
    Friend WithEvents txtPrinterName As System.Windows.Forms.TextBox
    Friend WithEvents txtBadLocation As System.Windows.Forms.TextBox
    Friend WithEvents txtExchangePer As System.Windows.Forms.TextBox
    Friend WithEvents txtDefCategory As System.Windows.Forms.TextBox
    Friend WithEvents SysSet_lblMemItemNo As System.Windows.Forms.Label
    Friend WithEvents SysSet_lblBankAcnt As System.Windows.Forms.Label
    Friend WithEvents SysSet_lblHistNo As System.Windows.Forms.Label
    Friend WithEvents SysSet_lblHistUnit As System.Windows.Forms.Label
    Friend WithEvents SysSet_lblFooter4 As System.Windows.Forms.Label
    Friend WithEvents SysSet_lblFooter3 As System.Windows.Forms.Label
    Friend WithEvents SysSet_lblFooter2 As System.Windows.Forms.Label
    Friend WithEvents SysSet_lblFooter1 As System.Windows.Forms.Label
    Friend WithEvents SysSet_lblHeader4 As System.Windows.Forms.Label
    Friend WithEvents SysSet_lblHeader3 As System.Windows.Forms.Label
    Friend WithEvents SysSet_lblHeader2 As System.Windows.Forms.Label
    Friend WithEvents SysSet_lblMax As System.Windows.Forms.Label
    Friend WithEvents SysSet_lblCompName As System.Windows.Forms.Label
    Friend WithEvents SysSet_lblHeader1 As System.Windows.Forms.Label
    Friend WithEvents SysSet_lblGST As System.Windows.Forms.Label
    Friend WithEvents SysSet_lblMdtNo As System.Windows.Forms.Label
End Class
