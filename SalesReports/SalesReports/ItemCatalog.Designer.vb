<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ItemCatalog
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
        Dim Label3 As System.Windows.Forms.Label
        Dim Label2 As System.Windows.Forms.Label
        Dim Label1 As System.Windows.Forms.Label
        Dim Label4 As System.Windows.Forms.Label
        Dim Label5 As System.Windows.Forms.Label
        Me.btnPrint = New System.Windows.Forms.Button
        Me.cbFromItemNo = New System.Windows.Forms.ComboBox
        Me.cbCategory = New System.Windows.Forms.ComboBox
        Me.cbToItemNo = New System.Windows.Forms.ComboBox
        Me.cbSupplier = New System.Windows.Forms.ComboBox
        Me.cbCompany = New System.Windows.Forms.ComboBox
        Label3 = New System.Windows.Forms.Label
        Label2 = New System.Windows.Forms.Label
        Label1 = New System.Windows.Forms.Label
        Label4 = New System.Windows.Forms.Label
        Label5 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Label3
        '
        Label3.AutoSize = True
        Label3.Location = New System.Drawing.Point(15, 20)
        Label3.Name = "Label3"
        Label3.Size = New System.Drawing.Size(143, 13)
        Label3.TabIndex = 107
        Label3.Text = "Category . . . . . . . . . . . . ."
        '
        'Label2
        '
        Label2.AutoSize = True
        Label2.Location = New System.Drawing.Point(15, 78)
        Label2.Name = "Label2"
        Label2.Size = New System.Drawing.Size(163, 13)
        Label2.TabIndex = 106
        Label2.Text = "From Item No . . . . . . . . . . . . ."
        '
        'Label1
        '
        Label1.AutoSize = True
        Label1.Location = New System.Drawing.Point(15, 109)
        Label1.Name = "Label1"
        Label1.Size = New System.Drawing.Size(151, 13)
        Label1.TabIndex = 112
        Label1.Text = "To Item No . . . . . . . . . . . . ."
        '
        'Label4
        '
        Label4.AutoSize = True
        Label4.Location = New System.Drawing.Point(15, 49)
        Label4.Name = "Label4"
        Label4.Size = New System.Drawing.Size(164, 13)
        Label4.TabIndex = 114
        Label4.Text = "Supplier Code . . . . . . . . . . . . ."
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnPrint.Location = New System.Drawing.Point(104, 167)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(75, 23)
        Me.btnPrint.TabIndex = 110
        Me.btnPrint.Text = "View"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'cbFromItemNo
        '
        Me.cbFromItemNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbFromItemNo.FormattingEnabled = True
        Me.cbFromItemNo.Location = New System.Drawing.Point(145, 70)
        Me.cbFromItemNo.Name = "cbFromItemNo"
        Me.cbFromItemNo.Size = New System.Drawing.Size(152, 21)
        Me.cbFromItemNo.TabIndex = 109
        '
        'cbCategory
        '
        Me.cbCategory.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbCategory.FormattingEnabled = True
        Me.cbCategory.Location = New System.Drawing.Point(145, 12)
        Me.cbCategory.Name = "cbCategory"
        Me.cbCategory.Size = New System.Drawing.Size(152, 21)
        Me.cbCategory.TabIndex = 111
        '
        'cbToItemNo
        '
        Me.cbToItemNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbToItemNo.FormattingEnabled = True
        Me.cbToItemNo.Location = New System.Drawing.Point(145, 101)
        Me.cbToItemNo.Name = "cbToItemNo"
        Me.cbToItemNo.Size = New System.Drawing.Size(152, 21)
        Me.cbToItemNo.TabIndex = 113
        '
        'cbSupplier
        '
        Me.cbSupplier.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbSupplier.FormattingEnabled = True
        Me.cbSupplier.Location = New System.Drawing.Point(145, 41)
        Me.cbSupplier.Name = "cbSupplier"
        Me.cbSupplier.Size = New System.Drawing.Size(152, 21)
        Me.cbSupplier.TabIndex = 115
        '
        'cbCompany
        '
        Me.cbCompany.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbCompany.FormattingEnabled = True
        Me.cbCompany.Location = New System.Drawing.Point(145, 131)
        Me.cbCompany.Name = "cbCompany"
        Me.cbCompany.Size = New System.Drawing.Size(152, 21)
        Me.cbCompany.TabIndex = 117
        '
        'Label5
        '
        Label5.AutoSize = True
        Label5.Location = New System.Drawing.Point(15, 139)
        Label5.Name = "Label5"
        Label5.Size = New System.Drawing.Size(143, 13)
        Label5.TabIndex = 116
        Label5.Text = "Company . . . . . . . . . . . . ."
        '
        'ItemCatalog
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(312, 202)
        Me.Controls.Add(Me.cbCompany)
        Me.Controls.Add(Label5)
        Me.Controls.Add(Me.cbSupplier)
        Me.Controls.Add(Label4)
        Me.Controls.Add(Me.cbToItemNo)
        Me.Controls.Add(Label1)
        Me.Controls.Add(Me.cbCategory)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.cbFromItemNo)
        Me.Controls.Add(Label3)
        Me.Controls.Add(Label2)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "ItemCatalog"
        Me.Text = "Item Catalog"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents cbFromItemNo As System.Windows.Forms.ComboBox
    Friend WithEvents cbCategory As System.Windows.Forms.ComboBox
    Friend WithEvents cbToItemNo As System.Windows.Forms.ComboBox
    Friend WithEvents cbSupplier As System.Windows.Forms.ComboBox
    Friend WithEvents cbCompany As System.Windows.Forms.ComboBox
End Class
