<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ProductListToPrint
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
        Dim Label4 As System.Windows.Forms.Label
        Dim Label3 As System.Windows.Forms.Label
        Dim Label1 As System.Windows.Forms.Label
        Me.dgvOrdItem = New System.Windows.Forms.DataGridView
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.txtStartPosition = New System.Windows.Forms.TextBox
        Me.btnUnselect = New System.Windows.Forms.Button
        Me.btnSelect = New System.Windows.Forms.Button
        Me.btnView = New System.Windows.Forms.Button
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.btnApply = New System.Windows.Forms.Button
        Me.cbSupplier = New System.Windows.Forms.ComboBox
        Me.cbCategory = New System.Windows.Forms.ComboBox
        Me.Check = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.CustNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CustName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ShipName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.AgentID = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Barcode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CompanyName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Label4 = New System.Windows.Forms.Label
        Label3 = New System.Windows.Forms.Label
        Label1 = New System.Windows.Forms.Label
        CType(Me.dgvOrdItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label4
        '
        Label4.AutoSize = True
        Label4.Location = New System.Drawing.Point(12, 47)
        Label4.Name = "Label4"
        Label4.Size = New System.Drawing.Size(151, 13)
        Label4.TabIndex = 118
        Label4.Text = "Supplier Code . . . . . . . . . . . . ."
        '
        'Label3
        '
        Label3.AutoSize = True
        Label3.Location = New System.Drawing.Point(12, 18)
        Label3.Name = "Label3"
        Label3.Size = New System.Drawing.Size(127, 13)
        Label3.TabIndex = 116
        Label3.Text = "Category . . . . . . . . . . . . ."
        '
        'Label1
        '
        Label1.AutoSize = True
        Label1.Location = New System.Drawing.Point(374, 15)
        Label1.Name = "Label1"
        Label1.Size = New System.Drawing.Size(161, 13)
        Label1.TabIndex = 121
        Label1.Text = "Starting Position . . . . . . . . . . . . ."
        '
        'dgvOrdItem
        '
        Me.dgvOrdItem.AllowUserToAddRows = False
        Me.dgvOrdItem.AllowUserToDeleteRows = False
        Me.dgvOrdItem.AllowUserToOrderColumns = True
        Me.dgvOrdItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvOrdItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Check, Me.CustNo, Me.CustName, Me.ShipName, Me.AgentID, Me.Barcode, Me.CompanyName})
        Me.dgvOrdItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvOrdItem.Location = New System.Drawing.Point(0, 76)
        Me.dgvOrdItem.Name = "dgvOrdItem"
        Me.dgvOrdItem.Size = New System.Drawing.Size(686, 196)
        Me.dgvOrdItem.TabIndex = 16
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.txtStartPosition)
        Me.Panel1.Controls.Add(Label1)
        Me.Panel1.Controls.Add(Me.btnUnselect)
        Me.Panel1.Controls.Add(Me.btnSelect)
        Me.Panel1.Controls.Add(Me.btnView)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 272)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(686, 37)
        Me.Panel1.TabIndex = 15
        '
        'txtStartPosition
        '
        Me.txtStartPosition.Location = New System.Drawing.Point(498, 8)
        Me.txtStartPosition.Name = "txtStartPosition"
        Me.txtStartPosition.Size = New System.Drawing.Size(85, 20)
        Me.txtStartPosition.TabIndex = 122
        Me.txtStartPosition.Text = "1"
        '
        'btnUnselect
        '
        Me.btnUnselect.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUnselect.Location = New System.Drawing.Point(111, 7)
        Me.btnUnselect.Name = "btnUnselect"
        Me.btnUnselect.Size = New System.Drawing.Size(93, 23)
        Me.btnUnselect.TabIndex = 2
        Me.btnUnselect.Text = "Unselect All"
        Me.btnUnselect.UseVisualStyleBackColor = True
        '
        'btnSelect
        '
        Me.btnSelect.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSelect.Location = New System.Drawing.Point(12, 6)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(93, 23)
        Me.btnSelect.TabIndex = 1
        Me.btnSelect.Text = "Select All"
        Me.btnSelect.UseVisualStyleBackColor = True
        '
        'btnView
        '
        Me.btnView.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnView.Location = New System.Drawing.Point(589, 7)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(93, 23)
        Me.btnView.TabIndex = 0
        Me.btnView.Text = "&View"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.btnApply)
        Me.Panel2.Controls.Add(Me.cbSupplier)
        Me.Panel2.Controls.Add(Label4)
        Me.Panel2.Controls.Add(Me.cbCategory)
        Me.Panel2.Controls.Add(Label3)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(686, 76)
        Me.Panel2.TabIndex = 17
        '
        'btnApply
        '
        Me.btnApply.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnApply.Location = New System.Drawing.Point(589, 42)
        Me.btnApply.Name = "btnApply"
        Me.btnApply.Size = New System.Drawing.Size(93, 23)
        Me.btnApply.TabIndex = 1
        Me.btnApply.Text = "&Apply"
        Me.btnApply.UseVisualStyleBackColor = True
        '
        'cbSupplier
        '
        Me.cbSupplier.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbSupplier.FormattingEnabled = True
        Me.cbSupplier.Location = New System.Drawing.Point(142, 39)
        Me.cbSupplier.Name = "cbSupplier"
        Me.cbSupplier.Size = New System.Drawing.Size(224, 21)
        Me.cbSupplier.TabIndex = 119
        '
        'cbCategory
        '
        Me.cbCategory.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbCategory.FormattingEnabled = True
        Me.cbCategory.Location = New System.Drawing.Point(142, 10)
        Me.cbCategory.Name = "cbCategory"
        Me.cbCategory.Size = New System.Drawing.Size(224, 21)
        Me.cbCategory.TabIndex = 117
        '
        'Check
        '
        Me.Check.HeaderText = ""
        Me.Check.Name = "Check"
        Me.Check.Width = 30
        '
        'CustNo
        '
        Me.CustNo.HeaderText = "Item No"
        Me.CustNo.Name = "CustNo"
        Me.CustNo.ReadOnly = True
        Me.CustNo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'CustName
        '
        Me.CustName.HeaderText = "Item Name"
        Me.CustName.Name = "CustName"
        Me.CustName.ReadOnly = True
        Me.CustName.Width = 200
        '
        'ShipName
        '
        Me.ShipName.HeaderText = "Base UOM"
        Me.ShipName.Name = "ShipName"
        Me.ShipName.ReadOnly = True
        Me.ShipName.Width = 200
        '
        'AgentID
        '
        Me.AgentID.HeaderText = "Price"
        Me.AgentID.Name = "AgentID"
        Me.AgentID.ReadOnly = True
        '
        'Barcode
        '
        Me.Barcode.HeaderText = "Barcode"
        Me.Barcode.Name = "Barcode"
        Me.Barcode.ReadOnly = True
        '
        'CompanyName
        '
        Me.CompanyName.HeaderText = "Company"
        Me.CompanyName.Name = "CompanyName"
        Me.CompanyName.ReadOnly = True
        '
        'ProductListToPrint
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(686, 309)
        Me.Controls.Add(Me.dgvOrdItem)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "ProductListToPrint"
        Me.Text = "Product List To Print"
        CType(Me.dgvOrdItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvOrdItem As System.Windows.Forms.DataGridView
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents cbSupplier As System.Windows.Forms.ComboBox
    Friend WithEvents cbCategory As System.Windows.Forms.ComboBox
    Friend WithEvents btnApply As System.Windows.Forms.Button
    Friend WithEvents btnUnselect As System.Windows.Forms.Button
    Friend WithEvents btnSelect As System.Windows.Forms.Button
    Friend WithEvents txtStartPosition As System.Windows.Forms.TextBox
    Friend WithEvents Check As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents CustNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CustName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ShipName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents AgentID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Barcode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CompanyName As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
