<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DisplayOrder
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DisplayOrder))
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.PnlButton = New System.Windows.Forms.Panel
        Me.btnDOWN = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnUP = New System.Windows.Forms.Button
        Me.ItemBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.SalesDataSet3 = New SalesPromo.SalesDataSet3
        Me.ItemTableAdapter = New SalesPromo.SalesDataSet3TableAdapters.ItemTableAdapter
        Me.ItemNoDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ShortDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DescriptionDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PnlButton.SuspendLayout()
        CType(Me.ItemBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SalesDataSet3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AutoGenerateColumns = False
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ItemNoDataGridViewTextBoxColumn, Me.ShortDesc, Me.DescriptionDataGridViewTextBoxColumn})
        Me.dgvItem.DataSource = Me.ItemBindingSource
        Me.dgvItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItem.Location = New System.Drawing.Point(0, 0)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvItem.Size = New System.Drawing.Size(598, 353)
        Me.dgvItem.TabIndex = 90
        '
        'PnlButton
        '
        Me.PnlButton.Controls.Add(Me.btnDOWN)
        Me.PnlButton.Controls.Add(Me.btnSave)
        Me.PnlButton.Controls.Add(Me.btnUP)
        Me.PnlButton.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PnlButton.Location = New System.Drawing.Point(0, 353)
        Me.PnlButton.Name = "PnlButton"
        Me.PnlButton.Size = New System.Drawing.Size(598, 40)
        Me.PnlButton.TabIndex = 91
        '
        'btnDOWN
        '
        Me.btnDOWN.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDOWN.Image = CType(resources.GetObject("btnDOWN.Image"), System.Drawing.Image)
        Me.btnDOWN.Location = New System.Drawing.Point(494, 8)
        Me.btnDOWN.Name = "btnDOWN"
        Me.btnDOWN.Size = New System.Drawing.Size(24, 23)
        Me.btnDOWN.TabIndex = 79
        Me.btnDOWN.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(519, 8)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(68, 23)
        Me.btnSave.TabIndex = 76
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUP.Image = CType(resources.GetObject("btnUP.Image"), System.Drawing.Image)
        Me.btnUP.Location = New System.Drawing.Point(469, 8)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(25, 23)
        Me.btnUP.TabIndex = 75
        Me.btnUP.UseVisualStyleBackColor = True
        '
        'ItemBindingSource
        '
        Me.ItemBindingSource.DataMember = "Item"
        Me.ItemBindingSource.DataSource = Me.SalesDataSet3
        '
        'SalesDataSet3
        '
        Me.SalesDataSet3.DataSetName = "SalesDataSet3"
        Me.SalesDataSet3.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'ItemTableAdapter
        '
        Me.ItemTableAdapter.ClearBeforeFill = True
        '
        'ItemNoDataGridViewTextBoxColumn
        '
        Me.ItemNoDataGridViewTextBoxColumn.DataPropertyName = "ItemNo"
        Me.ItemNoDataGridViewTextBoxColumn.HeaderText = "ItemNo"
        Me.ItemNoDataGridViewTextBoxColumn.Name = "ItemNoDataGridViewTextBoxColumn"
        Me.ItemNoDataGridViewTextBoxColumn.ReadOnly = True
        '
        'ShortDesc
        '
        Me.ShortDesc.DataPropertyName = "ShortDesc"
        Me.ShortDesc.HeaderText = "ShortDesc"
        Me.ShortDesc.Name = "ShortDesc"
        Me.ShortDesc.ReadOnly = True
        '
        'DescriptionDataGridViewTextBoxColumn
        '
        Me.DescriptionDataGridViewTextBoxColumn.DataPropertyName = "Description"
        Me.DescriptionDataGridViewTextBoxColumn.HeaderText = "Description"
        Me.DescriptionDataGridViewTextBoxColumn.Name = "DescriptionDataGridViewTextBoxColumn"
        Me.DescriptionDataGridViewTextBoxColumn.ReadOnly = True
        Me.DescriptionDataGridViewTextBoxColumn.Width = 350
        '
        'DisplayOrder
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(598, 393)
        Me.Controls.Add(Me.dgvItem)
        Me.Controls.Add(Me.PnlButton)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "DisplayOrder"
        Me.Text = "Display Order"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PnlButton.ResumeLayout(False)
        CType(Me.ItemBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SalesDataSet3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents PnlButton As System.Windows.Forms.Panel
    Friend WithEvents btnDOWN As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents SalesDataSet3 As SalesPromo.SalesDataSet3
    Friend WithEvents ItemBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents ItemTableAdapter As SalesPromo.SalesDataSet3TableAdapters.ItemTableAdapter
    Friend WithEvents ItemNoDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ShortDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DescriptionDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
