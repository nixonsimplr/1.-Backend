<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class VoucherList
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
        Me.dgvVouList = New System.Windows.Forms.DataGridView
        Me.VoucherBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.VoucherDataSet = New SalesPromo.VoucherDataSet
        Me.VoucherTableAdapter = New SalesPromo.VoucherDataSetTableAdapters.VoucherTableAdapter
        Me.VoucherNoDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.VoucherFromNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.VoucherToNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Active = New System.Windows.Forms.DataGridViewCheckBoxColumn
        CType(Me.dgvVouList, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.VoucherBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.VoucherDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvVouList
        '
        Me.dgvVouList.AllowUserToAddRows = False
        Me.dgvVouList.AllowUserToDeleteRows = False
        Me.dgvVouList.AutoGenerateColumns = False
        Me.dgvVouList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvVouList.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.VoucherNoDataGridViewTextBoxColumn, Me.VoucherFromNo, Me.VoucherToNo, Me.Active})
        Me.dgvVouList.DataSource = Me.VoucherBindingSource
        Me.dgvVouList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvVouList.Location = New System.Drawing.Point(0, 0)
        Me.dgvVouList.Name = "dgvVouList"
        Me.dgvVouList.ReadOnly = True
        Me.dgvVouList.Size = New System.Drawing.Size(845, 323)
        Me.dgvVouList.TabIndex = 0
        '
        'VoucherBindingSource
        '
        Me.VoucherBindingSource.DataMember = "Voucher"
        Me.VoucherBindingSource.DataSource = Me.VoucherDataSet
        '
        'VoucherDataSet
        '
        Me.VoucherDataSet.DataSetName = "VoucherDataSet"
        Me.VoucherDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'VoucherTableAdapter
        '
        Me.VoucherTableAdapter.ClearBeforeFill = True
        '
        'VoucherNoDataGridViewTextBoxColumn
        '
        Me.VoucherNoDataGridViewTextBoxColumn.DataPropertyName = "VoucherNo"
        Me.VoucherNoDataGridViewTextBoxColumn.HeaderText = "Voucher No"
        Me.VoucherNoDataGridViewTextBoxColumn.Name = "VoucherNoDataGridViewTextBoxColumn"
        Me.VoucherNoDataGridViewTextBoxColumn.ReadOnly = True
        '
        'VoucherFromNo
        '
        Me.VoucherFromNo.DataPropertyName = "VoucherFromNo"
        Me.VoucherFromNo.HeaderText = "Voucher From"
        Me.VoucherFromNo.Name = "VoucherFromNo"
        Me.VoucherFromNo.ReadOnly = True
        '
        'VoucherToNo
        '
        Me.VoucherToNo.DataPropertyName = "VoucherToNo"
        Me.VoucherToNo.HeaderText = "Voucher To"
        Me.VoucherToNo.Name = "VoucherToNo"
        Me.VoucherToNo.ReadOnly = True
        '
        'Active
        '
        Me.Active.DataPropertyName = "Active"
        Me.Active.HeaderText = "Active"
        Me.Active.Name = "Active"
        Me.Active.ReadOnly = True
        '
        'VoucherList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(845, 323)
        Me.Controls.Add(Me.dgvVouList)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "VoucherList"
        Me.Text = "Voucher List"
        CType(Me.dgvVouList, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.VoucherBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.VoucherDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvVouList As System.Windows.Forms.DataGridView
    Friend WithEvents VoucherDataSet As SalesPromo.VoucherDataSet
    Friend WithEvents VoucherBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents VoucherTableAdapter As SalesPromo.VoucherDataSetTableAdapters.VoucherTableAdapter
    Friend WithEvents VoucherNoDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents VoucherFromNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents VoucherToNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Active As System.Windows.Forms.DataGridViewCheckBoxColumn
End Class
