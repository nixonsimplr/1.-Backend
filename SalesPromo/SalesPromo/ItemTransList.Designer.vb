<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ItemTransList
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
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.DocNoDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DocDtDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DocTypeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LocationDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ItemTransBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.SalesDataSet2 = New SalesPromo.SalesDataSet2
        Me.ItemTransTableAdapter = New SalesPromo.SalesDataSet2TableAdapters.ItemTransTableAdapter
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ItemTransBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SalesDataSet2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DocNoDataGridViewTextBoxColumn, Me.DocDtDataGridViewTextBoxColumn, Me.DocTypeDataGridViewTextBoxColumn, Me.LocationDataGridViewTextBoxColumn})
        Me.DataGridView1.DataSource = Me.ItemTransBindingSource
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(0, 0)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(444, 260)
        Me.DataGridView1.TabIndex = 0
        '
        'DocNoDataGridViewTextBoxColumn
        '
        Me.DocNoDataGridViewTextBoxColumn.DataPropertyName = "DocNo"
        Me.DocNoDataGridViewTextBoxColumn.HeaderText = "DocNo"
        Me.DocNoDataGridViewTextBoxColumn.Name = "DocNoDataGridViewTextBoxColumn"
        '
        'DocDtDataGridViewTextBoxColumn
        '
        Me.DocDtDataGridViewTextBoxColumn.DataPropertyName = "DocDt"
        Me.DocDtDataGridViewTextBoxColumn.HeaderText = "DocDt"
        Me.DocDtDataGridViewTextBoxColumn.Name = "DocDtDataGridViewTextBoxColumn"
        '
        'DocTypeDataGridViewTextBoxColumn
        '
        Me.DocTypeDataGridViewTextBoxColumn.DataPropertyName = "DocType"
        Me.DocTypeDataGridViewTextBoxColumn.HeaderText = "DocType"
        Me.DocTypeDataGridViewTextBoxColumn.Name = "DocTypeDataGridViewTextBoxColumn"
        '
        'LocationDataGridViewTextBoxColumn
        '
        Me.LocationDataGridViewTextBoxColumn.DataPropertyName = "Location"
        Me.LocationDataGridViewTextBoxColumn.HeaderText = "Location"
        Me.LocationDataGridViewTextBoxColumn.Name = "LocationDataGridViewTextBoxColumn"
        '
        'ItemTransBindingSource
        '
        Me.ItemTransBindingSource.DataMember = "ItemTrans"
        Me.ItemTransBindingSource.DataSource = Me.SalesDataSet2
        '
        'SalesDataSet2
        '
        Me.SalesDataSet2.DataSetName = "SalesDataSet2"
        Me.SalesDataSet2.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'ItemTransTableAdapter
        '
        Me.ItemTransTableAdapter.ClearBeforeFill = True
        '
        'ItemTransList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(444, 260)
        Me.Controls.Add(Me.DataGridView1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "ItemTransList"
        Me.Text = "Item Transaction List"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ItemTransBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SalesDataSet2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents SalesDataSet2 As SalesPromo.SalesDataSet2
    Friend WithEvents ItemTransBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents ItemTransTableAdapter As SalesPromo.SalesDataSet2TableAdapters.ItemTransTableAdapter
    Friend WithEvents DocNoDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DocDtDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DocTypeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LocationDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
