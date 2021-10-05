<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MDTList
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
        Me.components = New System.ComponentModel.Container
        Me.dgvMDT = New System.Windows.Forms.DataGridView
        Me.MDTSalesDataSet = New SalesPromo.MDTSalesDataSet
        Me.MDTBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.MDTTableAdapter = New SalesPromo.MDTSalesDataSetTableAdapters.MDTTableAdapter
        Me.MDTNoDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DescriptionDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.AgentIdDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LocationDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.VehicleIDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvMDT, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MDTSalesDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MDTBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvMDT
        '
        Me.dgvMDT.AllowUserToAddRows = False
        Me.dgvMDT.AllowUserToDeleteRows = False
        Me.dgvMDT.AutoGenerateColumns = False
        Me.dgvMDT.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvMDT.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.MDTNoDataGridViewTextBoxColumn, Me.DescriptionDataGridViewTextBoxColumn, Me.AgentIdDataGridViewTextBoxColumn, Me.LocationDataGridViewTextBoxColumn, Me.VehicleIDDataGridViewTextBoxColumn})
        Me.dgvMDT.DataSource = Me.MDTBindingSource
        Me.dgvMDT.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvMDT.Location = New System.Drawing.Point(0, 0)
        Me.dgvMDT.Name = "dgvMDT"
        Me.dgvMDT.ReadOnly = True
        Me.dgvMDT.Size = New System.Drawing.Size(652, 394)
        Me.dgvMDT.TabIndex = 0
        '
        'MDTSalesDataSet
        '
        Me.MDTSalesDataSet.DataSetName = "MDTSalesDataSet"
        Me.MDTSalesDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'MDTBindingSource
        '
        Me.MDTBindingSource.DataMember = "MDT"
        Me.MDTBindingSource.DataSource = Me.MDTSalesDataSet
        '
        'MDTTableAdapter
        '
        Me.MDTTableAdapter.ClearBeforeFill = True
        '
        'MDTNoDataGridViewTextBoxColumn
        '
        Me.MDTNoDataGridViewTextBoxColumn.DataPropertyName = "MDTNo"
        Me.MDTNoDataGridViewTextBoxColumn.HeaderText = "MDT No"
        Me.MDTNoDataGridViewTextBoxColumn.Name = "MDTNoDataGridViewTextBoxColumn"
        Me.MDTNoDataGridViewTextBoxColumn.ReadOnly = True
        '
        'DescriptionDataGridViewTextBoxColumn
        '
        Me.DescriptionDataGridViewTextBoxColumn.DataPropertyName = "Description"
        Me.DescriptionDataGridViewTextBoxColumn.HeaderText = "Description"
        Me.DescriptionDataGridViewTextBoxColumn.Name = "DescriptionDataGridViewTextBoxColumn"
        Me.DescriptionDataGridViewTextBoxColumn.ReadOnly = True
        Me.DescriptionDataGridViewTextBoxColumn.Width = 150
        '
        'AgentIdDataGridViewTextBoxColumn
        '
        Me.AgentIdDataGridViewTextBoxColumn.DataPropertyName = "AgentId"
        Me.AgentIdDataGridViewTextBoxColumn.HeaderText = "Agent ID"
        Me.AgentIdDataGridViewTextBoxColumn.Name = "AgentIdDataGridViewTextBoxColumn"
        Me.AgentIdDataGridViewTextBoxColumn.ReadOnly = True
        '
        'LocationDataGridViewTextBoxColumn
        '
        Me.LocationDataGridViewTextBoxColumn.DataPropertyName = "Location"
        Me.LocationDataGridViewTextBoxColumn.HeaderText = "Location"
        Me.LocationDataGridViewTextBoxColumn.Name = "LocationDataGridViewTextBoxColumn"
        Me.LocationDataGridViewTextBoxColumn.ReadOnly = True
        '
        'VehicleIDDataGridViewTextBoxColumn
        '
        Me.VehicleIDDataGridViewTextBoxColumn.DataPropertyName = "VehicleID"
        Me.VehicleIDDataGridViewTextBoxColumn.HeaderText = "Vehicle ID"
        Me.VehicleIDDataGridViewTextBoxColumn.Name = "VehicleIDDataGridViewTextBoxColumn"
        Me.VehicleIDDataGridViewTextBoxColumn.ReadOnly = True
        '
        'MDTList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(652, 394)
        Me.Controls.Add(Me.dgvMDT)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "MDTList"
        Me.Text = "MDT List"
        CType(Me.dgvMDT, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MDTSalesDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MDTBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvMDT As System.Windows.Forms.DataGridView
    Friend WithEvents MDTSalesDataSet As SalesPromo.MDTSalesDataSet
    Friend WithEvents MDTBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents MDTTableAdapter As SalesPromo.MDTSalesDataSetTableAdapters.MDTTableAdapter
    Friend WithEvents MDTNoDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DescriptionDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents AgentIdDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LocationDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents VehicleIDDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
