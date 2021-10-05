<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AssetTrans
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
        Me.SalesDataSet = New SalesReports.SalesDataSet
        Me.SyncHistoryBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.SalesDataSet1 = New SalesReports.SalesDataSet1
        Me.SyncHistoryTableAdapter = New SalesReports.SalesDataSetTableAdapters.SyncHistoryTableAdapter
        Me.SyncHistoryBindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
        Me.SyncHistoryTableAdapter1 = New SalesReports.SalesDataSet1TableAdapters.SyncHistoryTableAdapter
        Me.dgvSync = New System.Windows.Forms.DataGridView
        Me.SyncTimeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SyncHistoryBindingSource2 = New System.Windows.Forms.BindingSource(Me.components)
        Me.btnPrint = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        CType(Me.SalesDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SyncHistoryBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SalesDataSet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SyncHistoryBindingSource1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvSync, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SyncHistoryBindingSource2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'SalesDataSet
        '
        Me.SalesDataSet.DataSetName = "SalesDataSet"
        Me.SalesDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'SyncHistoryBindingSource
        '
        Me.SyncHistoryBindingSource.DataMember = "SyncHistory"
        Me.SyncHistoryBindingSource.DataSource = Me.SalesDataSet
        '
        'SalesDataSet1
        '
        Me.SalesDataSet1.DataSetName = "SalesDataSet1"
        Me.SalesDataSet1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'SyncHistoryTableAdapter
        '
        Me.SyncHistoryTableAdapter.ClearBeforeFill = True
        '
        'SyncHistoryBindingSource1
        '
        Me.SyncHistoryBindingSource1.DataMember = "SyncHistory"
        '
        'SyncHistoryTableAdapter1
        '
        Me.SyncHistoryTableAdapter1.ClearBeforeFill = True
        '
        'dgvSync
        '
        Me.dgvSync.AllowUserToAddRows = False
        Me.dgvSync.AllowUserToDeleteRows = False
        Me.dgvSync.AutoGenerateColumns = False
        Me.dgvSync.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSync.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.SyncTimeDataGridViewTextBoxColumn, Me.DataGridViewTextBoxColumn1, Me.DataGridViewTextBoxColumn2})
        Me.dgvSync.DataSource = Me.SyncHistoryBindingSource2
        Me.dgvSync.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvSync.Location = New System.Drawing.Point(0, 0)
        Me.dgvSync.Name = "dgvSync"
        Me.dgvSync.ReadOnly = True
        Me.dgvSync.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvSync.Size = New System.Drawing.Size(478, 390)
        Me.dgvSync.TabIndex = 7
        '
        'SyncTimeDataGridViewTextBoxColumn
        '
        Me.SyncTimeDataGridViewTextBoxColumn.DataPropertyName = "SyncTime"
        Me.SyncTimeDataGridViewTextBoxColumn.HeaderText = "SyncTime"
        Me.SyncTimeDataGridViewTextBoxColumn.Name = "SyncTimeDataGridViewTextBoxColumn"
        Me.SyncTimeDataGridViewTextBoxColumn.ReadOnly = True
        Me.SyncTimeDataGridViewTextBoxColumn.Width = 200
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.DataPropertyName = "AgentID"
        Me.DataGridViewTextBoxColumn1.HeaderText = "AgentID"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.ReadOnly = True
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.DataPropertyName = "Location"
        Me.DataGridViewTextBoxColumn2.HeaderText = "Location"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.ReadOnly = True
        '
        'SyncHistoryBindingSource2
        '
        Me.SyncHistoryBindingSource2.DataMember = "SyncHistory"
        Me.SyncHistoryBindingSource2.DataSource = Me.SalesDataSet1
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPrint.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.btnPrint.Location = New System.Drawing.Point(384, 3)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(91, 21)
        Me.btnPrint.TabIndex = 2
        Me.btnPrint.Text = "View"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnPrint)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 390)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(478, 29)
        Me.Panel1.TabIndex = 8
        '
        'AssetTrans
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(478, 419)
        Me.Controls.Add(Me.dgvSync)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "AssetTrans"
        Me.Text = "Asset Transaction"
        CType(Me.SalesDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SyncHistoryBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SalesDataSet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SyncHistoryBindingSource1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvSync, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SyncHistoryBindingSource2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SalesDataSet As SalesReports.SalesDataSet
    Friend WithEvents SyncHistoryBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents SalesDataSet1 As SalesReports.SalesDataSet1
    Friend WithEvents SyncHistoryTableAdapter As SalesReports.SalesDataSetTableAdapters.SyncHistoryTableAdapter
    Friend WithEvents SyncHistoryBindingSource1 As System.Windows.Forms.BindingSource
    Friend WithEvents SyncHistoryTableAdapter1 As SalesReports.SalesDataSet1TableAdapters.SyncHistoryTableAdapter
    Friend WithEvents dgvSync As System.Windows.Forms.DataGridView
    Friend WithEvents SyncTimeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SyncHistoryBindingSource2 As System.Windows.Forms.BindingSource
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
End Class
