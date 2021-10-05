<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class VanInventory
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
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.btnPrint = New System.Windows.Forms.Button
        Me.dgvSync = New System.Windows.Forms.DataGridView
        Me.SyncTimeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.AgentName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SyncHistoryBindingSource2 = New System.Windows.Forms.BindingSource(Me.components)
        Me.SalesDataSet1 = New SalesReports.SalesDataSet1
        Me.SyncHistoryBindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
        Me.SyncHistoryBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.SalesDataSet = New SalesReports.SalesDataSet
        Me.SyncHistoryTableAdapter = New SalesReports.SalesDataSetTableAdapters.SyncHistoryTableAdapter
        Me.SyncHistoryTableAdapter1 = New SalesReports.SalesDataSet1TableAdapters.SyncHistoryTableAdapter
        Me.Panel1.SuspendLayout()
        CType(Me.dgvSync, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SyncHistoryBindingSource2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SalesDataSet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SyncHistoryBindingSource1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SyncHistoryBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SalesDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnPrint)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 373)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(545, 29)
        Me.Panel1.TabIndex = 6
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPrint.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.btnPrint.Location = New System.Drawing.Point(451, 3)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(91, 21)
        Me.btnPrint.TabIndex = 2
        Me.btnPrint.Text = "View"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'dgvSync
        '
        Me.dgvSync.AllowUserToAddRows = False
        Me.dgvSync.AllowUserToDeleteRows = False
        Me.dgvSync.AutoGenerateColumns = False
        Me.dgvSync.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSync.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.SyncTimeDataGridViewTextBoxColumn, Me.DataGridViewTextBoxColumn1, Me.DataGridViewTextBoxColumn2, Me.AgentName})
        Me.dgvSync.DataSource = Me.SyncHistoryBindingSource2
        Me.dgvSync.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvSync.Location = New System.Drawing.Point(0, 0)
        Me.dgvSync.Name = "dgvSync"
        Me.dgvSync.ReadOnly = True
        Me.dgvSync.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvSync.Size = New System.Drawing.Size(545, 373)
        Me.dgvSync.TabIndex = 3
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
        'AgentName
        '
        Me.AgentName.DataPropertyName = "Name"
        Me.AgentName.HeaderText = "Agent Name"
        Me.AgentName.Name = "AgentName"
        Me.AgentName.ReadOnly = True
        '
        'SyncHistoryBindingSource2
        '
        Me.SyncHistoryBindingSource2.DataMember = "SyncHistory"
        Me.SyncHistoryBindingSource2.DataSource = Me.SalesDataSet1
        '
        'SalesDataSet1
        '
        Me.SalesDataSet1.DataSetName = "SalesDataSet1"
        Me.SalesDataSet1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'SyncHistoryBindingSource1
        '
        Me.SyncHistoryBindingSource1.DataMember = "SyncHistory"
        '
        'SyncHistoryBindingSource
        '
        Me.SyncHistoryBindingSource.DataMember = "SyncHistory"
        Me.SyncHistoryBindingSource.DataSource = Me.SalesDataSet
        '
        'SalesDataSet
        '
        Me.SalesDataSet.DataSetName = "SalesDataSet"
        Me.SalesDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'SyncHistoryTableAdapter
        '
        Me.SyncHistoryTableAdapter.ClearBeforeFill = True
        '
        'SyncHistoryTableAdapter1
        '
        Me.SyncHistoryTableAdapter1.ClearBeforeFill = True
        '
        'VanDailySales
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(545, 402)
        Me.Controls.Add(Me.dgvSync)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "VanDailySales"
        Me.Text = "Van Inventory"
        Me.Panel1.ResumeLayout(False)
        CType(Me.dgvSync, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SyncHistoryBindingSource2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SalesDataSet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SyncHistoryBindingSource1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SyncHistoryBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SalesDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents SalesDataSet As SalesReports.SalesDataSet
    Friend WithEvents SyncHistoryBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents SyncHistoryTableAdapter As SalesReports.SalesDataSetTableAdapters.SyncHistoryTableAdapter
    'Friend WithEvents SalesDataSet1 As SalesReports.SalesDataSet1
    Friend WithEvents SyncHistoryBindingSource1 As System.Windows.Forms.BindingSource
    ' Friend WithEvents SyncHistoryTableAdapter1 As SalesReports.SalesDataSet1TableAdapters.SyncHistoryTableAdapter
    Friend WithEvents SyncDateDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents AgentIDDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LocationDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgvSync As System.Windows.Forms.DataGridView
    Friend WithEvents SalesDataSet1 As SalesReports.SalesDataSet1
    Friend WithEvents SyncHistoryBindingSource2 As System.Windows.Forms.BindingSource
    Friend WithEvents SyncHistoryTableAdapter1 As SalesReports.SalesDataSet1TableAdapters.SyncHistoryTableAdapter
    Friend WithEvents SyncTimeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents AgentName As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
