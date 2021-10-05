<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class GoodsCheckINOUT
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
        Me.btnPrint = New System.Windows.Forms.Button
        Me.dgvSync = New System.Windows.Forms.DataGridView
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.SyncHistoryBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.SalesDataSet2 = New SalesReports.SalesDataSet2
        Me.SyncHistoryTableAdapter = New SalesReports.SalesDataSet2TableAdapters.SyncHistoryTableAdapter
        Me.SyncTimeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.AgentIDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LocationDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Name = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvSync, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me.SyncHistoryBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SalesDataSet2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPrint.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.btnPrint.Location = New System.Drawing.Point(453, 3)
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
        Me.dgvSync.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.SyncTimeDataGridViewTextBoxColumn, Me.AgentIDDataGridViewTextBoxColumn, Me.LocationDataGridViewTextBoxColumn, Me.Name})
        Me.dgvSync.DataSource = Me.SyncHistoryBindingSource
        Me.dgvSync.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvSync.Location = New System.Drawing.Point(0, 0)
        Me.dgvSync.Name = "dgvSync"
        Me.dgvSync.ReadOnly = True
        Me.dgvSync.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvSync.Size = New System.Drawing.Size(547, 372)
        Me.dgvSync.TabIndex = 7
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnPrint)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 372)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(547, 29)
        Me.Panel1.TabIndex = 8
        '
        'SyncHistoryBindingSource
        '
        Me.SyncHistoryBindingSource.DataMember = "SyncHistory"
        Me.SyncHistoryBindingSource.DataSource = Me.SalesDataSet2
        '
        'SalesDataSet2
        '
        Me.SalesDataSet2.DataSetName = "SalesDataSet2"
        Me.SalesDataSet2.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'SyncHistoryTableAdapter
        '
        Me.SyncHistoryTableAdapter.ClearBeforeFill = True
        '
        'SyncTimeDataGridViewTextBoxColumn
        '
        Me.SyncTimeDataGridViewTextBoxColumn.DataPropertyName = "SyncTime"
        Me.SyncTimeDataGridViewTextBoxColumn.HeaderText = "SyncTime"
        Me.SyncTimeDataGridViewTextBoxColumn.Name = "SyncTimeDataGridViewTextBoxColumn"
        Me.SyncTimeDataGridViewTextBoxColumn.ReadOnly = True
        Me.SyncTimeDataGridViewTextBoxColumn.Width = 200
        '
        'AgentIDDataGridViewTextBoxColumn
        '
        Me.AgentIDDataGridViewTextBoxColumn.DataPropertyName = "AgentID"
        Me.AgentIDDataGridViewTextBoxColumn.HeaderText = "AgentID"
        Me.AgentIDDataGridViewTextBoxColumn.Name = "AgentIDDataGridViewTextBoxColumn"
        Me.AgentIDDataGridViewTextBoxColumn.ReadOnly = True
        '
        'LocationDataGridViewTextBoxColumn
        '
        Me.LocationDataGridViewTextBoxColumn.DataPropertyName = "Location"
        Me.LocationDataGridViewTextBoxColumn.HeaderText = "Location"
        Me.LocationDataGridViewTextBoxColumn.Name = "LocationDataGridViewTextBoxColumn"
        Me.LocationDataGridViewTextBoxColumn.ReadOnly = True
        '
        'Name
        '
        Me.Name.DataPropertyName = "Name"
        Me.Name.HeaderText = "Name"
        Me.Name.Name = "Name"
        Me.Name.ReadOnly = True
        '
        'GoodsCheckINOUT
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(547, 401)
        Me.Controls.Add(Me.dgvSync)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        ' Me.Name = "GoodsCheckINOUT"
        Me.Text = "Goods Check IN and Check OUT"
        CType(Me.dgvSync, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        CType(Me.SyncHistoryBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SalesDataSet2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents dgvSync As System.Windows.Forms.DataGridView
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents SalesDataSet2 As SalesReports.SalesDataSet2
    Friend WithEvents SyncHistoryBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents SyncHistoryTableAdapter As SalesReports.SalesDataSet2TableAdapters.SyncHistoryTableAdapter
    Friend WithEvents SyncTimeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents AgentIDDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LocationDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Name As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
