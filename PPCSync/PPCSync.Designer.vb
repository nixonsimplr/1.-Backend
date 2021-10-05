<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PPCSync
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
        Me.btnUploadInvn = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.btnUpdate = New System.Windows.Forms.Button
        Me.btnSync = New System.Windows.Forms.Button
        Me.btnUpload = New System.Windows.Forms.Button
        Me.btnDownload = New System.Windows.Forms.Button
        Me.dgvSync = New System.Windows.Forms.DataGridView
        Me.MDTNoDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DescriptionDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.AgentIdDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LocationDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.AgentName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.MDTBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.SalesDataSet = New Global.PPCSync.SalesDataSet
        Me.MDTTableAdapter = New Global.PPCSync.SalesDataSetTableAdapters.MDTTableAdapter
        Me.Panel1.SuspendLayout()
        CType(Me.dgvSync, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MDTBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SalesDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnUploadInvn)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.btnUpdate)
        Me.Panel1.Controls.Add(Me.btnSync)
        Me.Panel1.Controls.Add(Me.btnUpload)
        Me.Panel1.Controls.Add(Me.btnDownload)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 255)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(597, 29)
        Me.Panel1.TabIndex = 4
        '
        'btnUploadInvn
        '
        Me.btnUploadInvn.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUploadInvn.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.btnUploadInvn.Location = New System.Drawing.Point(378, 3)
        Me.btnUploadInvn.Name = "btnUploadInvn"
        Me.btnUploadInvn.Size = New System.Drawing.Size(119, 21)
        Me.btnUploadInvn.TabIndex = 8
        Me.btnUploadInvn.Text = "Upload Inventory"
        Me.btnUploadInvn.UseVisualStyleBackColor = True
        Me.btnUploadInvn.Visible = False
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(135, 3)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 7
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(2, 3)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(127, 23)
        Me.btnUpdate.TabIndex = 6
        Me.btnUpdate.Text = "New / Change User"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'btnSync
        '
        Me.btnSync.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSync.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.btnSync.Location = New System.Drawing.Point(503, 3)
        Me.btnSync.Name = "btnSync"
        Me.btnSync.Size = New System.Drawing.Size(91, 21)
        Me.btnSync.TabIndex = 2
        Me.btnSync.Text = "Synchronization"
        Me.btnSync.UseVisualStyleBackColor = True
        '
        'btnUpload
        '
        Me.btnUpload.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpload.Location = New System.Drawing.Point(216, 3)
        Me.btnUpload.Name = "btnUpload"
        Me.btnUpload.Size = New System.Drawing.Size(75, 21)
        Me.btnUpload.TabIndex = 1
        Me.btnUpload.Text = "Upload"
        Me.btnUpload.UseVisualStyleBackColor = True
        Me.btnUpload.Visible = False
        '
        'btnDownload
        '
        Me.btnDownload.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDownload.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.btnDownload.Location = New System.Drawing.Point(297, 3)
        Me.btnDownload.Name = "btnDownload"
        Me.btnDownload.Size = New System.Drawing.Size(75, 21)
        Me.btnDownload.TabIndex = 0
        Me.btnDownload.Text = "Download"
        Me.btnDownload.UseVisualStyleBackColor = True
        Me.btnDownload.Visible = False
        '
        'dgvSync
        '
        Me.dgvSync.AllowUserToAddRows = False
        Me.dgvSync.AllowUserToDeleteRows = False
        Me.dgvSync.AutoGenerateColumns = False
        Me.dgvSync.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSync.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.MDTNoDataGridViewTextBoxColumn, Me.DescriptionDataGridViewTextBoxColumn, Me.AgentIdDataGridViewTextBoxColumn, Me.LocationDataGridViewTextBoxColumn, Me.AgentName})
        Me.dgvSync.DataSource = Me.MDTBindingSource
        Me.dgvSync.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvSync.Location = New System.Drawing.Point(0, 0)
        Me.dgvSync.Name = "dgvSync"
        Me.dgvSync.ReadOnly = True
        Me.dgvSync.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvSync.Size = New System.Drawing.Size(597, 255)
        Me.dgvSync.TabIndex = 5
        '
        'MDTNoDataGridViewTextBoxColumn
        '
        Me.MDTNoDataGridViewTextBoxColumn.DataPropertyName = "MDTNo"
        Me.MDTNoDataGridViewTextBoxColumn.HeaderText = "MDTNo"
        Me.MDTNoDataGridViewTextBoxColumn.Name = "MDTNoDataGridViewTextBoxColumn"
        Me.MDTNoDataGridViewTextBoxColumn.ReadOnly = True
        '
        'DescriptionDataGridViewTextBoxColumn
        '
        Me.DescriptionDataGridViewTextBoxColumn.DataPropertyName = "Description"
        Me.DescriptionDataGridViewTextBoxColumn.HeaderText = "Description"
        Me.DescriptionDataGridViewTextBoxColumn.Name = "DescriptionDataGridViewTextBoxColumn"
        Me.DescriptionDataGridViewTextBoxColumn.ReadOnly = True
        '
        'AgentIdDataGridViewTextBoxColumn
        '
        Me.AgentIdDataGridViewTextBoxColumn.DataPropertyName = "AgentId"
        Me.AgentIdDataGridViewTextBoxColumn.HeaderText = "AgentId"
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
        'AgentName
        '
        Me.AgentName.DataPropertyName = "Name"
        Me.AgentName.HeaderText = "Agent Name"
        Me.AgentName.Name = "AgentName"
        Me.AgentName.ReadOnly = True
        '
        'MDTBindingSource
        '
        Me.MDTBindingSource.DataMember = "MDT"
        Me.MDTBindingSource.DataSource = Me.SalesDataSet
        '
        'SalesDataSet
        '
        Me.SalesDataSet.DataSetName = "SalesDataSet"
        Me.SalesDataSet.EnforceConstraints = False
        Me.SalesDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'MDTTableAdapter
        '
        Me.MDTTableAdapter.ClearBeforeFill = True
        '
        'PPCSync
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(597, 284)
        Me.Controls.Add(Me.dgvSync)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "PPCSync"
        Me.Text = "PDA Synchronization"
        Me.Panel1.ResumeLayout(False)
        CType(Me.dgvSync, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MDTBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SalesDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnUpload As System.Windows.Forms.Button
    Friend WithEvents btnDownload As System.Windows.Forms.Button
    Friend WithEvents dgvSync As System.Windows.Forms.DataGridView
    Friend WithEvents SalesDataSet As SalesDataSet
    Friend WithEvents MDTBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents MDTTableAdapter As SalesDataSetTableAdapters.MDTTableAdapter
    Friend WithEvents btnSync As System.Windows.Forms.Button
    Friend WithEvents MDTNoDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DescriptionDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents AgentIdDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LocationDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents AgentName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents btnUploadInvn As System.Windows.Forms.Button

End Class
