<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class POList
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
        Me.dgv = New System.Windows.Forms.DataGridView
        Me.PODataSet = New SalesPromo.PODataSet
        Me.POBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.POTableAdapter = New SalesPromo.PODataSetTableAdapters.POTableAdapter
        Me.PONoDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PODtDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GSTDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn5 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn6 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn7 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PrintNoDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.VoidDataGridViewCheckBoxColumn = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.VoidDateDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn8 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TermDaysDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn9 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn10 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ExportedDataGridViewCheckBoxColumn = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ExportDateDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CustRefNoDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DTGDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.AccountStringDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.MDTNoDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.RemarksDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgv, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PODataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.POBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgv
        '
        Me.dgv.AllowUserToAddRows = False
        Me.dgv.AllowUserToDeleteRows = False
        Me.dgv.AutoGenerateColumns = False
        Me.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.PONoDataGridViewTextBoxColumn, Me.PODtDataGridViewTextBoxColumn, Me.DataGridViewTextBoxColumn1, Me.DataGridViewTextBoxColumn2, Me.DataGridViewTextBoxColumn3, Me.DataGridViewTextBoxColumn4, Me.GSTDataGridViewTextBoxColumn, Me.DataGridViewTextBoxColumn5, Me.DataGridViewTextBoxColumn6, Me.DataGridViewTextBoxColumn7, Me.PrintNoDataGridViewTextBoxColumn, Me.VoidDataGridViewCheckBoxColumn, Me.VoidDateDataGridViewTextBoxColumn, Me.DataGridViewTextBoxColumn8, Me.TermDaysDataGridViewTextBoxColumn, Me.DataGridViewTextBoxColumn9, Me.DataGridViewTextBoxColumn10, Me.ExportedDataGridViewCheckBoxColumn, Me.ExportDateDataGridViewTextBoxColumn, Me.CustRefNoDataGridViewTextBoxColumn, Me.DTGDataGridViewTextBoxColumn, Me.AccountStringDataGridViewTextBoxColumn, Me.MDTNoDataGridViewTextBoxColumn, Me.RemarksDataGridViewTextBoxColumn})
        Me.dgv.DataSource = Me.POBindingSource
        Me.dgv.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgv.Location = New System.Drawing.Point(0, 0)
        Me.dgv.Name = "dgv"
        Me.dgv.ReadOnly = True
        Me.dgv.Size = New System.Drawing.Size(631, 266)
        Me.dgv.TabIndex = 0
        '
        'PODataSet
        '
        Me.PODataSet.DataSetName = "PODataSet"
        Me.PODataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'POBindingSource
        '
        Me.POBindingSource.DataMember = "PO"
        Me.POBindingSource.DataSource = Me.PODataSet
        '
        'POTableAdapter
        '
        Me.POTableAdapter.ClearBeforeFill = True
        '
        'PONoDataGridViewTextBoxColumn
        '
        Me.PONoDataGridViewTextBoxColumn.DataPropertyName = "PONo"
        Me.PONoDataGridViewTextBoxColumn.HeaderText = "PO No"
        Me.PONoDataGridViewTextBoxColumn.Name = "PONoDataGridViewTextBoxColumn"
        Me.PONoDataGridViewTextBoxColumn.ReadOnly = True
        '
        'PODtDataGridViewTextBoxColumn
        '
        Me.PODtDataGridViewTextBoxColumn.DataPropertyName = "PODt"
        Me.PODtDataGridViewTextBoxColumn.HeaderText = "PO Dt"
        Me.PODtDataGridViewTextBoxColumn.Name = "PODtDataGridViewTextBoxColumn"
        Me.PODtDataGridViewTextBoxColumn.ReadOnly = True
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.DataPropertyName = "CustId"
        Me.DataGridViewTextBoxColumn1.HeaderText = "Cust Id"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.ReadOnly = True
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.DataPropertyName = "AgentId"
        Me.DataGridViewTextBoxColumn2.HeaderText = "Agent Id"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.ReadOnly = True
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.DataPropertyName = "Discount"
        Me.DataGridViewTextBoxColumn3.HeaderText = "Discount"
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        Me.DataGridViewTextBoxColumn3.ReadOnly = True
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.DataPropertyName = "SubTotal"
        Me.DataGridViewTextBoxColumn4.HeaderText = "SubTotal"
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        Me.DataGridViewTextBoxColumn4.ReadOnly = True
        '
        'GSTDataGridViewTextBoxColumn
        '
        Me.GSTDataGridViewTextBoxColumn.DataPropertyName = "GST"
        Me.GSTDataGridViewTextBoxColumn.HeaderText = "GST"
        Me.GSTDataGridViewTextBoxColumn.Name = "GSTDataGridViewTextBoxColumn"
        Me.GSTDataGridViewTextBoxColumn.ReadOnly = True
        '
        'DataGridViewTextBoxColumn5
        '
        Me.DataGridViewTextBoxColumn5.DataPropertyName = "GstAmt"
        Me.DataGridViewTextBoxColumn5.HeaderText = "GstAmt"
        Me.DataGridViewTextBoxColumn5.Name = "DataGridViewTextBoxColumn5"
        Me.DataGridViewTextBoxColumn5.ReadOnly = True
        '
        'DataGridViewTextBoxColumn6
        '
        Me.DataGridViewTextBoxColumn6.DataPropertyName = "TotalAmt"
        Me.DataGridViewTextBoxColumn6.HeaderText = "TotalAmt"
        Me.DataGridViewTextBoxColumn6.Name = "DataGridViewTextBoxColumn6"
        Me.DataGridViewTextBoxColumn6.ReadOnly = True
        '
        'DataGridViewTextBoxColumn7
        '
        Me.DataGridViewTextBoxColumn7.DataPropertyName = "PaidAmt"
        Me.DataGridViewTextBoxColumn7.HeaderText = "PaidAmt"
        Me.DataGridViewTextBoxColumn7.Name = "DataGridViewTextBoxColumn7"
        Me.DataGridViewTextBoxColumn7.ReadOnly = True
        '
        'PrintNoDataGridViewTextBoxColumn
        '
        Me.PrintNoDataGridViewTextBoxColumn.DataPropertyName = "PrintNo"
        Me.PrintNoDataGridViewTextBoxColumn.HeaderText = "PrintNo"
        Me.PrintNoDataGridViewTextBoxColumn.Name = "PrintNoDataGridViewTextBoxColumn"
        Me.PrintNoDataGridViewTextBoxColumn.ReadOnly = True
        '
        'VoidDataGridViewCheckBoxColumn
        '
        Me.VoidDataGridViewCheckBoxColumn.DataPropertyName = "Void"
        Me.VoidDataGridViewCheckBoxColumn.HeaderText = "Void"
        Me.VoidDataGridViewCheckBoxColumn.Name = "VoidDataGridViewCheckBoxColumn"
        Me.VoidDataGridViewCheckBoxColumn.ReadOnly = True
        '
        'VoidDateDataGridViewTextBoxColumn
        '
        Me.VoidDateDataGridViewTextBoxColumn.DataPropertyName = "VoidDate"
        Me.VoidDateDataGridViewTextBoxColumn.HeaderText = "VoidDate"
        Me.VoidDateDataGridViewTextBoxColumn.Name = "VoidDateDataGridViewTextBoxColumn"
        Me.VoidDateDataGridViewTextBoxColumn.ReadOnly = True
        '
        'DataGridViewTextBoxColumn8
        '
        Me.DataGridViewTextBoxColumn8.DataPropertyName = "PayTerms"
        Me.DataGridViewTextBoxColumn8.HeaderText = "PayTerms"
        Me.DataGridViewTextBoxColumn8.Name = "DataGridViewTextBoxColumn8"
        Me.DataGridViewTextBoxColumn8.ReadOnly = True
        '
        'TermDaysDataGridViewTextBoxColumn
        '
        Me.TermDaysDataGridViewTextBoxColumn.DataPropertyName = "TermDays"
        Me.TermDaysDataGridViewTextBoxColumn.HeaderText = "TermDays"
        Me.TermDaysDataGridViewTextBoxColumn.Name = "TermDaysDataGridViewTextBoxColumn"
        Me.TermDaysDataGridViewTextBoxColumn.ReadOnly = True
        '
        'DataGridViewTextBoxColumn9
        '
        Me.DataGridViewTextBoxColumn9.DataPropertyName = "CurCode"
        Me.DataGridViewTextBoxColumn9.HeaderText = "CurCode"
        Me.DataGridViewTextBoxColumn9.Name = "DataGridViewTextBoxColumn9"
        Me.DataGridViewTextBoxColumn9.ReadOnly = True
        '
        'DataGridViewTextBoxColumn10
        '
        Me.DataGridViewTextBoxColumn10.DataPropertyName = "CurExRate"
        Me.DataGridViewTextBoxColumn10.HeaderText = "CurExRate"
        Me.DataGridViewTextBoxColumn10.Name = "DataGridViewTextBoxColumn10"
        Me.DataGridViewTextBoxColumn10.ReadOnly = True
        '
        'ExportedDataGridViewCheckBoxColumn
        '
        Me.ExportedDataGridViewCheckBoxColumn.DataPropertyName = "Exported"
        Me.ExportedDataGridViewCheckBoxColumn.HeaderText = "Exported"
        Me.ExportedDataGridViewCheckBoxColumn.Name = "ExportedDataGridViewCheckBoxColumn"
        Me.ExportedDataGridViewCheckBoxColumn.ReadOnly = True
        '
        'ExportDateDataGridViewTextBoxColumn
        '
        Me.ExportDateDataGridViewTextBoxColumn.DataPropertyName = "ExportDate"
        Me.ExportDateDataGridViewTextBoxColumn.HeaderText = "ExportDate"
        Me.ExportDateDataGridViewTextBoxColumn.Name = "ExportDateDataGridViewTextBoxColumn"
        Me.ExportDateDataGridViewTextBoxColumn.ReadOnly = True
        '
        'CustRefNoDataGridViewTextBoxColumn
        '
        Me.CustRefNoDataGridViewTextBoxColumn.DataPropertyName = "CustRefNo"
        Me.CustRefNoDataGridViewTextBoxColumn.HeaderText = "CustRefNo"
        Me.CustRefNoDataGridViewTextBoxColumn.Name = "CustRefNoDataGridViewTextBoxColumn"
        Me.CustRefNoDataGridViewTextBoxColumn.ReadOnly = True
        '
        'DTGDataGridViewTextBoxColumn
        '
        Me.DTGDataGridViewTextBoxColumn.DataPropertyName = "DTG"
        Me.DTGDataGridViewTextBoxColumn.HeaderText = "DTG"
        Me.DTGDataGridViewTextBoxColumn.Name = "DTGDataGridViewTextBoxColumn"
        Me.DTGDataGridViewTextBoxColumn.ReadOnly = True
        '
        'AccountStringDataGridViewTextBoxColumn
        '
        Me.AccountStringDataGridViewTextBoxColumn.DataPropertyName = "AccountString"
        Me.AccountStringDataGridViewTextBoxColumn.HeaderText = "AccountString"
        Me.AccountStringDataGridViewTextBoxColumn.Name = "AccountStringDataGridViewTextBoxColumn"
        Me.AccountStringDataGridViewTextBoxColumn.ReadOnly = True
        '
        'MDTNoDataGridViewTextBoxColumn
        '
        Me.MDTNoDataGridViewTextBoxColumn.DataPropertyName = "MDTNo"
        Me.MDTNoDataGridViewTextBoxColumn.HeaderText = "MDTNo"
        Me.MDTNoDataGridViewTextBoxColumn.Name = "MDTNoDataGridViewTextBoxColumn"
        Me.MDTNoDataGridViewTextBoxColumn.ReadOnly = True
        '
        'RemarksDataGridViewTextBoxColumn
        '
        Me.RemarksDataGridViewTextBoxColumn.DataPropertyName = "Remarks"
        Me.RemarksDataGridViewTextBoxColumn.HeaderText = "Remarks"
        Me.RemarksDataGridViewTextBoxColumn.Name = "RemarksDataGridViewTextBoxColumn"
        Me.RemarksDataGridViewTextBoxColumn.ReadOnly = True
        '
        'DataList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(631, 266)
        Me.Controls.Add(Me.dgv)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "DataList"
        Me.Text = "PO List"
        CType(Me.dgv, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PODataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.POBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgv As System.Windows.Forms.DataGridView
    '  Friend WithEvents InvoiceDataSet As InvoiceDataSet
    ' Friend WithEvents InvoiceTableAdapter As InvoiceDataSetTableAdapters.InvoiceTableAdapter
    Friend WithEvents InvNoDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents InvDtDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents OrdNoDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CustIdDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents AgentIDDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PayTermsDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CurCodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CurExRateDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DiscountDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SubTotalDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GSTAmtDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TotalAmtDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PaidAmtDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PODataSet As SalesPromo.PODataSet
    Friend WithEvents POBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents POTableAdapter As SalesPromo.PODataSetTableAdapters.POTableAdapter
    Friend WithEvents PONoDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PODtDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GSTDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn7 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PrintNoDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents VoidDataGridViewCheckBoxColumn As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents VoidDateDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn8 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TermDaysDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn9 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn10 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ExportedDataGridViewCheckBoxColumn As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ExportDateDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CustRefNoDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DTGDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents AccountStringDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents MDTNoDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RemarksDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
