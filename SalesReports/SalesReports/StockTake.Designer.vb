<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class StockTake
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
        Dim Label1 As System.Windows.Forms.Label
        Dim Label3 As System.Windows.Forms.Label
        Dim Label4 As System.Windows.Forms.Label
        Me.dtpFromDate = New System.Windows.Forms.DateTimePicker
        Me.btnPrint = New System.Windows.Forms.Button
        Me.dtpToDate = New System.Windows.Forms.DateTimePicker
        Me.cmbLocation = New System.Windows.Forms.ComboBox
        Label1 = New System.Windows.Forms.Label
        Label3 = New System.Windows.Forms.Label
        Label4 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Label1
        '
        Label1.AutoSize = True
        Label1.Location = New System.Drawing.Point(12, 19)
        Label1.Name = "Label1"
        Label1.Size = New System.Drawing.Size(164, 13)
        Label1.TabIndex = 113
        Label1.Text = "From Date. . . . . . . . . . . . . . . . . . "
        '
        'Label3
        '
        Label3.AutoSize = True
        Label3.Location = New System.Drawing.Point(12, 45)
        Label3.Name = "Label3"
        Label3.Size = New System.Drawing.Size(154, 13)
        Label3.TabIndex = 117
        Label3.Text = "To Date. . . . . . . . . . . . . . . . . . "
        '
        'dtpFromDate
        '
        Me.dtpFromDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpFromDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFromDate.Location = New System.Drawing.Point(124, 12)
        Me.dtpFromDate.Name = "dtpFromDate"
        Me.dtpFromDate.Size = New System.Drawing.Size(153, 20)
        Me.dtpFromDate.TabIndex = 112
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(124, 105)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(75, 23)
        Me.btnPrint.TabIndex = 111
        Me.btnPrint.Text = "Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'dtpToDate
        '
        Me.dtpToDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpToDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpToDate.Location = New System.Drawing.Point(124, 38)
        Me.dtpToDate.Name = "dtpToDate"
        Me.dtpToDate.Size = New System.Drawing.Size(153, 20)
        Me.dtpToDate.TabIndex = 116
        '
        'Label4
        '
        Label4.AutoSize = True
        Label4.Location = New System.Drawing.Point(12, 71)
        Label4.Name = "Label4"
        Label4.Size = New System.Drawing.Size(135, 13)
        Label4.TabIndex = 119
        Label4.Text = "Location . . . . . . . . . . . . . . "
        '
        'cmbLocation
        '
        Me.cmbLocation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbLocation.FormattingEnabled = True
        Me.cmbLocation.Location = New System.Drawing.Point(124, 68)
        Me.cmbLocation.Name = "cmbLocation"
        Me.cmbLocation.Size = New System.Drawing.Size(153, 21)
        Me.cmbLocation.TabIndex = 120
        '
        'StockTake
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(301, 140)
        Me.Controls.Add(Me.cmbLocation)
        Me.Controls.Add(Label4)
        Me.Controls.Add(Me.dtpToDate)
        Me.Controls.Add(Label3)
        Me.Controls.Add(Me.dtpFromDate)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Label1)
        Me.Name = "StockTake"
        Me.Text = "Stock Take"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dtpFromDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents dtpToDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents cmbLocation As System.Windows.Forms.ComboBox
End Class
