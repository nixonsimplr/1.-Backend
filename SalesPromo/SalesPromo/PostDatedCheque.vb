Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.odbc
Imports DataInterface
Imports SalesInterface.MobileSales
Public Class PostDatedCheque
    Implements ISalesBase
    Dim bCmbAgentLoading As Boolean = False
    Private aAgent As New ArrayList()
    Private objDO As New DataInterface.IbizDO
    Dim strPay As String = " "

    Public Function GetListViewForm() As String Implements SalesInterface.MobileSales.ISalesBase.GetListViewForm
        Return ""
    End Function

    Public Sub ListViewClick() Implements SalesInterface.MobileSales.ISalesBase.ListViewClick

    End Sub

    Public Event LoadDataForm(ByVal Location As String, ByVal LoadType As String, ByVal ParentLoadType As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.LoadDataForm

    Public Event LoadForm(ByVal Location As String, ByVal LoadType As String) Implements SalesInterface.MobileSales.ISalesBase.LoadForm

    Private Function GetReportFolder() As String
        Dim ds As New DataSet
        Dim dataDirectory As String
        dataDirectory = Windows.Forms.Application.StartupPath
        ds.ReadXml(dataDirectory & "\ibiz.xml")
        Dim table As DataTable
        For Each table In ds.Tables
            Dim row As DataRow
            If table.TableName = "ReportFolder" Then
                For Each row In table.Rows
                    Return row("Value").ToString()
                Next row
            End If
        Next table
        Return ""
    End Function
    Public Sub MoveFirstClick() Implements SalesInterface.MobileSales.ISalesBase.MoveFirstClick

    End Sub

    Public Sub MoveLastClick() Implements SalesInterface.MobileSales.ISalesBase.MoveLastClick

    End Sub

    Public Sub MoveNextClick() Implements SalesInterface.MobileSales.ISalesBase.MoveNextClick

    End Sub

    Public Sub MovePositionClick(ByVal Position As Long) Implements SalesInterface.MobileSales.ISalesBase.MovePositionClick

    End Sub

    Public Sub MovePreviousClick() Implements SalesInterface.MobileSales.ISalesBase.MovePreviousClick

    End Sub

    Public Sub Print() Implements SalesInterface.MobileSales.ISalesBase.Print

    End Sub

    Public Sub Print(ByVal PageNo As Integer) Implements SalesInterface.MobileSales.ISalesBase.Print

    End Sub

    Public Event ResultData(ByVal ChildLoadType As String, ByVal Value As String) Implements SalesInterface.MobileSales.ISalesBase.ResultData

    Public Sub ReturnedData(ByVal Value As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.ReturnedData

    End Sub

    Public Sub ReturnSearch(ByVal SQL As String) Implements SalesInterface.MobileSales.ISalesBase.ReturnSearch

    End Sub

    Public Event SearchField(ByVal FormName As String, ByVal FieldName As String, ByVal FieldType As String, ByVal CurValue As String) Implements SalesInterface.MobileSales.ISalesBase.SearchField

    Public Sub SetCulture(ByVal CultureName As String) Implements SalesInterface.MobileSales.ISalesBase.SetCulture

    End Sub

    Private Sub btnApprove_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnApprove.Click
        Dim strSql As String = ""
        If dgvPDC.Rows.Count = 0 Then Exit Sub
        If dgvPDC.Rows.Count = 1 And dgvPDC.Item(1, 1).Value = "" Then
            Exit Sub
        End If
        Dim i As Integer
        Dim iCnt As Integer = 0
        For i = 0 To dgvPDC.Rows.Count - 2
            If Not dgvPDC.Item(0, i).Value Is Nothing Then
                If Not dgvPDC.Item(1, i).Value Is Nothing Then
                    If dgvPDC.Item(1, i).Value <> "" Then
                        If dgvPDC.Item(0, i).Value = True Then
                            iCnt += 1
                        End If
                    End If
                End If
            End If
        Next
        If iCnt = 0 Then
            MsgBox("No rows selected")
            Exit Sub
        End If
        objDO.ExecuteSQL("Delete PaymentReport")
        If MsgBox("You are about to send " & iCnt & " Payment Information to Navision." & vbCrLf & "Do you want to continue?", MsgBoxStyle.YesNo, "Export") = MsgBoxResult.Yes Then
            For i = 0 To dgvPDC.Rows.Count - 2
                If Not dgvPDC.Item(0, i).Value Is Nothing Then
                    If Not dgvPDC.Item(1, i).Value Is Nothing Then
                        If dgvPDC.Item(1, i).Value <> "" Then
                            If dgvPDC.Item(0, i).Value = True Then
                                objDO.ExecuteSQL("Update Receipt Set Exported = 0 where RcptNo=" & objDO.SafeSQL(dgvPDC.Item(4, i).Value.ToString))
                                strSql = "Insert into PaymentReport (RcptNo, RcptDt, ChqDt, PaidAmt, ChqNo, BankName) Values (" & _
                                "" & objDO.SafeSQL(dgvPDC.Item(4, i).Value.ToString) & ", " & objDO.SafeSQL(Format(CDate(dgvPDC.Item(5, i).Value), "yyyyMMdd HH:mm")) & ", " & objDO.SafeSQL(Format(CDate(dgvPDC.Item(1, i).Value), "yyyyMMdd HH:mm")) & ", " & _
                                "" & dgvPDC.Item(8, i).Value.ToString.Replace(",", "") & ", " & _
                                "" & objDO.SafeSQL(dgvPDC.Item(2, i).Value.ToString) & ", " & objDO.SafeSQL(dgvPDC.Item(3, i).Value.ToString) & " )"
                                objDO.ExecuteSQL(strSql)
                            End If
                        End If
                    End If
                End If
            Next
            LoadDataGrid()
            strSql = "Select * from PaymentReport Order by Rcptno"
            Dim RptName As String = "BankInSlipRpt"
            Dim DA As New SqlDataAdapter(strSql, My.Settings.ConnectionString)
            Dim DS As New DataSet
            DA.Fill(DS)
            Dim ConInfo As New CrystalDecisions.Shared.TableLogOnInfo
            ConInfo.ConnectionInfo.IntegratedSecurity = True
            ConInfo.ConnectionInfo.ServerName = ".\SQLEXPRESS"
            ConInfo.ConnectionInfo.DatabaseName = "Sales"
            Dim strReportPath As String = Application.StartupPath & "\" & RptName & ".rpt"
            If Not IO.File.Exists(strReportPath) Then
                Throw (New Exception("Unable to locate report file:" & vbCrLf & strReportPath))
            End If
            Dim rptDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument
            rptDocument.Load(strReportPath)
            rptDocument.SetDataSource(DS.Tables(0))
            If IO.Directory.Exists(Application.StartupPath & "\Bank-In Slip") = False Then
                IO.Directory.CreateDirectory(Application.StartupPath & "\Bank-In Slip")
            End If
            rptDocument.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, Application.StartupPath & "\Bank-In Slip\" & Format(Date.Now, "ddMMyyyy HH-mm-ss") & ".pdf")
            Dim strReportFolder As String = ""
            Try
                strReportFolder = GetReportFolder()
            Catch ex As Exception
                strReportFolder = ""
            End Try
            rptDocument.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, strReportFolder & "\Bank-In Slip\" & Format(Date.Now, "ddMMyyyy HH-mm-ss") & ".pdf")


            Dim frm As New ViewReport
            frm.crvReport.ShowRefreshButton = False
            frm.crvReport.ShowCloseButton = False
            frm.crvReport.ShowGroupTreeButton = False
            frm.crvReport.ReportSource = rptDocument
            frm.Show()
            MsgBox("Approved Successfully", MsgBoxStyle.OkOnly, "Post Dated Cheque")
        End If


        'Dim rs As SqlDataReader
        'rs = objDO.ReadRecord("Select RcptNo, RcptDT, ChqNo, ChqDT, BankName, Receipt.CustID, CustName, Amount from Receipt, Customer where Customer.CustNo=Receipt.CustID and (Void =0 or Void is Null) and (Paymethod='CHEQUE' or Paymethod='CHECK') and Receipt.Exported=1 and ChqDt >= " & objDO.SafeSQL(Format(Date.Now, "yyyyMMdd")) & " Order by ChqDT")
        'dgvPDC.Rows.Clear()
        'While rs.Read
        '    Dim row As String() = New String() _
        '{False, Format(rs("chqDT"), "dd/MM/yyyy"), rs("ChqNo").ToString, rs("BankName").ToString, rs("RcptNo").ToString, Format(rs("RcptDT"), "dd/MM/yyyy"), rs("CustID").ToString, rs("CustName").ToString, Format(rs("Amount"), "0,0.00")}
        '    dgvPDC.Rows.Add(row)
        'End While
        'rs.Close()
    End Sub

    Private Sub PostDatedCheque_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        objDO.DisconnectDB()
    End Sub
    Public Sub loadCombo()
        bCmbAgentLoading = True
        Dim dtr As SqlDataReader
        '  dtr = objDO.ReadRecord("Select MDT.MDTNo as Code, MDT.Description from MDT order by Description")
        dtr = objDO.ReadRecord("Select SalesAgent.Code, SalesAgent.Name as Description from SalesAgent order by Code")
        cmbAgent.DataSource = Nothing
        aAgent.Clear()
        aAgent.Add(New ComboValues("ALL", "ALL"))
        While dtr.Read()
            aAgent.Add(New ComboValues(dtr("Code").ToString, dtr("Description").ToString))

            '    iSelIndex = iIndex
            'End If
            'iIndex = iIndex + 1
        End While
        dtr.Close()
        cmbAgent.DataSource = aAgent
        cmbAgent.DisplayMember = "Desc"
        cmbAgent.ValueMember = "Code"
        dtpFromDate.Value = Date.Now
        dtpToDate.Value = Date.Now
        bCmbAgentLoading = False
    End Sub
    Private Sub PostDatedCheque_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objDO.ConnectDB()
        loadCombo()
        chkAllCheque.Checked = True
        LoadDataGrid()
    End Sub
    Private Sub LoadDataGrid()
        Dim rs As SqlDataReader
        If chkAllCheque.Checked = True Then
            rs = objDO.ReadRecord("Select RcptNo, RcptDT, ChqNo, ChqDT, BankName, Receipt.CustID, CustName, Amount from Receipt, Customer where Customer.CustNo=Receipt.CustID and (ExportNav =0 or ExportNav is Null) and (Void =0 or Void is Null) and (Paymethod='CHEQUE' or Paymethod='CHECK') and Receipt.Exported=1 " & strPay & " Order by ChqDT")
        Else
            rs = objDO.ReadRecord("Select RcptNo, RcptDT, ChqNo, ChqDT, BankName, Receipt.CustID, CustName, Amount from Receipt, Customer where Customer.CustNo=Receipt.CustID and (ExportNav =0 or ExportNav is Null) and (Void =0 or Void is Null) and (Paymethod='CHEQUE' or Paymethod='CHECK') and Receipt.Exported=1 and ChqDt between " & objDO.SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00.000")) & " and " & objDO.SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59.000")) & strPay & " Order by ChqDT")
        End If
        dgvPDC.Rows.Clear()
        While rs.Read
            Dim row As String() = New String() _
                           {False, Format(rs("chqDT"), "dd/MMM/yyyy"), rs("ChqNo").ToString, rs("BankName").ToString, rs("RcptNo").ToString, Format(rs("RcptDT"), "dd/MMM/yyyy"), rs("CustID").ToString, rs("CustName").ToString, Format(rs("Amount"), "0,0.00")}
            dgvPDC.Rows.Add(row)
        End While
        rs.Close()
        rs.Dispose()
    End Sub
    Private Sub btnDecline_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDecline.Click
        If dgvPDC.Rows.Count = 0 Then Exit Sub
        If dgvPDC.Rows.Count = 1 And dgvPDC.Item(1, 0).Value Is Nothing Then
            Exit Sub
        End If
        Dim bRows As Boolean = False
        For idx As Integer = 0 To dgvPDC.Rows.Count - 2
            If Not dgvPDC.Item(1, idx).Value Is Nothing Then
                If dgvPDC.Item(0, idx).Value = True And dgvPDC.Item(1, idx).Value <> "" Then
                    bRows = True
                End If
            End If
        Next
        If bRows = False Then
            MsgBox("No Rows selected")
            Exit Sub
        End If

        Dim frm As New Password()
        Dim dlgResult As DialogResult = frm.ShowDialog()
        If dlgResult = Windows.Forms.DialogResult.Cancel Then
            Exit Sub
        End If

        Dim i As Integer
        Dim iCnt As Integer = 0
        For i = 0 To dgvPDC.Rows.Count - 2
            If dgvPDC.Item(0, i).Value = True Then
                iCnt += 1
            End If
        Next
        If MsgBox("You are about to Decline " & iCnt & " Post Dated Cheques." & vbCrLf & "Do you want to continue?", MsgBoxStyle.YesNo, "Export") = MsgBoxResult.Yes Then
            For i = 0 To dgvPDC.Rows.Count - 2
                If Not dgvPDC.Item(0, i).Value Is Nothing Then
                    If Not dgvPDC.Item(1, i).Value Is Nothing Then
                        If dgvPDC.Item(1, i).Value <> "" Then
                            If dgvPDC.Item(0, i).Value = True Then
                                objDO.ExecuteSQL("Update Receipt Set ExportNav=1, Void = 1 where RcptNo=" & objDO.SafeSQL(dgvPDC.Item(4, i).Value.ToString))
                                '                        MsgBox("Declined Successfully")

                            End If
                        End If
                    End If
                End If
            Next
            MsgBox("Declined Successfully", MsgBoxStyle.OkOnly, "Post Dated Cheque")
            LoadDataGrid()
        End If

        'Dim rs As SqlDataReader
        'rs = objDO.ReadRecord("Select RcptNo, RcptDT, BankName, ChqNo, ChqDT, Receipt.CustID, CustName, Amount from Receipt, Customer where Customer.CustNo=Receipt.CustID and (Void =0 or Void is Null) and (Paymethod='CHEQUE' or Paymethod='CHECK') and Receipt.Exported=1 and ChqDt >= " & objDO.SafeSQL(Format(Date.Now, "yyyyMMdd")) & " Order by ChqDT")
        'dgvPDC.Rows.Clear()
        'While rs.Read
        '    Dim row As String() = New String() _
        '        {False, Format(rs("chqDT"), "dd/MM/yyyy"), rs("ChqNo").ToString, rs("BankName").ToString, rs("RcptNo").ToString, Format(rs("RcptDT"), "dd/MM/yyyy"), rs("CustID").ToString, rs("CustName").ToString, Format(rs("Amount"), "0,0.00")}
        '    dgvPDC.Rows.Add(row)
        'End While
        'rs.Close()
    End Sub

    Private Sub cmbAgent_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbAgent.SelectedIndexChanged
        If bCmbAgentLoading = True Then Exit Sub
        If cmbAgent.Text = "ALL" Then
            strPay = " "
        Else
            strPay = "and Receipt.AgentID = " & "'" & cmbAgent.SelectedValue & "'"
        End If
        LoadDataGrid()
    End Sub

    Private Sub dtpToDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpToDate.ValueChanged
        If cmbAgent.Text = "ALL" Then
            strPay = ""
        Else
            strPay = "and Receipt.AgentID = " & "'" & cmbAgent.SelectedValue & "'"
        End If
        LoadDataGrid()
    End Sub

    Private Sub dtpFromDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpFromDate.ValueChanged
        If cmbAgent.Text = "ALL" Then
            strPay = " "
        Else
            strPay = "and Receipt.AgentID = " & "'" & cmbAgent.SelectedValue & "'"
        End If
        LoadDataGrid()
    End Sub

    Private Sub chkSelAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSelAll.CheckedChanged
        Dim i As Integer
        If chkSelAll.Checked = True Then
            For i = 0 To dgvPDC.Rows.Count - 2
                dgvPDC.Item(0, i).Value = True
            Next
        Else
            For i = 0 To dgvPDC.Rows.Count - 2
                dgvPDC.Item(0, i).Value = False
            Next
        End If
    End Sub

    Private Sub btnApprove_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnApprove.LocationChanged

    End Sub

    Private Sub chkAllCheque_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAllCheque.CheckedChanged
        If chkAllCheque.Checked = True Then
            dtpFromDate.Enabled = False
            dtpToDate.Enabled = False
            LoadDataGrid()
        Else
            dtpFromDate.Enabled = True
            dtpToDate.Enabled = True
            LoadDataGrid()
        End If
    End Sub

End Class