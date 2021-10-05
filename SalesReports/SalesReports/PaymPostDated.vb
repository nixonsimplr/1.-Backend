Imports System
Imports System.Resources
Imports System.Threading
Imports System.Reflection
Imports System.Globalization

Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Imports DataInterface

Public Class PaymPostDated
    Implements ISalesBase


#Region "Localization Variants"
    Private rMgr As ResourceManager
    Private cInfo As CultureInfo
    Private sLang As String
#End Region

    Private objDo As New DataInterface.IbizDO
    Dim strPay As String = "and PayMethod='cash'"
    Dim strPay1 As String = " "

    Private Sub PayCollection_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        objDo.DisconnectDB()
        objDo.DisconnectAnotherDB()
    End Sub

    Private Sub PayCollection_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objDo.ConnectDB()
        objDo.ConnectAnotherDB()
        loadCombo()
    End Sub

    Private Sub loadCombo()
        Dim dtr As SqlDataReader
        Dim aTerms As New ArrayList
        dtr = ReadRecord("Select Code, Name from SalesAgent")
        aTerms.Add(New ComboValues("ALL", "ALL"))
        Do While dtr.Read = True
            aTerms.Add(New ComboValues(dtr("Code").ToString, dtr("Name").ToString))
        Loop
        dtr.Close()

        cmbAgent.DataSource = aTerms
        cmbAgent.DisplayMember = "Desc"
        cmbAgent.ValueMember = "Code"

        'If cmbTerms.Items.Count > 0 Then
        '    cmbTerms.SelectedIndex = 0
        'End If
        If cmbAgent.Items.Count > 0 Then
            cmbAgent.SelectedIndex = 0
        End If
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Try
            If cmbAgent.Text = "ALL" Then
                strPay1 = " "
            Else
                strPay1 = " and R.AgentID=" & SafeSQL(cmbAgent.SelectedValue.ToString)
            End If
            'If cmbTerms.Text = "ALL" Then
            '    strPay = " "
            'Else
            '    strPay = " and R.PayMethod='" & cmbTerms.Text & "'"
            'End If

            objDo.ExecuteSQL("Delete PaymentReport")
            Dim strSql As String
            Dim sDate, sInvDt, sChqDt, sDateDesc As String
            Dim dSubTot, dDiscount, dGstAmt, dPaidAmt As Double
            strSql = "SELECT R.RcptNo, R.RcptDt, C.CustNo, C.CustName, " & _
                   "R.Amount, R.AgentID, R.PayMethod, " & _
                   "R.ChqNo, R.ChqDt, R.BankName " & _
                   "FROM Receipt R, Customer C WHERE (R.void=0 or R.Void is Null) " & _
                   "and R.CustId = C.CustNo " & _
                   "and R.Exported=1 and (R.ExportNav=0 or R.ExportNav is Null) and R.ChqDt Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & " and (PayMethod = 'CHEQUE' or PayMethod = 'CHECK') " & _
                   "" & strPay1 & " Order by R.ChqDt"

            'and R.ChqDt > " & SafeSQL(Format(Date.Now, "yyyyMMdd 23:59:59")) & " 
            'strSql = "SELECT Customer.CustName, Receipt.CustID, Receipt.RcptNo, Receipt.RcptDt, Receipt.AgentID, Invoice.InvDt, RcptItem.InvNo, Receipt.PayMethod, Receipt.ChqDt, Receipt.ChqNo, Receipt.Amount, Receipt.BankName, SalesAgent.Name FROM Receipt INNER JOIN RcptItem ON Receipt.RcptNo = RcptItem.RcptNo INNER JOIN Invoice ON RcptItem.InvNo = Invoice.InvNo AND RcptItem.RcptNo = Invoice.InvNo INNER JOIN Customer ON Receipt.CustID = Customer.CustNo INNER JOIN SalesAgent ON Receipt.AgentID = SalesAgent.Code and R.ChqDt Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & " and R.ChqDt > " & SafeSQL(Format(Date.Now, "yyyyMMdd 23:59:59")) & " and PayMethod = 'CHEQUE' " & _
            '         "" & strPay1 & " Order by Receipt.ChqDT"
            'strSql = "SELECT R.RcptNo, R.RcptDt, C.CustNo, C.CustName, " & _
            '         "RI.InvNo, RI.AmtPaid, S.Name, R.PayMethod, " & _
            '         "I.Subtotal, I.Discount, I.GstAmt, R.ChqNo, I.InvDt, R.ChqDt, R.BankName " & _
            '         "FROM Receipt R, Customer C, RcptItem RI, Invoice I, SalesAgent S WHERE (R.void=0 or R.Void is Null) " & _
            '         "and R.CustId = C.CustNo and RI.InvNo=I.InvNo and R.RcptNo = RI.RcptNo and R.AgentId = S.Code  " & _
            '         "and R.ChqDt Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & " and R.ChqDt > " & SafeSQL(Format(Date.Now, "yyyyMMdd 23:59:59")) & " and PayMethod = 'CHEQUE' " & _
            '         "" & strPay1 & " Order by I.InvDt, I.InvNo, R.RcptNo"

            'strSql = "SELECT R.RcptNo, R.RcptDt, C.CustNo, C.CustName, " & _
            '         "RI.InvNo, RI.AmtPaid, S.Name, R.PayMethod, " & _
            '         "I.Subtotal, I.Discount, I.GstAmt, R.ChqNo, I.InvDt, R.ChqDt, R.BankName " & _
            '         "FROM Receipt R, Customer C, RcptItem RI, Invoice I, SalesAgent S WHERE (R.void=0 or R.Void is Null) " & _
            '         "and R.CustId = C.CustNo and RI.InvNo=I.InvNo and R.RcptNo = RI.RcptNo and R.AgentId = S.Code  " & _
            '         "and R.ChqDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & " and PayMethod = 'CHEQUE' " & _
            '         "" & strPay1 & " Order by I.InvDt, I.InvNo, R.RcptNo"

            Dim rs As SqlDataReader
            rs = objDo.ReadRecord(strSql)

            While rs.Read = True
                If IsDBNull(rs("RcptDt")) = True Then
                    sDate = ""
                Else
                    sDate = Format(rs("RcptDt"), "yyyyMMdd HH:mm:ss")
                End If
                '  If IsDBNull(rs("InvDt")) = True Then
                sInvDt = ""
                'Else
                'sInvDt = Format(rs("InvDt"), "yyyyMMdd HH:mm:ss")
                'End If
                If IsDBNull(rs("ChqDt")) = True Then
                    sChqDt = ""
                Else
                    sChqDt = Format(rs("ChqDt"), "yyyyMMdd HH:mm:ss")
                End If

                If IsDBNull(rs("Amount")) = True Then
                    dPaidAmt = 0
                Else
                    dPaidAmt = CDbl(rs("Amount"))
                End If

                'If IsDBNull(rs("Subtotal")) = True Then
                dSubTot = 0
                'Else
                '    dSubTot = CDbl(rs("Subtotal"))
                'End If

                'If IsDBNull(rs("Discount")) = True Then
                dDiscount = 0
                'Else
                '    dDiscount = CDbl(rs("Discount"))
                'End If

                'If IsDBNull(rs("GstAmt")) = True Then
                dGstAmt = 0
                'Else
                '    dGstAmt = CDbl(rs("GstAmt"))
                'End If
                Dim sChqNo As String = ""
                Dim sBankName As String = ""
                If rs("PayMethod").ToString = "CASH" Then
                    sBankName = ""
                    sChqNo = ""
                    sChqDt = ""
                Else
                    sBankName = rs("BankName").ToString
                    sChqNo = rs("ChqNo").ToString
                    sChqDt = IIf(IsDBNull(rs("ChqDt")) = True, "", Format(rs("ChqDt"), "yyyyMMdd HH:mm:ss"))
                End If
                sDateDesc = "From Date:" & Format(dtpFromDate.Value, "dd/MM/yyyy").ToString & " To Date :" & Format(dtpToDate.Value, "dd/MM/yyyy").ToString
                If cmbAgent.Text = "ALL" Then
                    strSql = "Insert into PaymentReport (RcptNo, RcptDt, CustNo, CustName, InvNo, PaidAmt, AgentId, " & _
                             "PayMethod, SubTotal, Discount, GSTAmt, ChqNo, InvDt, ChqDt, BankName, DateDesc) Values ( " & _
                             "" & objDo.SafeSQL(rs("RcptNo").ToString) & ", " & objDo.SafeSQL(sDate) & ", " & _
                             "" & objDo.SafeSQL(rs("CustNo").ToString) & ", " & objDo.SafeSQL(rs("CustName").ToString) & ", " & _
                             "" & objDo.SafeSQL("") & ", " & dPaidAmt & ", " & _
                             "" & objDo.SafeSQL("ALL") & ", " & objDo.SafeSQL(rs("PayMethod").ToString) & ", " & _
                             "" & dSubTot & ", " & dDiscount & ", " & dGstAmt & ", " & objDo.SafeSQL(sChqNo) & ", " & _
                             "" & objDo.SafeSQL(sInvDt) & ", " & objDo.SafeSQL(sChqDt) & ", " & objDo.SafeSQL(sBankName) & "," & SafeSQL(sDateDesc) & " )"
                Else
                    strSql = "Insert into PaymentReport (RcptNo, RcptDt, CustNo, CustName, InvNo, PaidAmt, AgentId, " & _
                             "PayMethod, SubTotal, Discount, GSTAmt, ChqNo, InvDt, ChqDt, BankName, DateDesc) Values ( " & _
                             "" & objDo.SafeSQL(rs("RcptNo").ToString) & ", " & objDo.SafeSQL(sDate) & ", " & _
                             "" & objDo.SafeSQL(rs("CustNo").ToString) & ", " & objDo.SafeSQL(rs("CustName").ToString) & ", " & _
                             "" & objDo.SafeSQL("") & ", " & dPaidAmt & ", " & _
                             "" & objDo.SafeSQL(rs("AgentID").ToString) & ", " & objDo.SafeSQL(rs("PayMethod").ToString) & ", " & _
                             "" & dSubTot & ", " & dDiscount & ", " & dGstAmt & ", " & objDo.SafeSQL(sChqNo) & ", " & _
                             "" & objDo.SafeSQL(sInvDt) & ", " & objDo.SafeSQL(sChqDt) & ", " & objDo.SafeSQL(sBankName) & "," & SafeSQL(sDateDesc) & " )"
                End If
                objDo.ExecuteSQLAnother(strSql)
            End While
            rs.Close()
            If rbCust.Checked = True Then
                strSql = "Select * from PaymentReport order by CustName, RcptNo"
            Else
                strSql = "Select * from PaymentReport order by ChqDT, RcptNo"
            End If

            Dim RptName As String = "rptPostDatedCheque"
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
            rptDocument.SetParameterValue("dDate", Format(dtpToDate.Value, "dd MMM yyyy"))
            Dim frm As New ViewReport
            frm.crvReport.ShowRefreshButton = False
            frm.crvReport.ShowCloseButton = False
            frm.crvReport.ShowGroupTreeButton = False
            frm.crvReport.ReportSource = rptDocument
            frm.Show()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub cmbTerms_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbTerms.SelectedIndexChanged

    End Sub

    Private Sub cmbAgent_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAgent.SelectedIndexChanged

    End Sub

    Public Function GetListViewForm() As String Implements SalesInterface.MobileSales.ISalesBase.GetListViewForm
        Return ""
    End Function

    Public Sub ListViewClick() Implements SalesInterface.MobileSales.ISalesBase.ListViewClick


    End Sub

    Public Event LoadDataForm(ByVal Location As String, ByVal LoadType As String, ByVal ParentLoadType As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.LoadDataForm

    Public Event LoadForm(ByVal Location As String, ByVal LoadType As String) Implements SalesInterface.MobileSales.ISalesBase.LoadForm

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

    Private Sub rbCust_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbCust.CheckedChanged
        If rbCust.Checked = True Then
            rbChqDt.Checked = False
        Else
            rbChqDt.Checked = True
        End If
    End Sub

    Private Sub rbChqDt_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbChqDt.CheckedChanged
        If rbChqDt.Checked = True Then
            rbCust.Checked = False
        Else
            rbCust.Checked = True
        End If
    End Sub
End Class