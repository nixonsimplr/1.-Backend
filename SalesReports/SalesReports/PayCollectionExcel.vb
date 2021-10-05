Imports System.Data.SqlClient
Imports SalesInterface.MobileSales

Public Class PayCollectionExcel
    Implements ISalesBase
    Private aAgent As New ArrayList()
    Private aTerms As New ArrayList()



    Private Sub PayCollection_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        DisconnectDB()
    End Sub

    Private Sub PayCollection_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ConnectDB()
        dtpDate.Value = DateTime.Now
        btnPrint_Click(Nothing, Nothing)
        Me.Close()
    End Sub
    
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        btnPrint.Enabled = False
        Dim strSql As String = ""
        strSql = "SELECT Distinct Receipt.RcptNo, Receipt.RcptDt, Customer.CustNo, Customer.CustName, RcptItem.InvNo, RcptItem.AmtPaid as PaidAmt,SalesAgent.Name as AgentID, PayMethod.Description as PayMethod, (Receipt.BankName + ' ' + Receipt.ChqNo + ' ' + convert(varchar(10),Receipt.ChqDt,101)) as ChqNo, Receipt.CompanyName, (isnull(Invoice.TotalAmt,0) - isnull(Invoice.PaidAmt,0)) as SubTotal  FROM Receipt, Customer, RcptItem, SalesAgent, PayMethod, Invoice WHERE Invoice.InvNo = RcptItem.InvNo and Receipt.CompanyName = Customer.CompanyName and (Receipt.void=0 or Receipt.Void is Null) and Receipt.CustId = Customer.CustNo and SalesAgent.Code=Receipt.AgentID and PayMethod.code=Receipt.Paymethod and Receipt.RcptNo = RcptItem.RcptNo and Receipt.RcptDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59"))
        strSql = strSql & " UNION "
        strSql = strSql & "SELECT Distinct Receipt.RcptNo, Receipt.RcptDt, Customer.CustNo, Customer.CustName, RcptItem.InvNo, RcptItem.AmtPaid as PaidAmt,SalesAgent.Name as AgentID, PayMethod.Description as PayMethod, (Receipt.BankName + ' ' + Receipt.ChqNo + ' ' + convert(varchar(10),Receipt.ChqDt,101)) as ChqNo, Receipt.CompanyName, (isnull(CreditNote.TotalAmt,0) - isnull(CreditNote.PaidAmt,0)) as SubTotal FROM Receipt, Customer, RcptItem, SalesAgent, PayMethod, CreditNote WHERE CreditNote.CreditNoteNo = RcptItem.InvNo and Receipt.CompanyName = Customer.CompanyName and (Receipt.void=0 or Receipt.Void is Null) and Receipt.CustId = Customer.CustNo and SalesAgent.Code=Receipt.AgentID and PayMethod.code=Receipt.Paymethod and Receipt.RcptNo = RcptItem.RcptNo and Receipt.RcptDt Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59"))
        strSql = strSql & " Order by RcptDt, RcptNo, InvNo "
        '  ExecuteReport(strSql, "PayCollectionRep")
        Dim DA As New SqlDataAdapter(strSql, My.Settings.ConnectionString)
        Dim DS As New DataSet
        DA.Fill(DS)
        Dim ConInfo As New CrystalDecisions.Shared.TableLogOnInfo
        ConInfo.ConnectionInfo.IntegratedSecurity = True
        'ConInfo.ConnectionInfo.UserID = objDataBase.UserName
        'ConInfo.ConnectionInfo.Password = objDataBase.Password
        ConInfo.ConnectionInfo.ServerName = ".\SQLEXPRESS"
        ConInfo.ConnectionInfo.DatabaseName = "Sales"

        Dim strReportPath As String = Application.StartupPath & "\" & "PayCollectionRep" & ".rpt"
        If Not IO.File.Exists(strReportPath) Then
            Throw (New Exception("Unable to locate report file:" & vbCrLf & strReportPath))
        End If
        Dim rptDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        rptDocument.Load(strReportPath)
        rptDocument.SetDataSource(DS.Tables(0))

        Dim sPath As String = GetPayColRptPath()
        If IO.File.Exists(sPath + "\PayColRpt" + Format(Date.Now, "ddMMyyyy") & ".xls") Then
xlsAgain:
            Try
                IO.File.Delete(sPath + "\PayColRpt" + Format(Date.Now, "ddMMyyyy") & ".xls")
            Catch ex As Exception
                MsgBox("Please Close the file : " + "PayColRpt" + Format(Date.Now, "ddMMyyyy") & ".xls" & vbCrLf & "and then click OK", MsgBoxStyle.OkOnly, "Close the file first")
                GoTo xlsAgain
            End Try

        End If

        If IO.File.Exists(sPath + "\PayColRpt" + Format(Date.Now, "ddMMyyyy") & ".pdf") Then
pdfAgain:
            Try
                IO.File.Delete(sPath + "\PayColRpt" + Format(Date.Now, "ddMMyyyy") & ".pdf")
            Catch ex As Exception
                MsgBox("Please Close the file : " + "PayColRpt" + Format(Date.Now, "ddMMyyyy") & ".pdf" & vbCrLf & "and then click OK", MsgBoxStyle.OkOnly, "Close the file first")
                GoTo pdfAgain
            End Try
        End If

        If IO.File.Exists(sPath + "\PayColRpt" + Format(Date.Now, "ddMMyyyy") & ".doc") Then
docAgain:
            Try
                IO.File.Delete(sPath + "\PayColRpt" + Format(Date.Now, "ddMMyyyy") & ".doc")
            Catch ex As Exception
                MsgBox("Please Close the file : " + "PayColRpt" + Format(Date.Now, "ddMMyyyy") & ".doc" & vbCrLf & "and then click OK", MsgBoxStyle.OkOnly, "Close the file first")
                GoTo docAgain
            End Try
        End If

        rptDocument.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.Excel, sPath + "\PayColRpt" & Format(Date.Now, "ddMMyyyy") & ".xls")
        rptDocument.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.WordForWindows, sPath + "\PayColRpt" & Format(Date.Now, "ddMMyyyy") & ".doc")
        rptDocument.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, sPath + "\PayColRpt" & Format(Date.Now, "ddMMyyyy") & ".pdf")
        MsgBox("Export Completed")
        btnPrint.Enabled = True
    End Sub

    Private Function GetPayColRptPath()
        Dim ds As New DataSet
        Dim dataDirectory As String
        dataDirectory = Windows.Forms.Application.StartupPath
        ds.ReadXml(dataDirectory & "\Ibiz.xml")
        Dim table As DataTable
        For Each table In ds.Tables
            Dim row As DataRow
            If table.TableName = "PayColRptPath" Then
                For Each row In table.Rows
                    Return row("value").ToString()
                Next row
            End If
        Next table
        Return ""
    End Function

    Private Sub cmbTerms_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub cmbAgent_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

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
End Class