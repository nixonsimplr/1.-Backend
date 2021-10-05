Imports System.Data
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Public Class Frequency
    Implements ISalesBase

    Dim sAgent As String = ""

    Private Sub btnCustNo_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\DataList.dll", "DataList.Customer", "SalesReports.Frequency", "txtCustNo", 0, 0)
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
    Public Sub loadCombo()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select AgentID from MDT")
        cmbAgent.Items.Add("ALL")
        Do While dtr.Read = True
            cmbAgent.Items.Add(dtr(0))
        Loop
        dtr.Close()
        If cmbAgent.Items.Count > 0 Then
            cmbAgent.SelectedIndex = 0
        End If
    End Sub
    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        ExecuteAnotherSQL("Delete from FrequencyRep")
        Dim sAgentName As String = ""
        If cmbAgent.Text = "ALL" Then
            sAgent = " "
        Else
            sAgent = "and AgentID=" & "'" & cmbAgent.Text & "'"
            sAgentName = cmbAgent.Text
        End If

        Dim rs1 As SqlDataReader
        'Dim strSql As String = "Select Count(*) as VisitNo, CustID, OrdDt, Customer.CustName from OrderHdr, Customer where OrderHdr.CustID=Customer.CustNo  and OrdDt Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & sAgent & " group by CustID, OrdDT, Customer.CustName Union Select Count(*) as VisitNo, CustID, RcptDt, CustName from Receipt, Customer where Receipt.CustID=Customer.CustNo  and RcptDt Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & sAgent & " group by CustID, RcptDT, Customer.CustName"
        Dim strSql As String = ""
        strSql = strSql & "select count(*) as VisitNo,CustId,Sales.dbo.fn_VisitedDate_FrequencyRep(" & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & "," & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & ",CustNo," & SafeSQL(sAgentName) & ") as OrdDt,Customer.CustName from "
        strSql = strSql & " (select CustId,OrdDt,AgentId from OrderHdr union select CustId,RcptDt,AgentId from Receipt) A "
        strSql = strSql & " ,Customer where A.CustId = Customer.CustNo " & sAgent & " and OrdDt between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & " group by custid,CustName,CustNo "

        rs1 = ReadRecord(strSql)
        While rs1.Read = True
            'If IsRecordExists(rs1("CustID"), rs1("OrdDt")) = False Then
            ExecuteAnotherSQL("Insert into FrequencyRep (CustNo, CustName, DateDesc, Amount, FromDate, ToDate, AgentID) Values" & " (" & SafeSQL(rs1("CustId")) & "," & SafeSQL(rs1("CustName")) & "," & SafeSQL(rs1("OrdDt")) & "," & rs1("VisitNo") & "," & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd")) & "," & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd")) & "," & SafeSQL(cmbAgent.Text) & ")")
            'End If
        End While
        rs1.Close()
        'Select count(*), CustID, OrdDt from OrderHdr where CustID not in (Select CustID from Receipt) and OrdDT='07/19/2007' group by CustId, OrdDt
        'strSql = "Select CustNo, OrdDate, CustName from Receipt, Customer where Receipt.CustID=Customer.CustNo  and RcptDt Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & sAgent & "  group by CustID, RcptDT, Customer.CustName"
        'rs1 = ReadRecord(strSql)
        'While rs1.Read = True
        '    ExecuteAnotherSQL("Insert into FrequencyRep (CustNo, CustName, OrdDate, Amount, FromDate, ToDate) Values" & " ('" & rs1("CustId") & "','" & rs1("CustName") & "'," & SafeSQL(Format(rs1("RcptDt"), "yyyyMMdd")) & "," & rs1("VisitNo") & "," & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd")) & "," & SafeSQL(Format(dtpToDate, "yyyyMMdd")) & ")")
        'End While
        'rs1.Close()
        strSql = "SELECT * from FrequencyRep"
        ExecuteReport(strSql, "FrequencyRep")
    End Sub

    Private Function IsRecordExists(ByVal sCustID As String, ByVal sDate As Date) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean
        dtr = ReadRecordAnother("Select * from FrequencyRep where CustNo = " & SafeSQL(sCustID) & " and OrdDate=" & SafeSQL(Format(sDate, "yyyyMMdd")))
        bAns = dtr.Read
        dtr.Close()
        Return bAns
    End Function

    Private Sub Frequency_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub
    Private Sub Frequency_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '   txtCustNo.Text = "ALL"
        ConnectDB()
        loadCombo()
    End Sub

    Public Sub ReturnSearch(ByVal SQL As String) Implements SalesInterface.MobileSales.ISalesBase.ReturnSearch

    End Sub

    Public Event SearchField(ByVal FormName As String, ByVal FieldName As String, ByVal FieldType As String, ByVal CurValue As String) Implements SalesInterface.MobileSales.ISalesBase.SearchField

    Public Sub SetCulture(ByVal CultureName As String) Implements SalesInterface.MobileSales.ISalesBase.SetCulture

    End Sub
End Class