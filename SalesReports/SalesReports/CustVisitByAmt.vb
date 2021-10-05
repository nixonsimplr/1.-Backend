Imports System.Data
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales

Public Class CustVisitByAmt
    Implements ISalesBase

    Private aAgent As New ArrayList()
    Dim strPay As String = " "
    Dim strTerms As String

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

    Private Sub CustVisitByAmt_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        DisconnectDB()
    End Sub

    Private Sub CustVisitByAmt_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ConnectDB()

        loadCombo()

    End Sub

    Public Sub loadCombo()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Code, Name from SalesAgent order by Name")
        cmbAgent.DataSource = Nothing
        aAgent.Clear()
        aAgent.Add(New ComboValues("ALL", "ALL"))
        While dtr.Read()
            aAgent.Add(New ComboValues(dtr("Code").ToString, dtr("Name").ToString))

            '    iSelIndex = iIndex
            'End If
            'iIndex = iIndex + 1
        End While
        dtr.Close()
        cmbAgent.DataSource = aAgent
        cmbAgent.DisplayMember = "Desc"
        cmbAgent.ValueMember = "Code"
        cmbAgent.SelectedIndex = 0
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim dtr As SqlDataReader
        Dim strAgent As String = ""
        Dim strAgentRet As String = ""
        Dim sWeek As String = "week1"
        Dim iDate As Integer
        Dim strSQL As String = ""

        If dtpDate.Value.DayOfWeek <> DayOfWeek.Monday Then
            MsgBox("Please select the date on Monday to generate this report")
            Return
        End If

        If cmbAgent.Text = "ALL" Then
            strAgent = " "
            strAgentRet = " "
        Else
            strAgent = " AND AgentId = " & SafeSQL(cmbAgent.SelectedValue.ToString.Trim)
            strAgentRet = " AND GoodsReturn.SalesPersonCode = " & SafeSQL(cmbAgent.SelectedValue.ToString.Trim)
        End If
        iDate = dtpDate.Value.Day
        If iDate <= 7 Then
            sWeek = "week1"
        ElseIf iDate > 7 And iDate <= 14 Then
            sWeek = "week2"
        ElseIf iDate > 14 And iDate <= 21 Then
            sWeek = "week3"
        ElseIf iDate > 21 And iDate <= 28 Then
            sWeek = "week4"
        End If

        ExecuteSQL("Delete from CustVisitByAmtRep")
        Dim bExists As Boolean
        Dim rsCheck As SqlDataReader
        'Insert Planned Data to temporary table
        strSQL = " select S.Code as AgentID,S.Name as AgentName,C.CustNo,C.CustName,RouteDay from RouteMaster RM, RouteDet RD, MDT M, SalesAgent S, Customer C "
        strSQL = strSQL & " where RM.RouteNo = RD.RouteNo and RouteWeek = " & SafeSQL(sWeek) & strAgent
        'strSQL = strSQL & " and FromDate<=" & SafeSQL(Format(dtpDate.Value, "yyyyMMdd")) & " and ToDate>=" & SafeSQL(Format(dtpDate.Value, "yyyyMMdd")) & " "
        strSQL = strSQL & " and M.VehicleId = RM.VehicleNo and C.CustNo = RD.CustNo and S.Code = M.AgentId and C.Active = 1"

        dtr = ReadRecord(strSQL)
        While dtr.Read = True
            bExists = False
            rsCheck = ReadRecordAnother("Select * from CustVisitByAmtRep where AgentId = " & SafeSQL(dtr("AgentID").ToString) & " and CustNo = " & SafeSQL(dtr("CustNo").ToString))
            If rsCheck.Read = True Then
                bExists = True
            End If
            rsCheck.Close()
            If bExists = True Then
                'Update the data on the temporary table
                ExecuteAnotherSQL("Update CustVisitByAmtRep set IsPlanned" & dtr("RouteDay").ToString & " = 1 where AgentId = " & SafeSQL(dtr("AgentID").ToString) & " and CustNo = " & SafeSQL(dtr("CustNo").ToString))
            Else
                'Insert the data on the temporary table
                ExecuteAnotherSQL("Insert into CustVisitByAmtRep (FromDate, ToDate, AgentID, AgentName, CustNo, CustName, IsPlanned" & dtr("RouteDay").ToString & " ) VALUES (" & SafeSQL(Format(dtpDate.Value, "yyyyMMdd")) & "," & SafeSQL(Format(dtpDate.Value.AddDays(6), "yyyyMMdd")) & "," & SafeSQL(dtr("AgentID").ToString) & "," & SafeSQL(dtr("AgentName").ToString) & "," & SafeSQL(dtr("CustNo").ToString) & "," & SafeSQL(dtr("CustName").ToString) & ",1)")
            End If
        End While
        dtr.Close()

        'Check from the service table

        strSQL = " select convert(varchar,ServiceDt,101) as TransDate, CustomerId as CustNo, CustName, AgentID, S.Name as AgentName, ReasonCode from Service, Customer C, SalesAgent S"
        strSQL = strSQL & " where ServiceDt between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd")) & " and " & SafeSQL(Format(dtpDate.Value.AddDays(6), "yyyyMMdd")) & ""
        strSQL = strSQL & " and C.Active = 1 and C.CustNo = Service.CustomerId and S.Code = Service.AgentID" & strAgent
        dtr = ReadRecord(strSQL)
        While dtr.Read = True
            bExists = False
            rsCheck = ReadRecordAnother("Select * from CustVisitByAmtRep where AgentId = " & SafeSQL(dtr("AgentID").ToString) & " and CustNo = " & SafeSQL(dtr("CustNo").ToString))
            If rsCheck.Read = True Then
                bExists = True
            End If
            rsCheck.Close()
            Dim iDay As Integer = 1
            iDay = DateDiff(DateInterval.Day, dtpDate.Value, CDate(dtr("TransDate").ToString))
            If iDay < 0 Then
                iDay = iDay * -1
            End If
            iDay = iDay + 1
            Dim sReasonCode As String = ""
            If IsDBNull(dtr("ReasonCode")) = False Then
                sReasonCode = dtr("ReasonCode")
            End If
            If bExists = True Then
                'Update the data on the temporary table
                ExecuteAnotherSQL("Update CustVisitByAmtRep set Description" & iDay.ToString & " = " & SafeSQL(sReasonCode) & " where AgentId = " & SafeSQL(dtr("AgentID").ToString) & " and CustNo = " & SafeSQL(dtr("CustNo").ToString))
            Else
                'Insert the data on the temporary table
                ExecuteAnotherSQL("Insert into CustVisitByAmtRep (FromDate, ToDate, AgentID, AgentName, CustNo, CustName, Description" & iDay.ToString & " ) VALUES (" & SafeSQL(Format(dtpDate.Value, "yyyyMMdd")) & "," & SafeSQL(Format(dtpDate.Value.AddDays(6), "yyyyMMdd")) & "," & SafeSQL(dtr("AgentID").ToString) & "," & SafeSQL(dtr("AgentName").ToString) & "," & SafeSQL(dtr("CustNo").ToString) & "," & SafeSQL(dtr("CustName").ToString) & "," & SafeSQL(sReasonCode) & ")")
            End If
        End While
        dtr.Close()

        'Check from the invoice table
        strSQL = " select convert(varchar,invdt,101) as TransDate,CustId as CustNo, C.CustName,AgentId,S.Name as AgentName,sum(TotalAmt) as TotalAmt from Invoice I, Customer C, SalesAgent S"
        strSQL = strSQL & " where invdt between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd")) & " and " & SafeSQL(Format(dtpDate.Value.AddDays(6), "yyyyMMdd")) & ""
        strSQL = strSQL & " and C.CustNo = I.CustId and S.Code = I.AgentId " & strAgent
        strSQL = strSQL & " group by convert(varchar,invdt,101),CustId,AgentId, C.CustName, S.Name "
        dtr = ReadRecord(strSQL)
        While dtr.Read = True
            bExists = False
            rsCheck = ReadRecordAnother("Select * from CustVisitByAmtRep where AgentId = " & SafeSQL(dtr("AgentID").ToString) & " and CustNo = " & SafeSQL(dtr("CustNo").ToString))
            If rsCheck.Read = True Then
                bExists = True
            End If
            rsCheck.Close()
            Dim iDay As Integer = 1
            iDay = DateDiff(DateInterval.Day, dtpDate.Value, CDate(dtr("TransDate").ToString))
            If iDay < 0 Then
                iDay = iDay * -1
            End If
            iDay = iDay + 1
            If bExists = True Then
                'Update the data on the temporary table
                ExecuteAnotherSQL("Update CustVisitByAmtRep set Description" & iDay.ToString & " = " & SafeSQL("$" & dtr("TotalAmt").ToString) & " where AgentId = " & SafeSQL(dtr("AgentID").ToString) & " and CustNo = " & SafeSQL(dtr("CustNo").ToString))
            Else
                'Insert the data on the temporary table
                ExecuteAnotherSQL("Insert into CustVisitByAmtRep (FromDate, ToDate, AgentID, AgentName, CustNo, CustName, Description" & iDay.ToString & " ) VALUES (" & SafeSQL(Format(dtpDate.Value, "yyyyMMdd")) & "," & SafeSQL(Format(dtpDate.Value.AddDays(6), "yyyyMMdd")) & "," & SafeSQL(dtr("AgentID").ToString) & "," & SafeSQL(dtr("AgentName").ToString) & "," & SafeSQL(dtr("CustNo").ToString) & "," & SafeSQL(dtr("CustName").ToString) & "," & SafeSQL("$" & dtr("TotalAmt").ToString) & ")")
            End If
        End While
        dtr.Close()

        ExecuteSQL("Update CustVisitByAmtRep set Description1 = 'XX' where Description1 is null and IsPlanned1 = 1")
        ExecuteSQL("Update CustVisitByAmtRep set Description2 = 'XX' where Description2 is null and IsPlanned2 = 1")
        ExecuteSQL("Update CustVisitByAmtRep set Description3 = 'XX' where Description3 is null and IsPlanned3 = 1")
        ExecuteSQL("Update CustVisitByAmtRep set Description4 = 'XX' where Description4 is null and IsPlanned4 = 1")
        ExecuteSQL("Update CustVisitByAmtRep set Description5 = 'XX' where Description5 is null and IsPlanned5 = 1")
        ExecuteSQL("Update CustVisitByAmtRep set Description6 = 'XX' where Description6 is null and IsPlanned6 = 1")
        ExecuteSQL("Update CustVisitByAmtRep set Description7 = 'XX' where Description7 is null and IsPlanned7 = 1")

        ExecuteReport("Select * from CustVisitByAmtRep", "CustVisitByAmtRep")
    End Sub
End Class