Imports System.Data
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales

Public Class ServiceLevel
    Implements ISalesBase

    Private aAgent As New ArrayList
    Private sAgent As String = " "
    Private sAgentCrNote As String = " "
    Private dtFromDate As DateTime = Date.Now
    Private dtToDate As DateTime = Date.Now
    Private iWeekNo As Integer = 1

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

    Private Sub ServiceLevel_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub


    Private Sub ServiceLevel_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ConnectDB()
        LoadAgent()
    End Sub

    Public Sub LoadAgent()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Code, Name from SalesAgent, MDT where MDT.AgentID=SalesAgent.Code order by Name")
        cmbAgent.DataSource = Nothing
        aAgent.Clear()
        aAgent.Add(New ComboValues("ALL", "ALL"))
        While dtr.Read()
            aAgent.Add(New ComboValues(dtr("Code").ToString, dtr("Name").ToString))
        End While
        dtr.Close()
        cmbAgent.DataSource = aAgent
        cmbAgent.DisplayMember = "Desc"
        cmbAgent.ValueMember = "Code"
        cmbAgent.SelectedIndex = 0
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        If cmbAgent.Text = "ALL" Then
            sAgent = " "
            sAgentCrNote = " "
        Else
            sAgent = " and Invoice.Agentid=" & SafeSQL(cmbAgent.SelectedValue)
            sAgentCrNote = " and CreditNote.SalesPersonCode = " & SafeSQL(cmbAgent.SelectedValue)
        End If

        Dim strSql As String = ""
        strSql = "Delete from ServiceLevelRep"
        ExecuteSQL(strSql)

        'Get From Date and To Date
        Dim aDate As DateTime = Date.Now
        aDate = dtpDate.Value

        dtFromDate = aDate.AddDays(-1 * (CInt(aDate.DayOfWeek) - 1))
        dtToDate = dtFromDate.AddDays(5)
        iWeekNo = Format(DatePart(DateInterval.WeekOfYear, dtFromDate))

        If cmbAgent.Text = "ALL" Then
            For i As Integer = 0 To cmbAgent.Items.Count - 1
                Dim cmbVal As ComboValues
                cmbVal = cmbAgent.Items(i)
                If cmbVal.Code <> "ALL" Then
                    LoadData(cmbVal.Code)
                End If
            Next
        Else
            LoadData(cmbAgent.SelectedValue)
        End If

        'Update FromDate and ToDate and Week No
        strSql = "Update ServiceLevelRep set FromDate = " & SafeSQL(Format(dtFromDate, "yyyyMMdd 00:00:00")) & ", ToDate = " & SafeSQL(Format(dtToDate, "yyyyMMdd 23:59:59")) & ", WeekNo = " & iWeekNo.ToString
        ExecuteSQL(strSql)

        strSql = "Select * from ServiceLevelRep"
        ExecuteReport(strSql, "ServiceLevelRep")
    End Sub

    Private Sub LoadData(ByVal strAgentId As String)
        Dim strSql As String = ""
        Dim dtr As SqlDataReader
        Dim iWeek As Integer = dtFromDate.Day
        Dim sWeek As String = "Week1"
        If iWeek <= 7 Then
            sWeek = "Week1"
        ElseIf iWeek > 7 And iWeek <= 14 Then
            sWeek = "Week2"
        ElseIf iWeek > 14 And iWeek <= 21 Then
            sWeek = "Week3"
        ElseIf iWeek > 21 And iWeek <= 28 Then
            sWeek = "Week4"
        Else
            sWeek = "Week1"
        End If

        'Insert the data to the temporarty table
        strSql = "insert into ServiceLevelRep ([CustomerClass],[LoadedQty],[FullQty],[SkipQty],[NSRQty],[MiscQty],[MissedQty],[DMissedQty],[TMissedQty],[AlertQty],[CINoQty],[ServicedQty],[AgentId],[FromDate],[ToDate],[WeekNo])"
        strSql = strSql & " select 'A', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0," & SafeSQL(strAgentId) & ",getdate(),getdate(),0 "
        ExecuteSQL(strSql)
        strSql = "insert into ServiceLevelRep ([CustomerClass],[LoadedQty],[FullQty],[SkipQty],[NSRQty],[MiscQty],[MissedQty],[DMissedQty],[TMissedQty],[AlertQty],[CINoQty],[ServicedQty],[AgentId],[FromDate],[ToDate],[WeekNo])"
        strSql = strSql & " select 'B', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0," & SafeSQL(strAgentId) & ",getdate(),getdate(),0 "
        ExecuteSQL(strSql)
        strSql = "insert into ServiceLevelRep ([CustomerClass],[LoadedQty],[FullQty],[SkipQty],[NSRQty],[MiscQty],[MissedQty],[DMissedQty],[TMissedQty],[AlertQty],[CINoQty],[ServicedQty],[AgentId],[FromDate],[ToDate],[WeekNo])"
        strSql = strSql & " select 'C', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0," & SafeSQL(strAgentId) & ",getdate(),getdate(),0 "
        ExecuteSQL(strSql)
        strSql = "insert into ServiceLevelRep ([CustomerClass],[LoadedQty],[FullQty],[SkipQty],[NSRQty],[MiscQty],[MissedQty],[DMissedQty],[TMissedQty],[AlertQty],[CINoQty],[ServicedQty],[AgentId],[FromDate],[ToDate],[WeekNo])"
        strSql = strSql & " select 'D', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0," & SafeSQL(strAgentId) & ",getdate(),getdate(),0 "
        ExecuteSQL(strSql)
        strSql = "insert into ServiceLevelRep ([CustomerClass],[LoadedQty],[FullQty],[SkipQty],[NSRQty],[MiscQty],[MissedQty],[DMissedQty],[TMissedQty],[AlertQty],[CINoQty],[ServicedQty],[AgentId],[FromDate],[ToDate],[WeekNo])"
        strSql = strSql & " select 'E', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0," & SafeSQL(strAgentId) & ",getdate(),getdate(),0 "
        ExecuteSQL(strSql)
        strSql = "insert into ServiceLevelRep ([CustomerClass],[LoadedQty],[FullQty],[SkipQty],[NSRQty],[MiscQty],[MissedQty],[DMissedQty],[TMissedQty],[AlertQty],[CINoQty],[ServicedQty],[AgentId],[FromDate],[ToDate],[WeekNo])"
        strSql = strSql & " select ' ', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0," & SafeSQL(strAgentId) & ",getdate(),getdate(),0 "
        ExecuteSQL(strSql)

        'Update Loaded Qty
        'Sales person visit the customer and there's an invoice. Need to make sure the invoice is not only for exchange
        strSql = " update servicelevelrep set LoadedQty = cnt from ( "
        strSql = strSql & " select  isnull(count(distinct custid),0) as cnt, AgentId, left(isnull(custclass,' '),1) custclass from invoice, invitem,customer where customer.custno = invoice.custid and invdt between " & SafeSQL(Format(dtFromDate, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtToDate, "yyyyMMdd 00:00:00")) & " and invoice.invno = invitem.invno and Invoice.AgentId = " & SafeSQL(strAgentId) & " and invitem.description not like 'EX:%' group by agentid,left(isnull(custclass,' '),1)"
        strSql = strSql & " ) a where a.agentid = servicelevelrep.agentid and a.custclass = servicelevelrep.customerclass "
        ExecuteSQL(strSql)

        'Update Full Qty
        'Sales person visit the customer with Service "Full", check should be no invoice
        strSql = "  update servicelevelrep set Fullqty = cnt from ( "
        strSql = strSql & " select ISNULL(COUNT(DISTINCT CustomerID),0) Cnt, Service.AgentId, left(isnull(CustClass,' '),1) CustClass from Service INNER JOIN Customer ON Service.CustomerID = Customer.CustNo INNER JOIN Reason ON Service.ReasonCode = Reason.Code WHERE Service.ServiceDt between " & SafeSQL(Format(dtFromDate, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtToDate, "yyyyMMdd 00:00:00")) & " and Reason.Description = 'Full' and Service.AgentId = " & SafeSQL(strAgentId) & " and Customer.CustNo+Service.AgentId not in  (Select CustId + AgentId from Invoice where Invoice.Invdt between " & SafeSQL(Format(dtFromDate, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtToDate, "yyyyMMdd 00:00:00")) & " and Invoice.AgentId = " & SafeSQL(strAgentId) & ") group by agentid,left(isnull(custclass,' '),1) "
        strSql = strSql & " ) a where a.agentid = servicelevelrep.agentid and a.custclass = servicelevelrep.customerclass "
        ExecuteSQL(strSql)

        'Update Skip Qty
        'Total of customer that on the schedule but not for that particular week and there's no service, no sales
        strSql = "  update servicelevelrep set SkipQty = cnt from ( "
        strSql = strSql & " SELECT ISNULL(COUNT(DISTINCT CustNo),0) Cnt, SalesAgent as AgentId, left(isnull(CustClass,' '),1) CustClass FROM Customer WHERE SalesAgent = " & SafeSQL(strAgentId)
        strSql = strSql & " AND CustNo  IN ( SELECT custno from routedet, routemaster WHERE routedet.RouteNo = RouteMaster.RouteNo and RouteMaster.VehicleNo = " & SafeSQL(strAgentId) & " ) "
        strSql = strSql & " AND CustNo NOT IN ( SELECT custno from routedet, routemaster WHERE routedet.RouteNo = RouteMaster.RouteNo and RouteMaster.VehicleNo = " & SafeSQL(strAgentId) & " and RouteWeek = " & SafeSQL(sWeek) & " ) "
        strSql = strSql & " AND CustNo + SalesAgent NOT IN (SELECT CustId + AgentId FROM Invoice WHERE Invoice.Invdt BETWEEN " & SafeSQL(Format(dtFromDate, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtToDate, "yyyyMMdd 00:00:00")) & " AND AgentID = " & SafeSQL(strAgentId) & ") "
        strSql = strSql & " AND CustNo + SalesAgent NOT IN (SELECT DISTINCT CustomerID + AgentID FROM Service INNER JOIN Customer ON Service.CustomerID = Customer.CustNo WHERE Service.ServiceDt BETWEEN " & SafeSQL(Format(dtFromDate, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtToDate, "yyyyMMdd 00:00:00")) & " AND AgentID = " & SafeSQL(strAgentId) & ") "
        strSql = strSql & " group by SalesAgent, left(isnull(CustClass,' '),1) "
        strSql = strSql & " ) a where a.agentid = servicelevelrep.agentid and a.custclass = servicelevelrep.customerclass "
        ExecuteSQL(strSql)

        'Update NSR Qty
        'Sales person visit the customer with Service "NSR", check should be no invoice
        strSql = "  update servicelevelrep set NSRQty = cnt from ( "
        strSql = strSql & " select ISNULL(COUNT(DISTINCT CustomerID),0) Cnt, Service.AgentId, left(isnull(CustClass,' '),1) CustClass from Service INNER JOIN Customer ON Service.CustomerID = Customer.CustNo INNER JOIN Reason ON Service.ReasonCode = Reason.Code WHERE Service.ServiceDt between " & SafeSQL(Format(dtFromDate, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtToDate, "yyyyMMdd 00:00:00")) & " and Reason.Description = 'NSR' and Service.AgentId = " & SafeSQL(strAgentId) & " and Customer.CustNo+Service.AgentId not in  (Select CustId + AgentId from Invoice where Invoice.Invdt between " & SafeSQL(Format(dtFromDate, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtToDate, "yyyyMMdd 00:00:00")) & " and Invoice.AgentId = " & SafeSQL(strAgentId) & ") and CustomerID + AgentID not in (select Service.CustomerID + Service.AgentId from Service INNER JOIN Customer ON Service.CustomerID = Customer.CustNo INNER JOIN Reason ON Service.ReasonCode = Reason.Code WHERE Service.ServiceDt between " & SafeSQL(Format(dtFromDate, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtToDate, "yyyyMMdd 00:00:00")) & " and Reason.Description = 'Full' and Service.AgentId = " & SafeSQL(strAgentId) & ") group by agentid,left(isnull(custclass,' '),1)"
        strSql = strSql & " ) a where a.agentid = servicelevelrep.agentid and a.custclass = servicelevelrep.customerclass "
        ExecuteSQL(strSql)

        'Update Misc Qty
        'sales person visit the customer no sales, and no service with code 'FULL' and 'NSR'
        strSql = "  update servicelevelrep set MiscQty = cnt from ( "
        strSql = strSql & " select ISNULL(COUNT(DISTINCT CustomerID),0) Cnt, Service.AgentId, left(isnull(CustClass,' '),1) CustClass from Service INNER JOIN Customer ON Service.CustomerID = Customer.CustNo LEFT JOIN Reason ON Service.ReasonCode = Reason.Code WHERE Service.ServiceDt between " & SafeSQL(Format(dtFromDate, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtToDate, "yyyyMMdd 00:00:00")) & " and Reason.Description not in  ('Full','NSR') and Service.AgentId = " & SafeSQL(strAgentId) & " and Customer.CustNo+Service.AgentId not in  (Select CustId + AgentId from Invoice where Invoice.Invdt between " & SafeSQL(Format(dtFromDate, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtToDate, "yyyyMMdd 00:00:00")) & " and Invoice.AgentId = " & SafeSQL(strAgentId) & ") "
        strSql = strSql & " and CustomerID + AgentID not in (select Service.CustomerID + Service.AgentId from Service INNER JOIN Customer ON Service.CustomerID = Customer.CustNo INNER JOIN Reason ON Service.ReasonCode = Reason.Code WHERE Service.ServiceDt between " & SafeSQL(Format(dtFromDate, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtToDate, "yyyyMMdd 00:00:00")) & " and Reason.Description in ('Full','NSR') and Service.AgentId = " & SafeSQL(strAgentId) & ") "
        strSql = strSql & " group by agentid,left(isnull(custclass,' '),1) "
        strSql = strSql & " ) a where a.agentid = servicelevelrep.agentid and a.custclass = servicelevelrep.customerclass "
        ExecuteSQL(strSql)

        'Update Missed Qty
        'total of customer that the sales person not visit for 1 time
        strSql = "  update servicelevelrep set MissedQty = cnt from ( "
        strSql = strSql & " select ISNULL(COUNT(DISTINCT DailySchedule.CustNo),0) Cnt, SalesAgent as AgentId, left(isnull(CustClass,' '),1) CustClass from DailySchedule, Customer WHERE DailySchedule.CustNo = Customer.CustNo "
        strSql = strSql & " AND DailySchedule.CustNo + SalesAgent NOT IN (SELECT CustId + AgentId FROM Invoice WHERE Invoice.Invdt BETWEEN '20090613 00:00:00' and '20090618 23:59:59') "
        strSql = strSql & " AND DailySchedule.CustNo + SalesAgent NOT IN (SELECT DISTINCT CustomerID + AgentID FROM Service INNER JOIN Customer ON Service.CustomerID = Customer.CustNo WHERE Service.ServiceDt BETWEEN '20090613 00:00:00' and '20090618 23:59:59') "
        strSql = strSql & " AND DailySchedule.NoofMissed = 1 "
        strSql = strSql & " group by SalesAgent, left(isnull(CustClass,' '),1) "
        strSql = strSql & " ) a where a.agentid = servicelevelrep.agentid and a.custclass = servicelevelrep.customerclass "
        ExecuteSQL(strSql)

        'Update D. Missed Qty
        'total of customer that the sales person not visit for 2 times
        strSql = "  update servicelevelrep set DMissedQty = cnt from ( "
        strSql = strSql & " select ISNULL(COUNT(DISTINCT DailySchedule.CustNo),0) Cnt, SalesAgent as AgentId, left(isnull(CustClass,' '),1) CustClass from DailySchedule, Customer WHERE DailySchedule.CustNo = Customer.CustNo "
        strSql = strSql & " AND DailySchedule.CustNo + SalesAgent NOT IN (SELECT CustId + AgentId FROM Invoice WHERE Invoice.Invdt BETWEEN '20090613 00:00:00' and '20090618 23:59:59') "
        strSql = strSql & " AND DailySchedule.CustNo + SalesAgent NOT IN (SELECT DISTINCT CustomerID + AgentID FROM Service INNER JOIN Customer ON Service.CustomerID = Customer.CustNo WHERE Service.ServiceDt BETWEEN '20090613 00:00:00' and '20090618 23:59:59') "
        strSql = strSql & " AND DailySchedule.NoofMissed = 2 "
        strSql = strSql & " group by SalesAgent, left(isnull(CustClass,' '),1) "
        strSql = strSql & " ) a where a.agentid = servicelevelrep.agentid and a.custclass = servicelevelrep.customerclass "
        ExecuteSQL(strSql)


        'Update T. Missed Qty
        'total of customer that the sales person not visit for 3 times
        strSql = "  update servicelevelrep set TMissedQty = cnt from ( "
        strSql = strSql & " select ISNULL(COUNT(DISTINCT DailySchedule.CustNo),0) Cnt, SalesAgent as AgentId, left(isnull(CustClass,' '),1) CustClass from DailySchedule, Customer WHERE DailySchedule.CustNo = Customer.CustNo "
        strSql = strSql & " AND DailySchedule.CustNo + SalesAgent NOT IN (SELECT CustId + AgentId FROM Invoice WHERE Invoice.Invdt BETWEEN '20090613 00:00:00' and '20090618 23:59:59') "
        strSql = strSql & " AND DailySchedule.CustNo + SalesAgent NOT IN (SELECT DISTINCT CustomerID + AgentID FROM Service INNER JOIN Customer ON Service.CustomerID = Customer.CustNo WHERE Service.ServiceDt BETWEEN '20090613 00:00:00' and '20090618 23:59:59') "
        strSql = strSql & " AND DailySchedule.NoofMissed = 3 "
        strSql = strSql & " group by SalesAgent, left(isnull(CustClass,' '),1) "
        strSql = strSql & " ) a where a.agentid = servicelevelrep.agentid and a.custclass = servicelevelrep.customerclass "
        ExecuteSQL(strSql)


        'Update Alert Qty
        'total of customer that the sales person not visit > 3 times       
        strSql = "  update servicelevelrep set ALertQty = cnt from ( "
        strSql = strSql & " select ISNULL(COUNT(DISTINCT DailySchedule.CustNo),0) Cnt, SalesAgent as AgentId, left(isnull(CustClass,' '),1) CustClass from DailySchedule, Customer WHERE DailySchedule.CustNo = Customer.CustNo "
        strSql = strSql & " AND DailySchedule.CustNo + SalesAgent NOT IN (SELECT CustId + AgentId FROM Invoice WHERE Invoice.Invdt BETWEEN '20090613 00:00:00' and '20090618 23:59:59') "
        strSql = strSql & " AND DailySchedule.CustNo + SalesAgent NOT IN (SELECT DISTINCT CustomerID + AgentID FROM Service INNER JOIN Customer ON Service.CustomerID = Customer.CustNo WHERE Service.ServiceDt BETWEEN '20090613 00:00:00' and '20090618 23:59:59') "
        strSql = strSql & " AND DailySchedule.NoofMissed > 1 "
        strSql = strSql & " group by SalesAgent, left(isnull(CustClass,' '),1) "
        strSql = strSql & " ) a where a.agentid = servicelevelrep.agentid and a.custclass = servicelevelrep.customerclass "
        ExecuteSQL(strSql)


        'Update CI No Qty
        'Total of the customer that belong the the sales person but not in the schedule and there's no sales or service at all
        strSql = "  update servicelevelrep set CINoQty = cnt from ( "
        strSql = strSql & " SELECT ISNULL(COUNT(DISTINCT CustNo),0) Cnt, SalesAgent as AgentId, left(isnull(CustClass,' '),1) CustClass FROM Customer WHERE SalesAgent = " & SafeSQL(strAgentId)
        strSql = strSql & " AND CustNo NOT IN ( SELECT custno from routedet, routemaster WHERE routedet.RouteNo = RouteMaster.RouteNo and RouteMaster.VehicleNo = " & SafeSQL(strAgentId) & " ) "
        strSql = strSql & " AND CustNo + SalesAgent NOT IN (SELECT CustId + AgentId FROM Invoice WHERE Invoice.Invdt BETWEEN " & SafeSQL(Format(dtFromDate, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtToDate, "yyyyMMdd 00:00:00")) & " AND AgentID = " & SafeSQL(strAgentId) & ") "
        strSql = strSql & " AND CustNo + SalesAgent NOT IN (SELECT DISTINCT CustomerID + AgentID FROM Service INNER JOIN Customer ON Service.CustomerID = Customer.CustNo WHERE Service.ServiceDt BETWEEN " & SafeSQL(Format(dtFromDate, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtToDate, "yyyyMMdd 00:00:00")) & " AND AgentID = " & SafeSQL(strAgentId) & ") "
        strSql = strSql & " group by SalesAgent, left(isnull(CustClass,' '),1) "
        strSql = strSql & " ) a where a.agentid = servicelevelrep.agentid and a.custclass = servicelevelrep.customerclass "
        ExecuteSQL(strSql)



    End Sub
End Class