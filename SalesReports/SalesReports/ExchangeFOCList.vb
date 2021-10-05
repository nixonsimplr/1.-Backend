Imports System.Data
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales

Public Class ExchangeFOCList
    Implements ISalesBase
    Private aAgent As New ArrayList()
    Dim sAgent As String = " "

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


    Private Sub ExchangeFOCList_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub

    Private Sub ExchangeFOCList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ConnectDB()
        LoadAgent()
    End Sub

    Public Sub LoadAgent()
        'Dim dtr As SqlDataReader
        'dtr = ReadRecord("Select Distinct AgentID from MDT order by AgentID")
        'cmbAgent.Items.Add("ALL")
        'Do While dtr.Read = True
        '    cmbAgent.Items.Add(dtr(0))
        'Loop
        'dtr.Close()
        'If cmbAgent.Items.Count > 0 Then
        '    cmbAgent.SelectedIndex = 0
        'End If
        'dtpFromDate.Value = Date.Now
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Code, Name from SalesAgent")
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

    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        If cmbAgent.Text = "ALL" Then
            sAgent = " "
        Else
            sAgent = "and OrderHdr.Agentid=" & "'" & cmbAgent.SelectedValue & "'"
        End If

        Dim strSql As String = ""
        'strSql = strSql & " Select Invoice.InvNo as TransNo, Invoice.InvDt as TransDate, Invoice.CustID as CustNo, Customer.CustName, Item.ItemNo, Item.ItemName, InvItem.Qty, InvItem.UOM, Reason.Description as Reason, SalesAgent.Name as AgentID, 'EXCHANGE' as ItemType from InvItem inner join Invoice on Invoice.InvNo = InvItem.InvNo inner join Item on Item.ItemNo = InvItem.ItemNo inner join SalesAgent on Invoice.AgentId = SalesAgent.Code  inner join Customer on Invoice.CustId = Customer.CustNo left join Reason on Reason.Code = InvItem.ReasonCode and Reason.ReasonType = 'EX' where  InvItem.Description like 'EX:%' and Invoice.InvDt between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & sAgent
        'strSql = strSql & " union "
        'strSql = strSql & "  Select Invoice.InvNo as TransNo, Invoice.InvDt as TransDate, Invoice.CustID as CustNo, Customer.CustName, Item.ItemNo, Item.ItemName, InvItem.Qty, InvItem.UOM, Reason.Description as Reason, SalesAgent.Name as AgentID, 'EXCHANGE' as ItemType from InvItem inner join Invoice on Invoice.InvNo = InvItem.InvNo inner join Item on Item.ItemNo = InvItem.ItemNo inner join SalesAgent on Invoice.AgentId = SalesAgent.Code  inner join Customer on Invoice.CustId = Customer.CustNo left join Reason on Reason.Code = InvItem.ReasonCode and Reason.ReasonType = 'FOC' where  InvItem.Description like 'FOC:%' and Invoice.InvDt between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & sAgent & " order by Invoice.InvNo, Item.ItemName"
        strSql = strSql & " Select OrderHdr.OrdNo as TransNo, OrderHdr.OrdDt as TransDate, OrderHdr.CustId as CustNo, Customer.CustName, Item.ItemNo, Item.ItemName, OrdItem.Qty, OrdItem.UOM, Reason.Description as Reason, SalesAgent.Name as AgentID, 'EXCHANGE' as ItemType, OrderHdr.CompanyName from OrdItem inner join OrderHdr on OrderHdr.OrdNo = OrdItem.OrdNo inner join Item on Item.CompanyName = OrderHdr.CompanyName and Item.ItemNo = OrdItem.ItemNo inner join SalesAgent on OrderHdr.AgentId = SalesAgent.Code  inner join Customer on OrderHdr.CompanyName = Customer.CompanyName and OrderHdr.CustId = Customer.CustNo left join Reason on Reason.Code = OrdItem.ReasonCode and Reason.ReasonType = 'EX' where  OrdItem.Description like 'EX:%' and OrderHdr.OrdDt between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 23:59:59")) & sAgent
        strSql = strSql & " union "
        strSql = strSql & " Select OrderHdr.OrdNo as TransNo, OrderHdr.OrdDt as TransDate, OrderHdr.CustId as CustNo, Customer.CustName, Item.ItemNo, Item.ItemName, OrdItem.Qty, OrdItem.UOM, Reason.Description as Reason, SalesAgent.Name as AgentID, 'FOC' as ItemType, OrderHdr.CompanyName from OrdItem inner join OrderHdr on OrderHdr.OrdNo = OrdItem.OrdNo inner join Item on Item.CompanyName = OrderHdr.CompanyName and Item.ItemNo = OrdItem.ItemNo inner join SalesAgent on OrderHdr.AgentId = SalesAgent.Code  inner join Customer on OrderHdr.CompanyName = Customer.CompanyName and OrderHdr.CustId = Customer.CustNo left join Reason on Reason.Code = OrdItem.ReasonCode and Reason.ReasonType = 'FOC' where  OrdItem.Description like 'FOC:%' and OrderHdr.OrdDt between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 23:59:59")) & sAgent
        strSql = strSql & " union "
        strSql = strSql & " Select OrderHdr.OrdNo as TransNo, OrderHdr.OrdDt as TransDate, OrderHdr.CustId as CustNo, NewCust.CustName, Item.ItemNo, Item.ItemName, OrdItem.Qty, OrdItem.UOM, Reason.Description as Reason, SalesAgent.Name as AgentID, 'EXCHANGE' as ItemType, OrderHdr.CompanyName from OrdItem inner join OrderHdr on OrderHdr.OrdNo = OrdItem.OrdNo inner join Item on Item.CompanyName = OrderHdr.CompanyName and Item.ItemNo = OrdItem.ItemNo inner join SalesAgent on OrderHdr.AgentId = SalesAgent.Code  inner join NewCust on OrderHdr.CompanyName = NewCust.CompanyName and OrderHdr.CustId = NewCust.CustId left join Reason on Reason.Code = OrdItem.ReasonCode and Reason.ReasonType = 'EX' where  OrdItem.Description like 'EX:%' and OrderHdr.OrdDt between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 23:59:59")) & sAgent
        strSql = strSql & " union "
        strSql = strSql & " Select OrderHdr.OrdNo as TransNo, OrderHdr.OrdDt as TransDate, OrderHdr.CustId as CustNo, NewCust.CustName, Item.ItemNo, Item.ItemName, OrdItem.Qty, OrdItem.UOM, Reason.Description as Reason, SalesAgent.Name as AgentID, 'FOC' as ItemType, OrderHdr.CompanyName from OrdItem inner join OrderHdr on OrderHdr.OrdNo = OrdItem.OrdNo inner join Item on Item.CompanyName = OrderHdr.CompanyName and Item.ItemNo = OrdItem.ItemNo inner join SalesAgent on OrderHdr.AgentId = SalesAgent.Code  inner join NewCust on OrderHdr.CompanyName = NewCust.CompanyName and OrderHdr.CustId = NewCust.CustId left join Reason on Reason.Code = OrdItem.ReasonCode and Reason.ReasonType = 'FOC' where  OrdItem.Description like 'FOC:%' and OrderHdr.OrdDt between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 23:59:59")) & sAgent & " order by OrderHdr.OrdNo, Item.ItemName"

        'System.IO.File.AppendAllText("C:\Report.txt", strSql)
        ExecuteReport(strSql, "ExchangeFOCListRep")
    End Sub
End Class