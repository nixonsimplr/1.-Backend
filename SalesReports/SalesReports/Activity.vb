Imports System.Data
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Public Class Activity
    Implements ISalesBase
    Private aAgent As New ArrayList()
    Dim sAgent As String = " "

    Private Sub Activity_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub

    Private Sub Activity_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ConnectDB()
        LoadAgent()
    End Sub
    Public Sub LoadAgent()
        'Dim dtr As SqlDataReader
        'dtr = ReadRecord("Select distinct AgentID from MDT Order By AgentID")
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

    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        
        If cmbAgent.Text = "ALL" Then
            sAgent = " "
        Else
            sAgent = "and Agentid=" & "'" & cmbAgent.SelectedValue & "'"
        End If
        ExecuteSQL("Delete from ActivityRep")
        Dim strSql As String = "Insert into ActivityRep([CustID],[CustName],[TransNo],[TransType],[TransDate],[AgentID], [Addr], [SalesAgent], [FromDate], [ToDate]) SELECT distinct CustVisit.CustID, Customer.CustName, CustVisit.TransNo, CustVisit.TransType, CustVisit.TransDate, SalesAgent.Name, Customer.[Address], SalesAgent2.Name, " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " as FromDate, " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & " ToDate FROM Customer, CustVisit, SalesAgent, SalesAgent as SalesAgent2 where CustVisit.AgentID = SalesAgent.Code and Customer.CustNo = CustVisit.CustId and Customer.SalesAgent = SalesAgent2.Code and TransDate Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & sAgent & " Order by TransDate"
        ExecuteSQL(strSql)
        strSql = "Update ActivityRep set Amt = OrderHdr.TotalAmt from OrderHdr where OrderHdr.OrdNo = ActivityRep.TransNo and ActivityRep.TransType = 'ORDER'"
        ExecuteSQL(strSql)
        strSql = "Update ActivityRep set Amt = Invoice.TotalAmt from Invoice where Invoice.InvNo = ActivityRep.TransNo and ActivityRep.TransType = 'INVOICE'"
        ExecuteSQL(strSql)
        'strSql = "Update ActivityRep set Amt = OrderHdr.TotalAmt where GoodsReturn.OrdNo = ActivityRep.TransNo and ActivityRep.TransType = 'RETURN'"
        strSql = "Update ActivityRep set Amt = CreditNote.TotalAmt from CreditNote where CreditNote.CreditNoteNo = ActivityRep.TransNo and ActivityRep.TransType = 'CREDIT NOTE'"
        ExecuteSQL(strSql)
        strSql = "Select Distinct * from ActivityRep Order By TransDate"
        ExecuteReport(strSql, "ActivityRep")
    End Sub

    Public Sub ReturnSearch(ByVal SQL As String) Implements SalesInterface.MobileSales.ISalesBase.ReturnSearch

    End Sub

    Public Event SearchField(ByVal FormName As String, ByVal FieldName As String, ByVal FieldType As String, ByVal CurValue As String) Implements SalesInterface.MobileSales.ISalesBase.SearchField

    Public Sub SetCulture(ByVal CultureName As String) Implements SalesInterface.MobileSales.ISalesBase.SetCulture

    End Sub
End Class