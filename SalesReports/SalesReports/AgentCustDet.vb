Imports System.Data
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Public Class AgentCustDet
    Implements ISalesBase
    
    Private aAgent As New ArrayList()
    Dim strSql As String = ""


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

    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        btnPrint.Enabled = False
        ExecuteSQL("Delete AgentCustDetRpt")
        'ExecuteSQL("Insert AgentCustDetRpt(AgentId, AgentName, CustNo, CustName, TimeIn, TimeOut,FromDate) Select CV.AgentId, A.[Name], C.CustNo, C.CustName, Min(CV.TransDate) as TimeIn, " & _
        '"Max(CV.TransDate) as TimeOut, " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd")) & " from CustVisit CV, SalesAgent A, " & _
        '"Customer C where A.Code = CV.AgentId and CV.AgentId like " & SafeSQL(cmbAgent.SelectedValue.ToString) & " and CV.TransDate between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00.000")) & " and " & _
        'SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59.000")) & " and (CV.TransType = 'CIN' or CV.TransType = 'COUT' or CV.TransType = 'ORDER' or " & _
        '"CV.TransType = 'RECEIPT' or CV.TransType = 'SERVICE') " & _
        '"and CV.CustId = C.CustNo group by C.CustNo, C.CustName, CV.AgentId, A.[Name]")


        ''''''New Cust
        ExecuteSQL("Insert AgentCustDetRpt(AgentId, AgentName, CustNo, CustName, TimeIn, TimeOut,FromDate) Select CV.AgentId, A.[Name], C.CustNo, C.CustName, Min(CV.TransDate) as TimeIn, " & _
        "Max(CV.TransDate) as TimeOut, " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd")) & " from CustVisit CV, SalesAgent A, " & _
        "Customer C where A.Code = CV.AgentId and CV.AgentId like " & SafeSQL(cmbAgent.SelectedValue.ToString) & " and CV.TransDate between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00.000")) & " and " & _
        SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59.000")) & " and (CV.TransType = 'CIN' or CV.TransType = 'COUT' or CV.TransType = 'ORDER' or " & _
        "CV.TransType = 'RECEIPT' or CV.TransType = 'SERVICE') " & _
        "and CV.CustId = C.CustNo group by C.CustNo, C.CustName, CV.AgentId, A.[Name] union Select CV.AgentId, A.[Name], C.CustID, C.CustName + ' - (NEW)', Min(CV.TransDate) as TimeIn, " & _
        "Max(CV.TransDate) as TimeOut, " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd")) & " from CustVisit CV, SalesAgent A, " & _
        "NewCust C where A.Code = CV.AgentId and CV.AgentId like " & SafeSQL(cmbAgent.SelectedValue.ToString) & " and CV.TransDate between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00.000")) & " and " & _
        SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59.000")) & " and (CV.TransType = 'CIN' or CV.TransType = 'COUT' or CV.TransType = 'ORDER' or " & _
        "CV.TransType = 'RECEIPT' or CV.TransType = 'SERVICE') " & _
        "and CV.CustId = C.CustID group by C.CustID, C.CustName, CV.AgentId, A.[Name]")
        ''''''

        ExecuteSQL("Update AgentCustDetRpt set OrderAmt = (select sum(TotalAmt) from orderhdr where " & _
        "(OrderHdr.Void is null or OrderHdr.Void = 0) and OrderHdr.AgentId = AgentCustDetRpt.AgentID and " & _
        "OrderHdr.CustId = AgentCustDetRpt.CustNo and OrdDt between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00.000")) & _
        " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59.000")) & ") , PaymentCollected = (select sum(Amount) from Receipt where " & _
        "(Receipt.Void is null or Receipt.Void = 0) and Receipt.AgentId = AgentCustDetRpt.AgentID and Receipt.CustId = AgentCustDetRpt.CustNo and " & _
        "RcptDt between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00.000")) & " and " & _
        SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59.000")) & "), Services = " & _
        "dbo.fn_AgentCustDet_Services(" & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00.000")) & "," & _
        SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59.000")) & ", AgentCustDetRpt.CustNo, AgentCustDetRpt.AgentID" & ")")

        ExecuteSQL("Update AgentCustDetRpt set OrderAmt = 0 where OrderAmt is null")

        ExecuteSQL("Update AgentCustDetRpt set PaymentCollected = 0 where PaymentCollected is null")

        ExecuteReport("Select * from AgentCustDetRpt Order By TimeIn asc", "AgentCustDetRpt")
        btnPrint.Enabled = True
    End Sub

    Private Sub AgentCustDet_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub

    Private Sub AgentCustDet_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ConnectDB()
        loadCombo()
    End Sub

    Public Sub loadCombo()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Code, Name from SalesAgent")
        cmbAgent.DataSource = Nothing
        aAgent.Clear()
        aAgent.Add(New ComboValues("%", "ALL"))
        While dtr.Read()
            aAgent.Add(New ComboValues(dtr("Code").ToString, dtr("Name").ToString))
        End While
        dtr.Close()
        cmbAgent.DataSource = aAgent
        cmbAgent.DisplayMember = "Desc"
        cmbAgent.ValueMember = "Code"
        cmbAgent.SelectedIndex = 0
    End Sub

End Class