Imports SalesInterface.MobileSales
Imports System.Data.SqlClient

Public Class OutstandingPayment
    Implements ISalesBase

    Private aTerms As New ArrayList()
    Private aAgent As New ArrayList()
    Dim strPay As String = " "
    Dim strPTerms As String
    Dim strTerms As String = " "
    Private Sub cmbAge_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAge.SelectedIndexChanged
        Select Case cmbAge.SelectedIndex
            Case 0
                strTerms = " "
            Case 1
                strTerms = " ((Age>0) And (Age<=30)) and"
            Case 2
                strTerms = " ((Age>30) And (Age<=60)) and"
            Case 3
                strTerms = " ((Age>60) And (Age<=90)) and"
            Case 4
                strTerms = " ((Age>90) And (Age<=120)) and"
            Case 5
                strTerms = " (age > 120) and"
        End Select
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        btnPrint.Enabled = False
        If cmbAgent.Text = "ALL" Then
            strPay = " "
        Else
            strPay = "and Invoice.Agentid=" & "'" & cmbAgent.SelectedValue & "'"
        End If
        If cmbTerms.Text = "ALL" Then
            strPTerms = ""
        Else
            strPTerms = " and Invoice.Payterms='" & cmbTerms.SelectedValue & "'"
        End If
        Dim strDel = "drop table Outstanding"
        ExecuteSQL(strDel)
        Dim strSqlSales = "SELECT DateDiff(d,Invoice.InvDt,GetDate()) AS Age, Customer.CustNo, Customer.CustName, Invoice.AgentID, Invoice.InvNo, Invoice.InvDt, Invoice.TotalAmt, Invoice.PaidAmt, Invoice.TotalAmt-Invoice.PaidAmt AS OutstandingAmt INTO Outstanding  FROM Customer INNER JOIN Invoice ON Customer.CompanyName = Invoice.CompanyName and Customer.CustNo=Invoice.CustId where Invoice.Void=0 " & strPay & strPTerms
        ExecuteSQL(strSqlSales)
        Dim strSql As String = "SELECT Outstanding.CustNo, Outstanding.CustName, Outstanding.Age, Outstanding.InvNo, Outstanding.InvDt, Outstanding.OutstandingAmt FROM Outstanding where " & strTerms & " (OutstandingAmt > 1) ORDER BY CustName"
        ExecuteReport(strSql, "OutstandingRep")
        btnPrint.Enabled = True
    End Sub

    Private Sub OutstandingPayment_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub
    Public Sub LoadDescription()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Code, Description from PayTerms order by Description")
        cmbTerms.DataSource = Nothing
        aTerms.Clear()
        aTerms.Add(New ComboValues("ALL", "ALL"))
        While dtr.Read()
            aTerms.Add(New ComboValues(dtr("Code").ToString, dtr("Description").ToString))

            '    iSelIndex = iIndex
            'End If
            'iIndex = iIndex + 1
        End While
        dtr.Close()
        cmbTerms.DataSource = aTerms
        cmbTerms.DisplayMember = "Desc"
        cmbTerms.ValueMember = "Code"
        cmbTerms.SelectedIndex = 0
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
    Private Sub OutstandingPayment_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ConnectDB()
        LoadDescription()
        loadCombo()
        If cmbAge.Items.Count > 0 Then
            cmbAge.SelectedIndex = 0
        End If
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