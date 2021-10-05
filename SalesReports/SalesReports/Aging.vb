Imports System.Data
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Public Class Aging
    Implements ISalesBase


    Dim strPay As String = " "
    Dim strTerms As String = " "
    Private Sub Aging_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ConnectDB()
        'Dim strSqlSales = "SELECT DateDiff('d',Invoice.InvDt,Now()) AS Age, Customer.CustNo, Customer.CustName, Invoice.AgentID, Invoice.InvNo, Invoice.InvDt, Invoice.TotalAmt, Invoice.PaidAmt, Invoice.TotalAmt-Invoice.PaidAmt AS OutstandingAmt INTO Aging IN 'C:\Documents and Settings\Jagadish\My Documents\Visual Studio 2005\Projects\Reports\Reports\bin\Debug\Report.mdb' FROM Customer INNER JOIN Invoice ON Customer.CustNo=Invoice.CustId;"
        'ExecuteSQL(strSqlSales)
        ' loadCombo()
        'DisconnectDB()
        Dim Terms As Integer
        Dim dtr As SqlDataReader
        'Terms = CInt((InStr(1, dtr("PayTerms"), "D"))) - 1
        dtr = ReadRecord("SELECT Invoice.InvNo, Invoice.InvDt, Invoice.DoNo, Invoice.DoDt, Invoice.OrdNo, Invoice.CustId, Invoice.AgentId, Invoice.Discount, Invoice.SubTotal, Invoice.GST, Invoice.GstAmt, Invoice.TotalAmt, Invoice.PaidAmt, Invoice.PrintNo, Invoice.Void, Invoice.VoidDate, Invoice.PayTerms, Invoice.TermDays, Invoice.CurCode, Invoice.CurExRate, Invoice.PoNo, Invoice.Exported, Invoice.ExportDate, Invoice.CustRefNo, Invoice.DTG, PayTerms.DueDateCalc FROM Invoice INNER JOIN  PayTerms ON Invoice.PayTerms = PayTerms.Code WHERE (Invoice.TotalAmt - Invoice.PaidAmt > 1)")
        Do While dtr.Read = True
            Terms = CInt(Mid(dtr("PayTerms"), 1, CInt((InStr(1, dtr("DueDateCalc"), "D"))) - 1))
        Loop
    End Sub

    Public Sub loadCombo()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select distinct(AgentId) from Aging")
        cmbAgent.Items.Add("All")
        Do While dtr.Read = True
            cmbAgent.Items.Add(dtr(0))
        Loop
        dtr.Close()
        If cmbAge.Items.Count > 0 Then
            cmbAge.SelectedIndex = 0
        End If
        If cmbAgent.Items.Count > 0 Then
            cmbAgent.SelectedIndex = 0
        End If
        dtpDate.Value = Date.Now
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Select Case cmbAge.SelectedIndex
            Case 0
                strTerms = " "
            Case 1
                If cmbAgent.Text = "All" Then
                    strTerms = "((Age>0) And (Age<=30))"
                Else
                    strTerms = " and ((Age>0) And (Age<=30))"

                End If

            Case 2
                If cmbAgent.Text = "All" Then
                    strTerms = "((Age>30) And (Age<=60))"
                Else
                    strTerms = " and ((Age>30) And (Age<=60))"

                End If

            Case 3
                If cmbAgent.Text = "All" Then
                    strTerms = "((Age>60) And (Age<=90))"
                Else
                    strTerms = " and ((Age>60) And (Age<=90))"

                End If
            Case 4
                If cmbAgent.Text = "All" Then
                    strTerms = "((Age>90) And (Age<=120))"
                Else
                    strTerms = " and ((Age>90) And (Age<=120))"

                End If
            Case 5
                If cmbAgent.Text = "All" Then
                    strTerms = " (Age > 120) "
                Else
                    strTerms = " and (Age > 120) "
                End If
        End Select
        Dim strSql As String = "Select * from Aging where " & strPay & strTerms & " (OutstandingAmt > 1) ORDER BY Age"
        ' MsgBox(strSql)
        ExecuteReport(strSql, "AgingRep")
    End Sub

    Private Sub cmbAgent_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAgent.SelectedIndexChanged
        If cmbAgent.Text = "All" Then
            strPay = " "
        Else
            strPay = "  Agentid=" & "'" & cmbAgent.Text & "' and"
        End If
    End Sub

    Private Sub cmbAge_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAge.SelectedIndexChanged

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
