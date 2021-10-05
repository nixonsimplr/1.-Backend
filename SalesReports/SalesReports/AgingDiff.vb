Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Public Class AgingDiff
    Implements ISalesBase


    Dim strPay As String = " "
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim strSql As String
        If cmbAgent.Text = "All" Then
            strSql = "SELECT " & SafeSQL(cmbAgent.Text) & " as AgentID, PaidAmt, TotalAmt, InvDt, InvNo, total, Above120, Days120, Days90, Days60, Days30, CustName, CustID FROM AgingSum"
        Else
            strSql = "SELECT " & SafeSQL(cmbAgent.Text) & " as AgentID, PaidAmt, TotalAmt, InvDt, InvNo, total, Above120, Days120, Days90, Days60, Days30, CustName, CustID FROM AgingSum Where AgentID= " & SafeSQL(cmbAgent.Text)
        End If
        ExecuteReport(strSql, "AgingRep")
    End Sub
    Public Sub loadCombo()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select AgentID from MDT")
        cmbAgent.Items.Add("All")
        Do While dtr.Read = True

            cmbAgent.Items.Add(dtr(0))

        Loop
        dtr.Close()
        If cmbAgent.Items.Count > 0 Then
            cmbAgent.SelectedIndex = 0
        End If

    End Sub
    Private Sub AgingDiff_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        DisconnectDB()
    End Sub

    Private Sub AgingDiff_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ConnectDB()
        ExecuteSQL("Truncate Table AgingSum")
        Dim dtr As SqlDataReader
        Dim sCustId, sCustName, strIns As String

        Dim lday As Long
        Dim d30, d60, d90, d120, d120a, total As Double

        Dim Terms As Integer
        dtr = ReadRecordAnother("SELECT Invoice.Payterms, Invoice.CustId, Customer.CustName, Invoice.InvNo, Invoice.InvDt, Invoice.TotalAmt, Invoice.PaidAmt, PayTerms.DueDateCalc,Invoice.AgentID FROM Invoice INNER JOIN Customer ON Invoice.CustId = Customer.CustNo INNER JOIN PayTerms ON Invoice.PayTerms = PayTerms.Code WHERE(Invoice.TotalAmt > Invoice.PaidAmt) and Invoice.Void =0 ORDER BY Invoice.CustId")

        While dtr.Read = True
            Terms = CInt(Mid(dtr("PayTerms"), 1, CInt((InStr(1, dtr("DueDateCalc"), "D"))) - 1))
            lday = DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, Terms, dtr("InvDt")), Now.Date)
            If lday <= 0 Then
                d30 = 0
                d60 = 0
                d90 = 0
                d120 = 0
                d120a = 0
            ElseIf lday <= 30 Then
                d30 = dtr("TotalAmt") - dtr("PaidAmt")
                d60 = 0
                d90 = 0
                d120 = 0
                d120a = 0
            ElseIf lday <= 60 Then
                d30 = 0
                d60 = dtr("TotalAmt") - dtr("PaidAmt")
                d90 = 0
                d120 = 0
                d120a = 0
            ElseIf lday <= 90 Then
                d30 = 0
                d60 = 0
                d90 = dtr("TotalAmt") - dtr("PaidAmt")
                d120 = 0
                d120a = 0
            ElseIf lday <= 120 Then
                d30 = 0
                d60 = 0
                d90 = 0
                d120 = dtr("TotalAmt") - dtr("PaidAmt")
                d120a = 0
            Else
                d30 = 0
                d60 = 0
                d90 = 0
                d120 = 0
                d120a = dtr("TotalAmt") - dtr("PaidAmt")
            End If
            total = dtr("TotalAmt")
            sCustId = dtr("CustId")
            sCustName = Convert.ToString(dtr("CustName"))
            strIns = "Insert into AgingSum (CustID, CustName, Days30, Days60, Days90, Days120, Above120, total,InvNo,InvDt,TotalAmt,PaidAmt,AgentID) Values" & " ('" & sCustId & "','" & sCustName & "'," & d30 & "," & d60 & "," & d90 & "," & d120 & "," & d120a & "," & d30 + d60 + d90 + d120 + d120a & "," & SafeSQL(dtr("InvNo")) & "," & SafeSQL(dtr("InvDt")) & "," & dtr("TotalAmt") & "," & dtr("PaidAmt") & "," & SafeSQL(dtr("AgentID")) & ")"
            ExecuteSQL(strIns)



            '       ExecuteSQL("Update AgingSum (CustID, CustName, 30Days, 60Days, 90Days, 120Days, Above120) Values ()"

        End While
        dtr.Close()
        loadCombo()
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