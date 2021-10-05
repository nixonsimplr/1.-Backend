Imports System.Data
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Public Class Effective
    Implements ISalesBase

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        btnPrint.Enabled = False
        ExecuteSQL("Delete from Effective")
        Dim strSql As String
        Dim dtr As SqlDataReader
        Dim iCnt As Integer = 0
        dtr = ReadRecord("SELECT COUNT(Customer.SalesAgent) as cnt, SalesAgent.Name FROM  RouteMaster INNER JOIN RouteDet ON RouteMaster.RouteNo = RouteDet.RouteNo INNER JOIN Customer ON RouteDet.CustNo = Customer.CustNo INNER JOIN SalesAgent ON Customer.SalesAgent = SalesAgent.Code WHERE (RouteMaster.RouteDay = " & dtpDate.Value.DayOfWeek & ") GROUP BY SalesAgent.Name")
        While dtr.Read = True
            ExecuteAnotherSQL("Insert into Effective(AgentID, NoofCustActual, NoofCustServed, EffDate) values(" & SafeSQL(dtr("Name").ToString) & "," & dtr("cnt") & ",0," & SafeSQL(Format(dtpDate.Value, "yyyyMMdd HH:mm:ss")) & ")")
        End While
        dtr.Close()
        dtr = ReadRecord("Select Count(AgentID) as cnt, SalesAgent.Name FROM Invoice, SalesAgent where Invoice.AgentID = SalesAgent.Code and InvDT between" & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & " group by Name")
        While dtr.Read = True
            ExecuteAnotherSQL("Update Effective Set NoofCustServed=" & dtr("cnt") & " Where AgentID=" & SafeSQL(dtr("Name").ToString))
        End While
        dtr.Close()
        strSql = "SELECT * from Effective"
        ExecuteReport(strSql, "EffectiveRep")
        btnPrint.Enabled = True
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

    Private Sub Effective_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub

    Private Sub Effective_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ConnectDB()
    End Sub
End Class