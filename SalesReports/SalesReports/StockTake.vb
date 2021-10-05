Imports System.Data.SqlClient
Imports SalesInterface.MobileSales

Public Class StockTake
    Implements ISalesBase
    Private aAgent As New ArrayList()
    Dim strLocation As String = " "

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

    Public Sub loadCombo()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Code, Name from Location order by [Name]")
        cmbLocation.DataSource = Nothing
        aAgent.Clear()
        aAgent.Add(New ComboValues("ALL", "ALL"))
        While dtr.Read()
            aAgent.Add(New ComboValues(dtr("Code").ToString, dtr("Name").ToString))
        End While
        dtr.Close()
        cmbLocation.DataSource = aAgent
        cmbLocation.DisplayMember = "Desc"
        cmbLocation.ValueMember = "Code"
        cmbLocation.SelectedIndex = 0
    End Sub

    Private Sub StockTake_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub

    Private Sub StockTake_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ConnectDB()
        loadCombo()
    End Sub

    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        btnPrint.Enabled = False
        If cmbLocation.Text = "ALL" Then
            strLocation = " "
        Else
            strLocation = "and Location.Code=" & SafeSQL(cmbLocation.SelectedValue) & " "
        End If

        ExecuteSQL("Delete from StockTakeRep")

        Dim strSql As String
        strSql = "SELECT StockTakeDet.StockTakeNo, StockTakeHdr.StockTakeDate, " & SafeSQL(Format(dtpFromDate.Value, "dd-MMM-yyyy 00:00:00")) & " as FromDate, " & SafeSQL(Format(dtpToDate.Value, "dd-MMM-yyyy 23:59:59")) & " as ToDate, StockTakeHdr.AgentId, SalesAgent.Name AS AgentName, StockTakeHdr.LocationCode, Location.Name as LocationName, StockTakeDet.ItemId, Item.ItemName, StockTakeDet.Qty, StockTakeDet.DiffQty, StockTakeDet.Uom, StockTakeDet.Remarks FROM StockTakeDet INNER JOIN StockTakeHdr ON StockTakeDet.StockTakeNo = StockTakeHdr.StockTakeNo INNER JOIN Location ON StockTakeHdr.LocationCode = Location.Code INNER JOIN SalesAgent ON StockTakeHdr.AgentId = SalesAgent.Code INNER JOIN Item ON StockTakeDet.ItemId = Item.ItemNo WHERE StockTakeHdr.StockTakeDate Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & " " & strLocation & " order by Item.ItemNo"

        ExecuteSQL("Insert StockTakeRep(StockTakeNo, StockTakeDate, FromDate, ToDate, AgentID, AgentName, LocationCode, LocationName, ItemNo, ItemName, Qty, Diff, UOM, Reason) " & strSql)

        ExecuteReport("Select * from StockTakeRep order by ItemNo", "StockTakeRep")
        btnPrint.Enabled = True
    End Sub
End Class