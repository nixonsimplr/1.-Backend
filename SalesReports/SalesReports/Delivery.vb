Imports System.Data
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales

Public Class Delivery
    Implements ISalesBase

    Dim sArea As String = " "

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


    Private Sub Delivery_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub


    Private Sub Delivery_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ConnectDB()
        LoadArea()
    End Sub

    Private Sub LoadArea()
        Dim rs As SqlDataReader
        cbDelivery.Items.Clear()
        cbDelivery.Items.Add("ALL")
        rs = ReadRecord("Select distinct ZoneCode from Customer where Active = 1 order by ZoneCode")
        While rs.Read
            cbDelivery.Items.Add(rs("ZoneCode").ToString)
        End While
        rs.Close()
        cbDelivery.SelectedIndex = 0
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        If cbDelivery.Text <> "ALL" Then
            '  sArea = " and OrderHdr.ShipName = " & SafeSQL(cbDelivery.Text)
            sArea = " and Customer.Zonecode = " & SafeSQL(cbDelivery.Text)
        Else
            sArea = ""
        End If

        Dim strSql As String = ""
        strSql = strSql & "Select OrderHdr.OrdNo, OrderHdr.OrdDt, OrderHdr.CustId, OrderHdr.DeliveryDate, OrderHdr.ShipName as DeliveryArea, Customer.CustName , SalesAgent.Name as AgentID, OrdItem.ItemNo, OrdItem.Qty, ITem.ItemName, OrdItem.UOM from OrderHdr, Customer, SalesAgent, Item, OrdItem where OrdItem.OrdNo = OrderHdr.OrdNo and OrdItem.ItemNo = Item.ItemNo and OrderHdr.CustId = Customer.CustNo and SalesAgent.Code = OrderHdr.AgentId and OrderHdr.DeliveryDate between " & SafeSQL(Format(dtpDelivery.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDelivery.Value, "yyyyMMdd 23:59:59")) & sArea & " order by CustNo"
        ExecuteReport(strSql, "DeliveryRep")
    End Sub
End Class