Imports System.Data
Imports System.Data.SqlClient
Imports DataInterface.IbizDO
Imports SalesInterface.MobileSales
Public Class VanInventory
    Implements ISalesBase
    Dim sAgent, sAgentID, sLocation As String
    Private objDO As New DataInterface.IbizDO
    Dim ArList As New ArrayList
    Private Structure ArrItemPrice
        Dim ItemCode As String
        Dim Desc As String
        Dim UOM As String
    End Structure
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub VanDailySales_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub

    Private Sub VanDailySales_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.SyncHistoryTableAdapter1.Fill(Me.SalesDataSet1.SyncHistory)
        ConnectDB()
    End Sub

    Private Sub btnPrint_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        btnPrint.Enabled = False
        Dim dSyncDate As Date = CDate("1/1/2000")
        Dim dLastSyncDate As Date = CDate("1/1/2000")
        Dim iCnt As Integer
        Try
            dSyncDate = dgvSync.Item(0, dgvSync.CurrentCell.RowIndex).Value
            sAgent = dgvSync.Item(1, dgvSync.CurrentCell.RowIndex).Value
            sLocation = dgvSync.Item(2, dgvSync.CurrentCell.RowIndex).Value
            For iCnt = dgvSync.CurrentCell.RowIndex + 1 To dgvSync.Rows.Count - 1
                If sAgent = dgvSync.Item(1, iCnt).Value Then
                    dLastSyncDate = dgvSync.Item(0, iCnt).Value
                    Exit For
                End If
            Next
        Catch ex As Exception
            MsgBox("Please select a date and try again.")
            btnPrint.Enabled = True
            Exit Sub
        End Try
        Dim dtr As SqlDataReader
        dtr = ReadRecord("SELECT  PreInvNo as Agent FROM  MDT WHERE AgentID=" & SafeSQL(sAgent))
        If dtr.Read = True Then
            sAgentID = dtr("Agent").ToString
        Else
            sAgentID = "0"
        End If
        dtr.Close()

        dtr = ReadRecord("SELECT  Name as Agent FROM  SalesAgent where Code=" & SafeSQL(sAgent))
        If dtr.Read = True Then
            sAgentID = sAgentID & " / " & dtr("Agent").ToString
        Else
            sAgentID = "0"
        End If
        dtr.Close()

        If dSyncDate = CDate("1/1/2000") Or dLastSyncDate = CDate("1/1/2000") Then
            MsgBox("Please select a valid date and try again.")
            btnPrint.Enabled = True
            Exit Sub
        End If
        'Dim dtr As SqlDataReader
        Dim rs1 As SqlDataReader
        Dim fFoc As Double
        Dim dOut, dDel, dEx, dRt, dRtn As Double
        Dim str As String
        ConnectDB()
        Dim strDel = "Delete from VanDailySales"
        ExecuteSQL(strDel)
        dtr = ReadRecord("SELECT Item.ItemNo, Description, ItemTrans.UOM FROM Item, ItemTrans where Item.ItemNo=ItemTrans.ItemID and DateDiff(s, " & objDO.SafeSQL(Format(dLastSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) >= 0 and DateDiff(s, " & objDO.SafeSQL(Format(dSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) < 0")
        While dtr.Read = True
            Dim aPr1 As ArrItemPrice
            aPr1.ItemCode = dtr("ItemNo").ToString
            aPr1.Desc = dtr("Description").ToString
            aPr1.UOM = dtr("UOM").ToString
            ArList.Add(aPr1)
        End While
        dtr.Close()
        Dim i As Integer
        For i = 0 To ArList.Count - 1
            Dim aPr As ArrItemPrice
            aPr = ArList(i)
            ''''Van Inventory
            str = "Select Sum(Qty) as Out from ItemTrans where LOCATION = " & SafeSQL(sLocation) & " and  DOCTYPE = 'VANINVN' and ItemTrans.ItemId=" & SafeSQL(aPr.ItemCode) & " and DateDiff(s, " & objDO.SafeSQL(Format(dLastSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) >= 0 and DateDiff(s, " & objDO.SafeSQL(Format(dSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) < 0 "
            rs1 = ReadRecord(str)
            If rs1.Read = True Then
                If rs1(0) Is System.DBNull.Value = True Then
                    dOut = 0
                Else
                    dOut = rs1(0)
                End If
            End If
            rs1.Close()
            If dOut = 0 Then Continue For
            Dim insSql As String = "Insert into VanDailySales(DocNo, ItemNo, ItemName, Uom, Out, Delivered, Exchange, Foc, AgentID, Date, GVar, ReturnItem) values (''," & SafeSQL(aPr.ItemCode) & "," & SafeSQL(aPr.Desc) & "," & SafeSQL(aPr.UOM) & "," & dOut & "," & dDel & "," & dEx & "," & fFoc & "," & SafeSQL(sAgentID) & ",'" & Format(dSyncDate, "yyyyMMdd HH:mm:ss") & "'," & dRt & "," & dRtn & ")"
            ExecuteSQL(insSql)
            'dtr.Close()
        Next
        Dim strSql As String = "Select Distinct VanDailySales.*, Item.ProdType, Item.Category, Item.Brand, Item.Packsize, Item.Description  from VanDailySales, Item where VanDailySales.ItemNo=Item.ItemNo order by Item.ProdType, Item.Category, Item.Brand, Item.Packsize, Item.Description"
        ExecuteReport(strSql, "VanInventoryRep")
        btnPrint.Enabled = True
    End Sub

    Private Sub dgvSync_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub SyncHistoryBindingSource_CurrentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SyncHistoryBindingSource.CurrentChanged

    End Sub

    Private Sub dgvSync_CellContentClick_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

    End Sub

    Private Sub dgvSync_CellContentClick_2(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSync.CellContentClick

    End Sub

    Public Function GetListViewForm() As String Implements SalesInterface.MobileSales.ISalesBase.GetListViewForm
        Return ""
    End Function

    Public Sub ListViewClick() Implements SalesInterface.MobileSales.ISalesBase.ListViewClick

    End Sub

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

    Public Event LoadDataForm(ByVal Location As String, ByVal LoadType As String, ByVal ParentLoadType As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.LoadDataForm

    Public Event LoadForm(ByVal Location As String, ByVal LoadType As String) Implements SalesInterface.MobileSales.ISalesBase.LoadForm
End Class