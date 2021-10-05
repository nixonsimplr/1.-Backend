Imports System.Data
Imports System.Data.SqlClient
Imports DataInterface.IbizDO
Imports SalesInterface.MobileSales
Public Class AssetTrans
    Implements ISalesBase

    Dim sAgentID, sLocation As String
    'Dim strTerms As String
    Private objDO As New DataInterface.IbizDO
    Dim ArList As New ArrayList
    Private Structure ArrItemPrice
        Dim sAssetID As String
        Dim Desc As String
    End Structure

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        btnPrint.Enabled = False
        Dim dSyncDate As Date = CDate("1/1/2000")
        Dim dLastSyncDate As Date = CDate("1/1/2000")
        Dim iCnt As Integer
        Try
            dSyncDate = dgvSync.Item(0, dgvSync.CurrentCell.RowIndex).Value
            sAgentID = dgvSync.Item(1, dgvSync.CurrentCell.RowIndex).Value
            sLocation = dgvSync.Item(2, dgvSync.CurrentCell.RowIndex).Value
            For iCnt = dgvSync.CurrentCell.RowIndex + 1 To dgvSync.Rows.Count - 1
                If sAgentID = dgvSync.Item(1, iCnt).Value Then
                    dLastSyncDate = dgvSync.Item(0, iCnt).Value
                    Exit For
                End If
            Next
        Catch ex As Exception
            MsgBox("Please select a date and try again.")
            btnPrint.Enabled = True
            Exit Sub
        End Try

        If dSyncDate = CDate("1/1/2000") Or dLastSyncDate = CDate("1/1/2000") Then
            MsgBox("Please select a valid date and try again.")
            btnPrint.Enabled = True
            Exit Sub
        End If
        Dim dtr As SqlDataReader
        Dim rs1 As SqlDataReader
        Dim fFoc As Double
        Dim dOut, dDel, dEx, dRt As Double
        Dim str As String
        ConnectDB()
        Dim strDel = "Delete from AssetSales"
        ExecuteSQL(strDel)
        dtr = ReadRecord("SELECT AssetID, AssetDesc FROM Asset")
        While dtr.Read = True
            Dim aPr1 As ArrItemPrice
            aPr1.sAssetID = dtr("AssetID").ToString
            aPr1.Desc = dtr("AssetDesc").ToString
            ArList.Add(aPr1)
        End While
        dtr.Close()
        Dim i As Integer
        For i = 0 To ArList.Count - 1
            Dim aPr As ArrItemPrice
            aPr = ArList(i)
            'GIN
            str = "Select Sum(InQty) as Out from AssetTrans where LOCATION = " & SafeSQL(sLocation) & " and  (DOCTYPE = 'GIN' or DOCTYPE = 'VANINVN') and AssetTrans.AssetId=" & SafeSQL(aPr.sAssetID) & " and DateDiff(s, " & objDO.SafeSQL(Format(dLastSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) >= 0 and DateDiff(s, " & objDO.SafeSQL(Format(dSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) < 0 "
            rs1 = ReadRecord(str)
            If rs1.Read = True Then
                If rs1(0) Is System.DBNull.Value = True Then
                    dOut = 0
                Else
                    dOut = rs1(0)
                End If
            End If
            rs1.Close()
            str = "SELECT SUM(InvItem.Qty) AS Delivered FROM InvItem INNER JOIN Invoice ON InvItem.InvNo = Invoice.InvNo WHERE (Invoice.Void = 0 and InvItem.Price > 0 and InvItem.ItemNo=" & SafeSQL(aPr.sAssetID) & " and Invoice.AgentID = " & SafeSQL(sAgentID) & " and DateDiff(s, " & objDO.SafeSQL(Format(dLastSyncDate, "yyyyMMdd HH:mm:ss")) & ", InvDt) >= 0 and DateDiff(s, " & objDO.SafeSQL(Format(dSyncDate, "yyyyMMdd HH:mm:ss")) & ", InvDt) < 0 )"
            rs1 = ReadRecord(str)
            If rs1.Read = True Then
                If rs1(0) Is System.DBNull.Value = True Then
                    dDel = 0
                Else
                    dDel = rs1(0)
                End If
            End If
            rs1.Close()

            str = "SELECT SUM(InvItem.Qty) AS Foc FROM InvItem INNER JOIN Invoice ON InvItem.InvNo = Invoice.InvNo WHERE (Invoice.Void = 0 and InvItem.Price = 0  and InvItem.ItemNo=" & SafeSQL(aPr.sAssetID) & " and Invoice.AgentID = " & SafeSQL(sAgentID) & " and DateDiff(s, " & objDO.SafeSQL(Format(dLastSyncDate, "yyyyMMdd HH:mm:ss")) & ", InvDt) >= 0 and DateDiff(s, " & objDO.SafeSQL(Format(dSyncDate, "yyyyMMdd HH:mm:ss")) & ", InvDt) < 0 )"
            rs1 = ReadRecord(str)
            If rs1.Read = True Then
                If rs1(0) Is System.DBNull.Value = True Then
                    fFoc = 0
                Else
                    fFoc = rs1(0)
                End If
            End If
            rs1.Close()

            str = "SELECT SUM(GoodsExchangeItem.Quantity) AS Exchange FROM GoodsExchangeItem  INNER JOIN GoodsExchange ON GoodsExchangeItem.ExchangeNo = GoodsExchange.ExchangeNo where GoodsExchange.SalesPersonCode = " & SafeSQL(sAgentID) & " and GoodsExchangeItem.ItemNo=" & SafeSQL(aPr.sAssetID) & " and DateDiff(s, " & objDO.SafeSQL(Format(dLastSyncDate, "yyyyMMdd HH:mm:ss")) & ", ExchangeDate) >= 0 and DateDiff(s, " & objDO.SafeSQL(Format(dSyncDate, "yyyyMMdd HH:mm:ss")) & ", ExchangeDate) < 0 "
            rs1 = ReadRecord(str)
            If rs1.Read = True Then
                If rs1(0) Is System.DBNull.Value = True Then
                    dEx = 0
                Else
                    dEx = rs1(0)
                End If
            End If
            rs1.Close()
            str = "Select Sum(Qty) as ReturnItem from ItemTrans where LOCATION = " & SafeSQL(sLocation) & " and  DOCTYPE = 'GVAR' and ItemTrans.ItemId=" & SafeSQL(aPr.sAssetID) & " and DateDiff(s, " & objDO.SafeSQL(Format(dLastSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) >= 0 and DateDiff(s, " & objDO.SafeSQL(Format(dSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) < 0 "
            rs1 = ReadRecord(str)
            If rs1.Read = True Then
                If rs1(0) Is System.DBNull.Value = True Then
                    dRt = 0
                Else
                    dRt = rs1(0)
                End If
            End If
            rs1.Close()
            If dOut = 0 And dDel = 0 And dEx = 0 And dRt = 0 Then Continue For
            Dim insSql As String = "Insert into VanDailySales(DocNo, ItemNo, ItemName, Uom, Out, Delivered, Exchange, Foc, AgentID, Date, ReturnItem) values (''," & SafeSQL(aPr.sAssetID) & "," & SafeSQL(aPr.Desc) & "," & SafeSQL("") & "," & dOut & "," & dDel & "," & dEx & "," & fFoc & "," & SafeSQL(sAgentID) & ",'" & Format(dSyncDate, "yyyyMMdd HH:mm:ss") & "'," & dRt & ")"
            ExecuteSQL(insSql)

            'dtr.Close()
        Next
        Dim strSql As String = "Select Distinct VanDailySales.*, Item.DisplayNo  from VanDailySales, Item where VanDailySales.ItemNo=Item.ItemNo order by DisplayNo"
        ExecuteReport(strSql, "VanDailySalesRep")
        btnPrint.Enabled = True
    End Sub

    Private Sub AssetTrans_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub

    Private Sub AssetTrans_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.SyncHistoryTableAdapter1.Fill(Me.SalesDataSet1.SyncHistory)
        ConnectDB()
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