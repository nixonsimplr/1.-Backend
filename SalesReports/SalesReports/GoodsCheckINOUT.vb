Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Public Class GoodsCheckINOUT
    Implements ISalesBase
    Dim sAgentID, sLocation As String
    Dim strAgent As String = " "
    Private objDO As New DataInterface.IbizDO
    Dim ArList As New ArrayList
    Private Structure ArrItemPrice
        Dim ItemCode As String
        Dim Desc As String
        Dim UOM As String
    End Structure
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub GoodsCheckINOUT_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'SalesDataSet2.SyncHistory' table. You can move, or remove it, as needed.
        Me.SyncHistoryTableAdapter.Fill(Me.SalesDataSet2.SyncHistory)
        ConnectDB()
        ' loadAgent()
        'dtpDate.Value = Date.Now
    End Sub
    'Public Sub loadAgent()
    '    cmbLoc.Items.Add("Select")
    '    Dim dtr As SqlDataReader
    '    dtr = ReadRecord("Select distinct(Location) from ItemTrans")
    '    Do While dtr.Read = True
    '        cmbLoc.Items.Add(dtr(0))
    '    Loop
    '    dtr.Close()
    '    If cmbLoc.Items.Count > 0 Then
    '        cmbLoc.SelectedIndex = 0
    '    End If
    'End Sub
    'Private Sub cmbAgent_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If cmbLoc.Text = "Select" Then
    '        strAgent = " "
    '    Else
    '        strAgent = " and Agentid=" & "'" & cmbLoc.Text & "'"
    '    End If
    'End Sub

    Private Sub btnPrint_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
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
        ' Dim dDocdt As Date
        '  Dim sDocNo, sLoc As String
        Dim dIss, dDeli, dTop, dEx, dGout, Actual, dRet As Double
        Dim str As String
        Dim strDel = "Delete from GoodsCheck"
        ExecuteSQL(strDel)
        dtr = ReadRecord("SELECT ItemNo, Description, BaseUOM FROM Item where ToPDA=1 order by DisplayNo")
        While dtr.Read = True
            Dim aPr1 As ArrItemPrice
            aPr1.ItemCode = dtr("ItemNo").ToString
            aPr1.Desc = dtr("Description").ToString
            aPr1.UOM = dtr("BaseUOM").ToString
            ArList.Add(aPr1)
        End While
        dtr.Close()
        Dim i As Integer
        For i = 0 To ArList.Count - 1
            Dim aPr As ArrItemPrice
            aPr = ArList(i)

            'Van Inventory

            'If aPr.ItemCode = "1110067" Then
            '    MsgBox("1")
            'End If
            str = "Select Sum( Qty) as Issue from ItemTrans, Uom where ItemTrans.ItemId = UOM.ItemNo and UOM.UOM = ItemTrans.UOM and ItemTrans.ItemId=" & SafeSQL(aPr.ItemCode) & " and Location = " & SafeSQL(sLocation) & " and DateDiff(s, " & SafeSQL(Format(dLastSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) >= 0 and DateDiff(s, " & SafeSQL(Format(dSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) < 0 and Doctype = 'VANINVN'"
            rs1 = ReadRecord(str)
            If rs1.Read = True Then
                If rs1(0) Is System.DBNull.Value = False Then
                    dIss = rs1(0)
                Else
                    dIss = 0
                End If
            End If
            rs1.Close()
            'Top- up
            str = "Select Sum( Qty) as Topup from ItemTrans, Uom where ItemTrans.ItemId = UOM.ItemNo and UOM.UOM = ItemTrans.UOM and ItemTrans.ItemId=" & SafeSQL(aPr.ItemCode) & " and Location = " & SafeSQL(sLocation) & " and DateDiff(s, " & SafeSQL(Format(dLastSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) >= 0 and DateDiff(s, " & SafeSQL(Format(dSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) < 0 and Doctype = 'GIN'"
            rs1 = ReadRecord(str)
            If rs1.Read = True Then
                If rs1(0) Is System.DBNull.Value = False Then
                    dTop = rs1(0)
                Else
                    dTop = 0
                End If
            End If
            rs1.Close()

            'Delivered

            str = "Select Sum( Qty) as DELI from ItemTrans, Uom where ItemTrans.ItemId = UOM.ItemNo and UOM.UOM = ItemTrans.UOM and ItemTrans.ItemId=" & SafeSQL(aPr.ItemCode) & " and Location = " & SafeSQL(sLocation) & "  and DateDiff(s, " & SafeSQL(Format(dLastSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) >= 0 and DateDiff(s, " & SafeSQL(Format(dSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) < 0 and Doctype = 'DELI' "
            rs1 = ReadRecord(str)
            If rs1.Read = True Then
                If rs1(0) Is System.DBNull.Value = False Then
                    dDeli = rs1(0)
                Else
                    dDeli = 0
                End If
            End If
            rs1.Close()

            'str = "Select Sum( Qty) as Ex from ItemTrans, Uom where ItemTrans.ItemId = UOM.ItemNo and UOM.UOM = ItemTrans.UOM and ItemTrans.ItemId=" & SafeSQL(dtr("ItemNo")) & " and Location = " & SafeSQL(sLocation) & "  and DateDiff(s, " & SafeSQL(Format(dLastSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) >= 0 and DateDiff(s, " & SafeSQL(Format(dSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) < 0 and Doctype = 'EX'"
            'rs1 = ReadRecordAnother(str)
            'If rs1.Read = True Then
            '    If rs1(0) Is System.DBNull.Value = False Then
            '        dEx = rs1(0)
            '    Else
            '        dEx = 0
            '    End If
            'End If
            'rs1.Close()

            ' Gout

            str = "Select Sum( Qty) as Gout from ItemTrans, Uom where ItemTrans.ItemId = UOM.ItemNo and UOM.UOM = ItemTrans.UOM and ItemTrans.ItemId=" & SafeSQL(aPr.ItemCode) & " and Location = " & SafeSQL(sLocation) & "  and DateDiff(s, " & SafeSQL(Format(dLastSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) >= 0 and DateDiff(s, " & SafeSQL(Format(dSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) < 0 and Doctype = 'GOUT'"
            rs1 = ReadRecord(str)
            If rs1.Read = True Then
                If rs1(0) Is System.DBNull.Value = False Then
                    dGout = rs1(0)
                Else
                    dGout = 0
                End If
            End If
            rs1.Close()

            'Return

            str = "Select Sum( Qty) as Gout from ItemTrans, Uom where ItemTrans.ItemId = UOM.ItemNo and UOM.UOM = ItemTrans.UOM and ItemTrans.ItemId=" & SafeSQL(aPr.ItemCode) & " and Location = " & SafeSQL(sLocation) & "  and DateDiff(s, " & SafeSQL(Format(dLastSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) >= 0 and DateDiff(s, " & SafeSQL(Format(dSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) < 0 and Doctype = 'RET'"
            rs1 = ReadRecord(str)
            If rs1.Read = True Then
                If rs1(0) Is System.DBNull.Value = False Then
                    dRet = rs1(0)
                Else
                    dRet = 0
                End If
            End If
            rs1.Close()

            'Item Request
            'str = "Select Sum(Qty) as Req from ItemRequisition, Uom where ItemRequisition.ItemNo = UOM.ItemNo and UOM.UOM = ItemRequisition.UOM and ItemRequisition.ItemNo=" & SafeSQL(aPr.ItemCode) & " and Location = " & SafeSQL(sLocation) & "  and DateDiff(s, " & SafeSQL(Format(dLastSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) >= 0 and DateDiff(s, " & SafeSQL(Format(dSyncDate, "yyyyMMdd HH:mm:ss")) & ", ReqDt) < 0"
            'rs1 = ReadRecord(str)
            'If rs1.Read = True Then
            '    If rs1(0) Is System.DBNull.Value = False Then
            '        dReq = rs1(0)
            '    Else
            '        dReq = 0
            '    End If
            'End If
            'rs1.Close()
            dEx = 0
            'dRet = 0
            Actual = dIss + dTop - dDeli - dEx - dGout - dRet
            'Uom,  "," & SafeSQL(dtr("Uom")) &
            If dGout = 0 And dTop = 0 Then Continue For
            Dim insSql As String = "Insert into GoodsCheck(DocNo,  ItemNo,ItemName, Uom, Issue, Topup, ActualReturn, GOUT, AgentID, Date) values (''," & SafeSQL(aPr.ItemCode) & "," & SafeSQL(aPr.Desc) & "," & SafeSQL(aPr.UOM) & "," & dIss & "," & dTop & "," & Actual & "," & dGout & "," & SafeSQL(sLocation) & ", '" & Format(dSyncDate, "yyyyMMdd HH:mm:ss") & "')"
            ExecuteSQL(insSql)
            'DisconnectAnotherDB()
        Next
        dtr.Close()
        Dim strSql As String = "Select Distinct * from GoodsCheck"
        ExecuteReport(strSql, "GoodsCheckINandOUT")
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
End Class