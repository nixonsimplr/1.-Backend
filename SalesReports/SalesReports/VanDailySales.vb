Imports System.Data
Imports System.Data.SqlClient
Imports DataInterface.IbizDO
Imports SalesInterface.MobileSales
Public Class VanDailySales
    Implements ISalesBase
    Dim sAgentID, sLocation, sAgentName As String
    'Dim strTerms As String
    Private objDO As New DataInterface.IbizDO
    Dim ArList As New ArrayList
    Dim sOrderBy As String
    Private Structure ArrItemPrice
        Dim ItemCode As String
        Dim Desc As String
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
        Dim frm As New status
        frm.Show()
        frm.Refresh()
        Dim sAgentPrefix As String
        Dim dSyncDate As Date = CDate("1/1/2000")
        Dim dLastSyncDate As Date = CDate("1/1/2000")
        Dim iCnt As Integer
        Try
            dSyncDate = dgvSync.Item(0, dgvSync.CurrentCell.RowIndex).Value
            sAgentID = dgvSync.Item(1, dgvSync.CurrentCell.RowIndex).Value
            sLocation = dgvSync.Item(2, dgvSync.CurrentCell.RowIndex).Value
            sAgentName = dgvSync.Item(3, dgvSync.CurrentCell.RowIndex).Value
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
        Dim dVInv, dChkIn, dOut, dDel, dEx, dVar, dRtn As Double
        Dim str As String
        ConnectDB()
        Dim strDel = "Delete from VanDailySales where Location=" & SafeSQL(sLocation)
        ExecuteSQL(strDel)
        dtr = ReadRecord("SELECT  PreInvNo as Agent, OrderBy FROM  MDT WHERE AgentID=" & SafeSQL(sAgentID))
        If dtr.Read = True Then
            sAgentPrefix = dtr("Agent").ToString
            sOrderBy = dtr("OrderBy").ToString
        Else
            sAgentPrefix = "0"
        End If
        dtr.Close()
        If sOrderBy = "GR" Then
            sOrderBy = "Item.ProdType, Item.Packsize, Item.Category, Item.Brand, Item.Description"
        Else
            sOrderBy = "Item.ProdType, Item.Category, Item.Brand, Item.Packsize, Item.Description"
        End If
        sOrderBy = "Item.Category, Item.Brand, Item.Description"
        ' str = "Select ItemNo, Description from Item, ItemTrans where Item.ItemNo=ItemTrans.ItemID and LOCATION = " & SafeSQL(sLocation) & " and DateDiff(s, " & objDO.SafeSQL(Format(dLastSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) >= 0 and DateDiff(s, " & objDO.SafeSQL(Format(dSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) < 0 order by " & sOrderBy
        'dtr = ReadRecord("SELECT ItemNo, Description FROM Item where Active=1 and ToPDA=1 Order By " & sOrderBy)
        'MsgBox(str)
        'dtr = ReadRecord(str)
        str = "Select Distinct ItemNo, " & sOrderBy & " from Item, ItemTrans where Item.ItemNo=ItemTrans.ItemID and LOCATION = " & SafeSQL(sLocation) & " and DateDiff(s, " & SafeSQL(Format(dLastSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) >= 0 and DateDiff(s, " & SafeSQL(Format(dSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) < 0 and Qty <> 0 order by " & sOrderBy
        dtr = ReadRecord(str)
        While dtr.Read = True
            Dim aPr1 As ArrItemPrice
            aPr1.ItemCode = dtr("ItemNo").ToString
            aPr1.Desc = dtr("Description").ToString
            ArList.Add(aPr1)
        End While
        dtr.Close()
        Dim i As Integer
        For i = 0 To ArList.Count - 1
            Dim aPr As ArrItemPrice
            aPr = ArList(i)
            ''''Van InVn
            str = "Select Sum(ItemTrans.Qty * Uom.BaseQty) as Out from ItemTrans, UOM where UOM.Uom = ItemTrans.Uom and ItemTrans.ItemId = UOM.ItemNo and ItemTrans.LOCATION = " & SafeSQL(sLocation) & " and  ItemTrans.DOCTYPE = 'VANINVN' and ItemTrans.ItemId=" & SafeSQL(aPr.ItemCode) & " and DateDiff(s, " & objDO.SafeSQL(Format(dLastSyncDate, "yyyyMMdd HH:mm:ss")) & ", ItemTrans.DocDt) >= 0 and DateDiff(s, " & objDO.SafeSQL(Format(dSyncDate, "yyyyMMdd HH:mm:ss")) & ", ItemTrans.DocDt) < 0 "
            rs1 = ReadRecord(str)
            If rs1.Read = True Then
                If rs1(0) Is System.DBNull.Value = True Then
                    dVInv = 0
                Else
                    dVInv = rs1(0)
                End If
            End If
            rs1.Close()


            ''''GIN
            str = "Select Sum(ItemTrans.Qty * Uom.BaseQty) as Out from ItemTrans, UOM where UOM.Uom = ItemTrans.Uom and ItemTrans.ItemId = UOM.ItemNo and ItemTrans.LOCATION = " & SafeSQL(sLocation) & " and ItemTrans.DOCTYPE = 'GIN' and ItemTrans.ItemId=" & SafeSQL(aPr.ItemCode) & " and DateDiff(s, " & objDO.SafeSQL(Format(dLastSyncDate, "yyyyMMdd HH:mm:ss")) & ", ItemTrans.DocDt) >= 0 and DateDiff(s, " & objDO.SafeSQL(Format(dSyncDate, "yyyyMMdd HH:mm:ss")) & ", ItemTrans.DocDt) < 0 "
            rs1 = ReadRecord(str)
            If rs1.Read = True Then
                If rs1(0) Is System.DBNull.Value = True Then
                    dChkIn = 0
                Else
                    dChkIn = rs1(0)
                End If
            End If
            rs1.Close()

            str = "Select Sum( ItemTrans.Qty * Uom.BaseQty) as Gout from ItemTrans, Uom where ItemTrans.ItemId = UOM.ItemNo and UOM.UOM = ItemTrans.UOM and ItemTrans.ItemId=" & SafeSQL(aPr.ItemCode) & " and Location = " & SafeSQL(sLocation) & "  and DateDiff(s, " & SafeSQL(Format(dLastSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) >= 0 and DateDiff(s, " & SafeSQL(Format(dSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) < 0 and Doctype = 'GOUT'"
            rs1 = ReadRecord(str)
            If rs1.Read = True Then
                If rs1(0) Is System.DBNull.Value = False Then
                    dOut = rs1(0)
                Else
                    dOut = 0
                End If
            End If
            rs1.Close()
            'dOut = dVInv + dChkIn

            ''''Delivered
            str = "SELECT SUM(InvItem.Qty * Uom.BaseQty) AS Delivered FROM InvItem INNER JOIN Invoice ON InvItem.InvNo = Invoice.InvNo INNER JOIN UOM ON UOM.ItemNo = InvItem.ItemNo and InvItem.UOM = UOM.Uom WHERE (Invoice.Void = 0 and InvItem.Price > 0 and InvItem.ItemNo=" & SafeSQL(aPr.ItemCode) & " and Invoice.AgentID = " & SafeSQL(sAgentID) & " and DateDiff(s, " & objDO.SafeSQL(Format(dLastSyncDate, "yyyyMMdd HH:mm:ss")) & ", InvDt) >= 0 and DateDiff(s, " & objDO.SafeSQL(Format(dSyncDate, "yyyyMMdd HH:mm:ss")) & ", InvDt) < 0 )"
            rs1 = ReadRecord(str)
            If rs1.Read = True Then
                If rs1(0) Is System.DBNull.Value = True Then
                    dDel = 0
                Else
                    dDel = rs1(0)
                End If
            End If
            rs1.Close()
            ''''FOC
            str = "SELECT SUM(InvItem.Qty * Uom.BaseQty) AS Foc FROM InvItem INNER JOIN Invoice ON InvItem.InvNo = Invoice.InvNo INNER JOIN UOM ON UOM.ItemNo = InvItem.ItemNo and InvItem.UOM = UOM.Uom WHERE (Invoice.Void = 0 and InvItem.Price = 0  and InvItem.Description like 'FOC:%' and InvItem.ItemNo=" & SafeSQL(aPr.ItemCode) & " and Invoice.AgentID = " & SafeSQL(sAgentID) & " and DateDiff(s, " & objDO.SafeSQL(Format(dLastSyncDate, "yyyyMMdd HH:mm:ss")) & ", InvDt) >= 0 and DateDiff(s, " & objDO.SafeSQL(Format(dSyncDate, "yyyyMMdd HH:mm:ss")) & ", InvDt) < 0 )"
            rs1 = ReadRecord(str)
            If rs1.Read = True Then
                If rs1(0) Is System.DBNull.Value = True Then
                    fFoc = 0
                Else
                    fFoc = rs1(0)
                End If
            End If
            rs1.Close()
            ''''Exchange
            str = "SELECT SUM(InvItem.Qty * Uom.BaseQty) AS Ex FROM InvItem INNER JOIN Invoice ON InvItem.InvNo = Invoice.InvNo INNER JOIN UOM ON UOM.ItemNo = InvItem.ItemNo and InvItem.UOM = UOM.Uom WHERE (Invoice.Void = 0 and InvItem.Price = 0  and InvItem.Description like 'EX:%' and InvItem.ItemNo=" & SafeSQL(aPr.ItemCode) & " and Invoice.AgentID = " & SafeSQL(sAgentID) & " and DateDiff(s, " & objDO.SafeSQL(Format(dLastSyncDate, "yyyyMMdd HH:mm:ss")) & ", InvDt) >= 0 and DateDiff(s, " & objDO.SafeSQL(Format(dSyncDate, "yyyyMMdd HH:mm:ss")) & ", InvDt) < 0 )"
            rs1 = ReadRecord(str)
            If rs1.Read = True Then
                If rs1(0) Is System.DBNull.Value = True Then
                    dEx = 0
                Else
                    dEx = rs1(0)
                End If
            End If
            rs1.Close()
         
            ''''Goods Return
            str = "SELECT SUM(GoodsReturnItem.Quantity * Uom.BaseQty) AS ReturnItem FROM GoodsReturnItem  INNER JOIN GoodsReturn ON GoodsReturnItem.ReturnNo = GoodsReturn.ReturnNo INNER JOIN UOM ON UOM.ItemNo = GoodsReturnItem.ItemNo and GoodsReturnItem.UOM = UOM.Uom where GoodsReturn.SalesPersonCode = " & SafeSQL(sAgentID) & " and GoodsReturnItem.ItemNo=" & SafeSQL(aPr.ItemCode) & " and DateDiff(s, " & objDO.SafeSQL(Format(dLastSyncDate, "yyyyMMdd HH:mm:ss")) & ", ReturnDate) >= 0 and DateDiff(s, " & objDO.SafeSQL(Format(dSyncDate, "yyyyMMdd HH:mm:ss")) & ", ReturnDate) < 0 "
            rs1 = ReadRecord(str)
            If rs1.Read = True Then
                If rs1(0) Is System.DBNull.Value = True Then
                    drtn = 0
                Else
                    drtn = rs1(0)
                End If
            End If
            rs1.Close()
            ''''Goods Variance
            str = "Select Sum(ItemTrans.Qty * Uom.BaseQty) as ReturnItem from ItemTrans, UOM where UOM.Uom = ItemTrans.Uom and ItemTrans.ItemId = UOM.ItemNo and ItemTrans.LOCATION = " & SafeSQL(sLocation) & " and  ItemTrans.DOCTYPE = 'GVAR' and ItemTrans.ItemId=" & SafeSQL(aPr.ItemCode) & " and DateDiff(s, " & objDO.SafeSQL(Format(dLastSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) >= 0 and DateDiff(s, " & objDO.SafeSQL(Format(dSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) < 0 "
            rs1 = ReadRecord(str)
            If rs1.Read = True Then
                If rs1(0) Is System.DBNull.Value = True Then
                    dVar = 0
                Else
                    dVar = rs1(0)
                End If
            End If
            rs1.Close()
            Dim sRemarks As String = ""
            str = "Select Remarks from ItemTrans where LOCATION = " & SafeSQL(sLocation) & " and  DOCTYPE = 'GVAR' and ItemTrans.ItemId=" & SafeSQL(aPr.ItemCode) & " and DateDiff(s, " & objDO.SafeSQL(Format(dLastSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) >= 0 and DateDiff(s, " & objDO.SafeSQL(Format(dSyncDate, "yyyyMMdd HH:mm:ss")) & ", DocDt) < 0 "
            rs1 = ReadRecord(str)
            If rs1.Read = True Then
                sRemarks = rs1("Remarks").ToString
            End If
            rs1.Close()

            If dVInv = 0 And dChkIn = 0 And dOut = 0 And dDel = 0 And dEx = 0 And dVar = 0 And dRtn = 0 Then Continue For
            Dim insSql As String = "Insert into VanDailySales(DocNo, ItemNo, ItemName, Uom, Out, Delivered, Exchange, Foc, AgentID, Date, GVar, ReturnItem, Location, VInvn, GIN, Remarks) values (''," & SafeSQL(aPr.ItemCode) & "," & SafeSQL(aPr.Desc) & "," & SafeSQL("") & "," & dOut & "," & dDel & "," & dEx & "," & fFoc & "," & SafeSQL(sAgentPrefix & " / " & sAgentName) & ",'" & Format(dSyncDate, "yyyyMMdd HH:mm:ss") & "'," & dVar & "," & dRtn & "," & SafeSQL(sLocation) & "," & dVInv & "," & dChkIn & "," & SafeSQL(sRemarks) & ")"
            ExecuteSQL(insSql)
            'dtr.Close()
            frm.Refresh()
        Next
        frm.Close()
        'Dim strSql As String = "Select Distinct VanDailySales.*, Item.ProdType, Item.Category, Item.Brand, Item.Packsize, Item.Description from VanDailySales, Item where VanDailySales.ItemNo=Item.ItemNo and Location=" & SafeSQL(sLocation) & " order by " & sOrderBy
        Dim strsql As String = "Select Distinct VanDailySales.*,Item.Category, Item.Brand, Item.Description from VanDailySales, Item where VanDailySales.ItemNo=Item.ItemNo and Location=" & SafeSQL(sLocation) & " order by " & sOrderBy
        ExecuteReport(strSql, "VanDailySalesRep")

        'PrintAllReport(strSql, "VanDailySalesRep")
        'PrintAllReport(strSql, "VanDailySalesRep")

        btnPrint.Enabled = True
    End Sub

    Private Function GetPrinterName() As String
        Dim ds As New DataSet
        Dim dataDirectory As String
        dataDirectory = Windows.Forms.Application.StartupPath
        ds.ReadXml(dataDirectory & "\ibiz.xml")
        Dim table As DataTable
        For Each table In ds.Tables
            Dim row As DataRow
            If table.TableName = "PrinterName" Then
                For Each row In table.Rows
                    Return row("Value").ToString()
                Next row
            End If
        Next table
        Return ""
    End Function

    Private Sub PrintAllReport(ByVal strSql As String, ByVal RptName As String)
        '  sPrinterName = GetPrinterName()
        Dim DA As New SqlDataAdapter(strSql, My.Settings.ConnectionString)
        Dim DS As New DataSet
        DA.Fill(DS)
        Dim ConInfo As New CrystalDecisions.Shared.TableLogOnInfo
        ConInfo.ConnectionInfo.IntegratedSecurity = True
        'ConInfo.ConnectionInfo.UserID = objDataBase.UserName
        'ConInfo.ConnectionInfo.Password = objDataBase.Password
        ConInfo.ConnectionInfo.ServerName = ".\SQLEXPRESS"
        ConInfo.ConnectionInfo.DatabaseName = "Sales"

        Dim strReportPath As String = Application.StartupPath & "\" & RptName & ".rpt"
        If Not IO.File.Exists(strReportPath) Then
            Throw (New Exception("Unable to locate report file:" & vbCrLf & strReportPath))
        End If
        Dim rptDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        rptDocument.Load(strReportPath)
        rptDocument.SetDataSource(DS.Tables(0))
        'rptDocument.PrintOptions.PrinterName = "Epson LQ-300 ESC/P 2"
        Dim frm As New ViewReport
        frm.crvReport.ShowRefreshButton = False
        frm.crvReport.ShowCloseButton = False
        frm.crvReport.ShowGroupTreeButton = False
        frm.crvReport.ReportSource = rptDocument
        'rptDocument.PrintOptions.PrinterName
        rptDocument.PrintOptions.PrinterName = GetPrinterName() '"HP LaserJet P3005 PCL 6"
        rptDocument.PrintToPrinter(1, False, 0, 0)
        'frm.crvReport.PrintOptions.PrinterName = "\\Epson LQ-300 ESC/P 2"
        'frm.Show()
        'frm.crvReport.PrintReport()
        'frm.crvReport.Dispose()
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