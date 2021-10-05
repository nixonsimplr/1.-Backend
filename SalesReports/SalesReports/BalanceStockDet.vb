Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Odbc
Imports SalesInterface.MobileSales
Imports System.IO

Public Class BalanceStockDet
    Implements ISalesBase

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


    Private Sub BalanceStockDet_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub

    Private Sub BalanceStockDet_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ConnectDB()
        LoadCombo()
    End Sub

    Private Sub LoadCombo()

        Dim rs As SqlDataReader
        cbCategory.Items.Clear()
        cbCategory.Items.Add("ALL")
        rs = ReadRecord("Select distinct Code from Category order by Code")
        While rs.Read
            cbCategory.Items.Add(rs("Code").ToString)
        End While
        rs.Close()
        cbCategory.SelectedIndex = 0

        'cbSupplier.Items.Clear()
        'cbSupplier.Items.Add("ALL")
        'rs = ReadRecord("Select distinct left(itemno,3) SupplierCode from Item order by left(itemno,3)")
        'While rs.Read
        '    cbSupplier.Items.Add(rs("SupplierCode").ToString)
        'End While
        'rs.Close()
        'cbSupplier.SelectedIndex = 0


        'Dim dtr As SqlDataReader
        'dtr = ReadRecord("Select distinct CompanyName from Customer where active = 1 order by CompanyName")
        'cbCompany.Items.Clear()
        'While dtr.Read()
        '    cbCompany.Items.Add(dtr("CompanyName").ToString)
        'End While
        'dtr.Close()
        'If cbCompany.Items.Count >= 1 Then cbCompany.SelectedIndex = 0
    End Sub


    Private Sub LoadItem()
        Dim sItem As String = " "
        If cbCategory.Text <> "ALL" Then
            sItem = " AND Category = " & SafeSQL(cbCategory.Text)
        End If
        'If cbSupplier.Text <> "ALL" Then
        '    sItem = sItem & " AND left(ItemNo,3) = " & SafeSQL(cbSupplier.Text)
        'End If
        Dim rs As SqlDataReader
        '" and companyName = " & SafeSQL(cbCompany.Text) &
        rs = ReadRecord("Select distinct ItemNo,Description+' '+ShortDesc as Description, BaseUOM, UnitPrice, ToPda, Barcode, CompanyName from Item where Active = 1 " & sItem & " order by Description+' '+ShortDesc")

        dgvOrdItem.Rows.Clear()
        While rs.Read

            Dim row As String() = New String() _
                        {True, rs("ItemNo").ToString, rs("Description").ToString, rs("BaseUOM").ToString, rs("UnitPrice").ToString, rs("Barcode").ToString, rs("CompanyName").ToString}
            dgvOrdItem.Rows.Add(row)

        End While
        rs.Close()
    End Sub

    Private Sub btnSelect_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim i As Integer
        For i = 0 To dgvOrdItem.Rows.Count - 1
            dgvOrdItem.Item(0, i).Value = True
        Next
    End Sub

    Private Sub btnUnselect_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim i As Integer
        For i = 0 To dgvOrdItem.Rows.Count - 1
            dgvOrdItem.Item(0, i).Value = False
        Next
    End Sub

    Private Sub btnView_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnView.Click

        If dgvOrdItem.RowCount = 0 Then
            MsgBox("Please view the item list first before view the report. Use the apply button to list the item.")
            Return
        End If
        btnView.Enabled = False
        ConnectNavDB()


        Dim strSql As String = ""
        Dim i As Integer
        Dim dtr As SqlDataReader
        Dim dtrNav As OdbcDataReader
        Dim sItem As String = ""
        strSql = "Delete BalanceStockRep"
        ExecuteSQL(strSql)
        Dim sUOM As String = ""
        For i = 0 To dgvOrdItem.Rows.Count - 1
            If dgvOrdItem.Item(0, i).Value = True Then
                sUOM = dgvOrdItem.Item(3, i).Value
                'ExecuteSQL("Insert into BalanceStockRep(ItemNo,  UOM) values (" & SafeSQL(dgvOrdItem.Item(1, i).Value) & "," & SafeSQL(sUOM) & ")")
                ExecuteSQL("Insert into BalanceStockRep(ItemNo,  UOM, Qty, OpenQty, IQty, DQty, QtyLoose, OpenQtyLoose, IQtyLoose, DQtyLoose) values (" & SafeSQL(dgvOrdItem.Item(1, i).Value) & "," & SafeSQL(sUOM) & ",0,0,0,0,0,0,0,0)")
                sItem = sItem & "," & SafeSQL(dgvOrdItem.Item(1, i).Value)
            End If
        Next

        'GET OPENING BALANCE
        dtrNav = ReadNavRecord("Select ""Item No_"" ItemNo, sum(""Quantity"") Qty from ""Item Ledger Entry"" where ""Posting Date"" < " & SafeSQL(Format(dtpDate.Value, "yyyy-MM-dd")) & " and ""Item No_"" in ('' " & sItem & " ) group by ""Item No_""")
        While dtrNav.Read
            ExecuteSQL("Update BalanceStockRep set OpenQty = " & dtrNav("Qty").ToString & " where ItemNo = " & SafeSQL(dtrNav("ItemNo").ToString))
        End While
        dtrNav.Close()

        'GET INCREASE QTY
        dtrNav = ReadNavRecord("Select ""Item No_"" ItemNo, sum(""Quantity"") Qty from ""Item Ledger Entry"" where ""Quantity"" > 0 and ""Posting Date"" between " & SafeSQL(Format(dtpDate.Value, "yyyy-MM-dd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyy-MM-dd 23:59:59")) & " and ""Item No_"" in ('' " & sItem & " ) group by ""Item No_"" ")
        While dtrNav.Read
            ExecuteSQL("Update BalanceStockRep set IQty = " & dtrNav("Qty").ToString & " where ItemNo = " & SafeSQL(dtrNav("ItemNo").ToString))
        End While
        dtrNav.Close()


        'GET DECREASE QTY
        dtrNav = ReadNavRecord("Select ""Item No_"" ItemNo, sum(""Quantity"") Qty from ""Item Ledger Entry"" where ""Quantity"" < 0 and ""Posting Date"" between " & SafeSQL(Format(dtpDate.Value, "yyyy-MM-dd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyy-MM-dd 23:59:59")) & " and ""Item No_"" in ('' " & sItem & " ) group by ""Item No_"" ")
        While dtrNav.Read
            ExecuteSQL("Update BalanceStockRep set DQty = " & dtrNav("Qty").ToString & " where ItemNo = " & SafeSQL(dtrNav("ItemNo").ToString))
        End While
        dtrNav.Close()

        'GET CLOSING BALANCE
        dtrNav = ReadNavRecord("Select ""Item No_"" ItemNo, sum(""Quantity"") Qty  from ""Item Ledger Entry"" where ""Posting Date"" < " & SafeSQL(Format(dtpToDate.Value.AddDays(1), "yyyy-MM-dd")) & " and ""Item No_"" in ('' " & sItem & " ) group by ""Item No_""")
        While dtrNav.Read
            ExecuteSQL("Update BalanceStockRep set Qty = " & dtrNav("Qty").ToString & " where ItemNo = " & SafeSQL(dtrNav("ItemNo").ToString))
        End While
        dtrNav.Close()



        strSql = "Update BalanceStockRep set CompanyName = 'BAN SENG WINE', StockDate = " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd")) & ", ToDate = " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd")) & ", ItemName = Item.Description, Packaging = Item.ShortDesc, UnitCost = Item.CostPrice from Item where Item.ItemNo = BalanceStockRep.ItemNo" 'and Item.CompanyName = " & SafeSQL(cbCompany.Text)
        ExecuteSQL(strSql)

        strSql = "select ItemNo, UOM, isnull(UnitCost,0) as UnitCost, isnull(Qty,0) as Qty, isnull(OpenQty,0) as OpenQty, isnull(IQty,0) as IQty, isnull(DQty,0) as DQty from BalanceStockRep"
        dtr = ReadRecord(strSql)
        While dtr.Read
            'ExecuteAnotherSQL("UPdate BalanceStockRep set QtyPrint = " & SafeSQL(GetUOMQtyForPrint(dtr("ItemNo").ToString, dtr("Qty"), dtr("UOM").ToString)) & ", OpenQtyPrint = " & SafeSQL(GetUOMQtyForPrint(dtr("ItemNo").ToString, dtr("OpenQty"), dtr("UOM").ToString)) & ", IQtyPrint = " & SafeSQL(GetUOMQtyForPrint(dtr("ItemNo").ToString, dtr("IQty"), dtr("UOM").ToString)) & ", DQtyPrint = " & SafeSQL(GetUOMQtyForPrint(dtr("ItemNo").ToString, dtr("DQty"), dtr("UOM").ToString)) & " where ItemNo = " & SafeSQL(dtr("ItemNo")).ToString)
            If dtr("Qty") <> 0 Then
                ExecuteAnotherSQL("Update BalanceStockRep set SubAmt = " & GetSubAmt(dtr("ItemNo").ToString, dtr("Qty"), dtr("UOM").ToString, dtr("UnitCost")) & " where ItemNo = " & SafeSQL(dtr("ItemNo").ToString))
            End If
        End While
        dtr.Close()

        strSql = "Update BalanceStockRep set Qty = 0 where Qty is null"
        ExecuteSQL(strSql)

        strSql = "Update BalanceStockRep set UOMLoose = '' , QtyLoose = 0 where QtyLoose is null"
        ExecuteSQL(strSql)

        UpdateQty()
        UpdateIQty()
        UpdateDQty()
        UpdateOpenQty()

        DisconnectNavDB()
        strSql = "Select * from BalanceStockRep order by ItemName, Packaging"
        ExecuteReport(strSql, "BalanceStockDetRep")
        btnView.Enabled = True
    End Sub

    Private Sub btnApply_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnApply.Click
        LoadItem()
    End Sub

    Public Function GetSubAmt(ByVal sItemId As String, ByVal dQty As Double, ByVal sUOM As String, ByVal dUnitPrice As Double) As Double
        Dim dSubAmt As Double = 0
        Dim dtr As SqlDataReader
        Dim bMinus As Boolean = False
        If dQty < 0 Then
            dQty = dQty * (-1)
            bMinus = True
        End If
        dtr = ReadRecordAnother("Select * from UOM where ItemNo=" & SafeSQL(sItemId) & " Order by BaseQty Desc")
        While dtr.Read = True
            If CDbl(dtr("BaseQty")) <= dQty Then
                If sUOM = dtr("UOM").ToString Then
                    dSubAmt = dSubAmt + ((dQty - (dQty Mod CDbl(dtr("BaseQty")))) * dUnitPrice)
                Else
                    dSubAmt = dSubAmt + (Math.Round(dQty / dtr("BaseQty")) / Math.Round(1 / dtr("BaseQty")) * dUnitPrice)
                End If

                dQty = dQty Mod CDbl(dtr("BaseQty"))
            End If
        End While
        dtr.Close()

        If bMinus = True Then dSubAmt = -1 * dSubAmt
        Return dSubAmt

    End Function

    Public Function GetUOMQtyForPrint(ByVal sItemId As String, ByVal dGQty As Double, ByVal sUOM As String) As String
        Dim sGQtyDesc As String = ""
        Dim dtr As SqlDataReader
        Dim bMinus As Boolean = False
        If dGQty < 0 Then
            dGQty = dGQty * (-1)
            bMinus = True
        End If
        dtr = ReadRecordAnother("Select * from UOM where ItemNo=" & SafeSQL(sItemId) & " Order by BaseQty Desc")
        While dtr.Read = True
            If CDbl(dtr("BaseQty")) <= dGQty Then
                If sGQtyDesc.ToString.Length > 0 Then sGQtyDesc = sGQtyDesc & ", "
                'sGQtyDesc = sGQtyDesc & (dGQty / dtr("BaseQty")) & " " & dtr("UOM").ToString
                If sUOM = dtr("UOM").ToString Then
                    sGQtyDesc = sGQtyDesc & dGQty - (dGQty Mod CDbl(dtr("BaseQty"))) & " " & dtr("UOM").ToString
                Else
                    sGQtyDesc = sGQtyDesc & Math.Round((dGQty / dtr("BaseQty"))) & " " & dtr("UOM").ToString
                End If

                dGQty = dGQty Mod CDbl(dtr("BaseQty"))
            End If
        End While
        dtr.Close()
        If sGQtyDesc.Trim = "" Then
            sGQtyDesc = "0"
        End If
        If bMinus = True Then sGQtyDesc = "- " & sGQtyDesc
        Return sGQtyDesc

    End Function

    Public Sub UpdateQty()
        Dim sLooseUOM As String = ""
        Dim dLooseQty As Double = 0
        Dim dAllQty As Double = 0
        Dim strSql As String
        strSql = "select * from BalanceStockRep where round(qty, 0) <> qty"
        Dim dtr, dtrUOM As SqlDataReader
        dtr = ReadRecord(strSql)
        While dtr.Read
            dLooseQty = 0
            sLooseUOM = ""
            dAllQty = dtr("Qty")
            dtrUOM = ReadRecordAnother("Select * from UOM where ItemNo = " & SafeSQL(dtr("ItemNo").ToString) & " order by BaseQty")
            If dtrUOM.Read = True Then
                sLooseUOM = dtrUOM("UOM").ToString
                If dAllQty < 0 Then
                    dLooseQty = (dAllQty - Math.Ceiling(dAllQty)) / dtrUOM("BaseQty")
                Else
                    dLooseQty = (dAllQty - Math.Floor(dAllQty)) / dtrUOM("BaseQty")
                End If

            End If
            dtrUOM.Close()
            If dAllQty < 0 Then
                ExecuteAnotherSQL("UPdate BalanceStockRep set Qty = " & Math.Ceiling(dAllQty).ToString & ", QtyLoose = " & dLooseQty.ToString & " where ItemNo = " & SafeSQL(dtr("ItemNo")).ToString & " and UOM = " & SafeSQL(dtr("UOM").ToString))
            Else
                ExecuteAnotherSQL("UPdate BalanceStockRep set Qty = " & Math.Floor(dAllQty).ToString & ", QtyLoose = " & dLooseQty.ToString & "  where ItemNo = " & SafeSQL(dtr("ItemNo")).ToString & " and UOM = " & SafeSQL(dtr("UOM").ToString))
            End If
        End While
        dtr.Close()
    End Sub

    Public Sub UpdateOpenQty()
        Dim sLooseUOM As String = ""
        Dim dLooseQty As Double = 0
        Dim dAllQty As Double = 0
        Dim strSql As String
        strSql = "select * from BalanceStockRep where round(Openqty, 0) <> Openqty"
        Dim dtr, dtrUOM As SqlDataReader
        dtr = ReadRecord(strSql)
        While dtr.Read
            dLooseQty = 0
            sLooseUOM = ""
            dAllQty = dtr("OpenQty")
            dtrUOM = ReadRecordAnother("Select * from UOM where ItemNo = " & SafeSQL(dtr("ItemNo").ToString) & " order by BaseQty")
            If dtrUOM.Read = True Then
                sLooseUOM = dtrUOM("UOM").ToString
                If dAllQty < 0 Then
                    dLooseQty = (dAllQty - Math.Ceiling(dAllQty)) / dtrUOM("BaseQty")
                Else
                    dLooseQty = (dAllQty - Math.Floor(dAllQty)) / dtrUOM("BaseQty")
                End If
            End If
            dtrUOM.Close()
            If dAllQty < 0 Then
                ExecuteAnotherSQL("UPdate BalanceStockRep set OpenQty = " & Math.Ceiling(dAllQty).ToString & ", OpenQtyLoose = " & dLooseQty.ToString & " where ItemNo = " & SafeSQL(dtr("ItemNo")).ToString & " and UOM = " & SafeSQL(dtr("UOM").ToString))
            Else
                ExecuteAnotherSQL("UPdate BalanceStockRep set OpenQty = " & Math.Floor(dAllQty).ToString & ", OpenQtyLoose = " & dLooseQty.ToString & "  where ItemNo = " & SafeSQL(dtr("ItemNo")).ToString & " and UOM = " & SafeSQL(dtr("UOM").ToString))
            End If
        End While
        dtr.Close()
    End Sub

    Public Sub UpdateIQty()
        Dim sLooseUOM As String = ""
        Dim dLooseQty As Double = 0
        Dim dAllQty As Double = 0
        Dim strSql As String
        strSql = "select * from BalanceStockRep where round(Iqty, 0) <> Iqty"
        Dim dtr, dtrUOM As SqlDataReader
        dtr = ReadRecord(strSql)
        While dtr.Read
            dLooseQty = 0
            sLooseUOM = ""
            dAllQty = dtr("IQty")
            dtrUOM = ReadRecordAnother("Select * from UOM where ItemNo = " & SafeSQL(dtr("ItemNo").ToString) & " order by BaseQty")
            If dtrUOM.Read = True Then
                sLooseUOM = dtrUOM("UOM").ToString
                If dAllQty < 0 Then
                    dLooseQty = (dAllQty - Math.Ceiling(dAllQty)) / dtrUOM("BaseQty")
                Else
                    dLooseQty = (dAllQty - Math.Floor(dAllQty)) / dtrUOM("BaseQty")
                End If
            End If
            dtrUOM.Close()
            If dAllQty < 0 Then
                ExecuteAnotherSQL("UPdate BalanceStockRep set IQty = " & Math.Ceiling(dAllQty).ToString & ", IQtyLoose = " & dLooseQty.ToString & " where ItemNo = " & SafeSQL(dtr("ItemNo")).ToString & " and UOM = " & SafeSQL(dtr("UOM").ToString))
            Else
                ExecuteAnotherSQL("UPdate BalanceStockRep set IQty = " & Math.Floor(dAllQty).ToString & ", IQtyLoose = " & dLooseQty.ToString & "  where ItemNo = " & SafeSQL(dtr("ItemNo")).ToString & " and UOM = " & SafeSQL(dtr("UOM").ToString))
            End If
        End While
        dtr.Close()
    End Sub

    Public Sub UpdateDQty()
        Dim sLooseUOM As String = ""
        Dim dLooseQty As Double = 0
        Dim dAllQty As Double = 0
        Dim strSql As String
        strSql = "select * from BalanceStockRep where round(Dqty, 0) <> Dqty"
        Dim dtr, dtrUOM As SqlDataReader
        dtr = ReadRecord(strSql)
        While dtr.Read
            dLooseQty = 0
            sLooseUOM = ""
            dAllQty = dtr("DQty")
            dtrUOM = ReadRecordAnother("Select * from UOM where ItemNo = " & SafeSQL(dtr("ItemNo").ToString) & " order by BaseQty")
            If dtrUOM.Read = True Then
                sLooseUOM = dtrUOM("UOM").ToString
                If dAllQty < 0 Then
                    dLooseQty = (dAllQty - Math.Ceiling(dAllQty)) / dtrUOM("BaseQty")
                Else
                    dLooseQty = (dAllQty - Math.Floor(dAllQty)) / dtrUOM("BaseQty")
                End If
            End If
            dtrUOM.Close()
            If dAllQty < 0 Then
                ExecuteAnotherSQL("UPdate BalanceStockRep set DQty = " & Math.Ceiling(dAllQty).ToString & ", DQtyLoose = " & dLooseQty.ToString & " where ItemNo = " & SafeSQL(dtr("ItemNo")).ToString & " and UOM = " & SafeSQL(dtr("UOM").ToString))
            Else
                ExecuteAnotherSQL("UPdate BalanceStockRep set DQty = " & Math.Floor(dAllQty).ToString & ", DQtyLoose = " & dLooseQty.ToString & "  where ItemNo = " & SafeSQL(dtr("ItemNo")).ToString & " and UOM = " & SafeSQL(dtr("UOM").ToString))
            End If
        End While
        dtr.Close()
    End Sub

    Private Sub chkSelAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSelAll.CheckedChanged
        Dim i As Integer
        If chkSelAll.Checked = True Then
            For i = 0 To dgvOrdItem.Rows.Count - 1
                If dgvOrdItem.Rows(i).IsNewRow = True Then Exit For
                dgvOrdItem.Item(0, i).Value = True
            Next
        Else
            For i = 0 To dgvOrdItem.Rows.Count - 1
                If dgvOrdItem.Rows(i).IsNewRow = True Then Exit For
                dgvOrdItem.Item(0, i).Value = False
            Next
        End If
    End Sub

    
    
End Class