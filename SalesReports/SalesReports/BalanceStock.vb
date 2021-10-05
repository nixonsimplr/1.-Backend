Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Odbc
Imports SalesInterface.MobileSales
Imports System.IO

Public Class BalanceStock
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

    Private Sub BalanceStock_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub

    Private Sub BalanceStock_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
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
        'and companyName = " & SafeSQL(cbCompany.Text) & " 
        rs = ReadRecord("Select distinct ItemNo,Description+' '+ShortDesc as Description, BaseUOM, UnitPrice, ToPda, Barcode, CompanyName from Item where Active = 1 " & sItem & " order by Description+' '+ShortDesc")

        dgvOrdItem.Rows.Clear()
        While rs.Read
            Dim row As String() = New String() _
                        {True, rs("ItemNo").ToString, rs("Description").ToString, rs("BaseUOM").ToString, rs("UnitPrice").ToString, rs("Barcode").ToString, rs("CompanyName").ToString}
            dgvOrdItem.Rows.Add(row)

        End While
        rs.Close()
    End Sub

   

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        'select [Item No_] ItemNo, sum([Remaining Quantity]) Qty, [Unit of Measure Code] UOM from dbo.[Two-T Trading$Item Ledger Entry]
        'group by [Item No_], [Unit of Measure Code]
        btnView.Enabled = False
        If dgvOrdItem.RowCount = 0 Then
            MsgBox("Please view the item list first before view the report. Click the Refresh button to list the item.")
            Return
        End If

        ConnectNavDB()


        Dim strSql As String = ""
        Dim i As Integer
        Dim dtr, dtrUOM As SqlDataReader
        Dim dtrNav As OdbcDataReader
        Dim sItem As String = ""
        strSql = "Delete BalanceStockRep"
        ExecuteSQL(strSql)
        Dim sUOM As String = ""
        For i = 0 To dgvOrdItem.Rows.Count - 1
            If dgvOrdItem.Item(0, i).Value = True Then
                sUOM = dgvOrdItem.Item(3, i).Value
                ExecuteSQL("Insert into BalanceStockRep(ItemNo,  UOM) values (" & SafeSQL(dgvOrdItem.Item(1, i).Value) & "," & SafeSQL(sUOM) & ")")
                sItem = sItem & "," & SafeSQL(dgvOrdItem.Item(1, i).Value)
            End If
        Next

        dtrNav = ReadNavRecord("Select ""Item No_"" as ItemNo, sum(""Quantity"") as Qty from ""Item Ledger Entry"" where ""Posting Date"" < " & SafeSQL(Format(dtpDate.Value.AddDays(1), "yyyy-MM-dd")) & " and ""Item No_"" in ('' " & sItem & " ) group by ""Item No_""")

        'dtrNav = ReadNavRecord("Select ""Item No_"" as ItemNo, sum(""Quantity"") as Qty from ""Item Ledger Entry"" where ""Item No_"" in ('' " & sItem & " ) group by ""Item No_""")

        While dtrNav.Read
            'ExecuteSQL("Insert into BalanceStockRep (ItemNo,UOM,Qty) values(" & SafeSQL(dtr("ItemNo").ToString) & "," & SafeSQL(dtr("UOM").ToString) & "," & dtr("Qty").ToString & ")")
            ExecuteSQL("Update BalanceStockRep set Qty = " & dtrNav("Qty").ToString & " where ItemNo = " & SafeSQL(dtrNav("ItemNo").ToString))
        End While
        dtrNav.Close()

        Dim sLooseUOM As String = ""
        Dim dLooseQty As Double = 0
        Dim dAllQty As Double = 0
        strSql = "select * from BalanceStockRep where round(qty, 0) <> qty"
        dtr = ReadRecord(strSql)
        While dtr.Read
            dLooseQty = 0
            sLooseUOM = ""
            dAllQty = dtr("Qty")
            dtrUOM = ReadRecordAnother("Select * from UOM where ItemNo = " & SafeSQL(dtr("ItemNo").ToString) & " order by BaseQty")
            If dtrUOM.Read = True Then
                sLooseUOM = dtrUOM("UOM").ToString
                'dLooseQty = (dAllQty - Math.Floor(dAllQty)) / dtrUOM("BaseQty")
                If dAllQty < 0 Then
                    dLooseQty = (dAllQty - Math.Ceiling(dAllQty)) / dtrUOM("BaseQty")
                Else
                    dLooseQty = (dAllQty - Math.Floor(dAllQty)) / dtrUOM("BaseQty")
                End If
            End If
            dtrUOM.Close()
            If dAllQty < 0 Then
                ExecuteAnotherSQL("UPdate BalanceStockRep set Qty = " & Math.Ceiling(dAllQty).ToString & ", UOMLoose = " & SafeSQL(sLooseUOM) & ", QtyLoose = " & dLooseQty.ToString & "  where ItemNo = " & SafeSQL(dtr("ItemNo")).ToString)
            Else
                ExecuteAnotherSQL("UPdate BalanceStockRep set Qty = " & Math.Floor(dAllQty).ToString & ", UOMLoose = " & SafeSQL(sLooseUOM) & ", QtyLoose = " & dLooseQty.ToString & "  where ItemNo = " & SafeSQL(dtr("ItemNo")).ToString)
            End If

        End While
        dtr.Close()

        strSql = "Update BalanceStockRep set CompanyName = 'BAN SENG WINE', StockDate = " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd")) & ", ItemName = Item.Description, Packaging = Item.ShortDesc from Item where Item.ItemNo = BalanceStockRep.ItemNo"
        ExecuteSQL(strSql)

        strSql = "Update BalanceStockRep set Qty = 0 where Qty is null"
        ExecuteSQL(strSql)

        strSql = "Update BalanceStockRep set UOMLoose = '' , QtyLoose = 0 where QtyLoose is null"
        ExecuteSQL(strSql)

        DisconnectNavDB()
        strSql = "Select * from BalanceStockRep order by ItemName, Packaging"
        ExecuteReport(strSql, "BalanceStockRep")
        btnView.Enabled = True
    End Sub

    Private Sub btnApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnApply.Click
        LoadItem()
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