Imports System
Imports System.Resources
Imports System.Threading
Imports System.Reflection
Imports System.Globalization
Imports DataInterface.IbizDO
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales

Public Class Inventory
    Implements ISalesBase



#Region "Localization Variants"
    Private rMgr As ResourceManager
    Private cInfo As CultureInfo
    Private sLang As String
#End Region

    Private objDO As New DataInterface.IbizDO
    Private IsNewRecord As Boolean
    Private IsModify As Boolean

    Public sItemId As String
    Private sMDT As String = ""
    Dim iRtnNo As Int32
    Dim dtr As SqlDataReader

    Private Sub btnLocation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLocation.Click
        Dim cnt As Integer = 0
        Dim strsql As String
        Dim dtr1 As SqlDataReader

        dgvInventory.Rows.Clear()
        strsql = "select distinct(ItemNo),UOM,sum(qty) from goodsInvn where Location=" & objDO.SafeSQL(cmbLocation.Text) & " group by ItemNo, UOM"
        dtr1 = objDO.ReadRecord(strsql)

        While dtr1.Read = True
            dgvInventory.Rows.Add()
            dgvInventory(0, cnt).Value = dtr1(0)
            dgvInventory(2, cnt).Value = dtr1(1)
            dgvInventory(3, cnt).Value = dtr1(2)
        End While
        dtr1.Close()

        'For cnt = 0 To dgvInventory.Rows.Count - 2
        '    dtr1 = objDO.ReadRecord("select Description from Item where itemNo='" & dgvInventory.Item(0, cnt).Value & "'")
        '    While dtr1.Read
        '        dgvInventory.Item(1, cnt).Value = dtr1("Description")
        '    End While
        '    dtr1.Close()
        'Next
        'dgvInventory.ReadOnly = True
    End Sub

    Private Sub dgvItemTrans_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvInventory.CellMouseClick
        If e.RowIndex < 0 Then Exit Sub
        If e.ColumnIndex < 0 Then Exit Sub
        'If e.ColumnIndex = 0 Then
        '    Dim frm As New ItemDataList
        '    frm.frm4 = Me
        '    frm.ShowDialog()
        '    If frm.ItemId <> " " Then
        '        dgvInventory.Item(e.ColumnIndex, e.RowIndex).Value = frm.ItemId
        '        dgvInventory.Item(1, e.RowIndex).Value = frm.Desc
        '        dgvInventory.Item(2, e.RowIndex).Value = frm.uom
        '    End If
        'End If
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        dgvInventory.ReadOnly = False
        btnNew.Enabled = False
        btnSave.Enabled = True
        clear()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        SaveItemTran()
        btnNew.Enabled = True
    End Sub

    Private Sub SaveItemTran()
        If txtLocation.Text = String.Empty Then
            MsgBox(rMgr.GetString("Msg_Location"))
            Return
        End If
        Dim confirm As Integer
        confirm = MessageBox.Show(rMgr.GetString("Msg_Save"), rMgr.GetString("Msg_Save_Cap"), MessageBoxButtons.YesNo)
        If confirm = 7 Then
            clear()
            Exit Sub
        End If
        Dim cnt As Integer
        Dim strSql As String
        Try
            For cnt = 0 To dgvInventory.Rows.Count() - 2
                strSql = "Insert into GoodsInvn (Location, ItemNo, UOM, QTY) Values (" & objDO.SafeSQL(txtLocation.Text) & ", " & objDO.SafeSQL(dgvInventory(0, cnt).Value) & " , " & objDO.SafeSQL(dgvInventory(2, cnt).Value) & " , " & (dgvInventory(3, cnt).Value) & ")"
                objDO.ExecuteSQL(strSql)
            Next
            MessageBox.Show(rMgr.GetString("Msg_Inventry"), rMgr.GetString("Msg_Inventry_Cap"), MessageBoxButtons.OK)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try
        '   objDO.ExecuteSQL("Update MDT set LastDoNo=" & iRtnNo & " where MdtNo = " & objDO.SafeSQL(sMDT))
    End Sub

    Private Sub Inventory_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        objDO.DisconnectDB()
    End Sub

    Public Sub clear()
        'dgvInventory.Rows.Clear()
        'txtLocation.Text = String.Empty
    End Sub

    Private Sub Inventory_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objDO.ConnectionString = My.Settings.ConnectionString
        objDO.ConnectDB()
        btnSave.Enabled = False
        PnlButton.Visible = False
        LoadCombo()
        If cmbLocation.Items.Count > 0 Then
            cmbLocation.SelectedIndex = 0
        End If
    End Sub

    Private Sub LoadCombo()
        dtr = objDO.ReadRecord("Select Code from Location")
        Do While dtr.Read = True
            cmbLocation.Items.Add(dtr("Code"))
        Loop
        dtr.Close()
    End Sub

    Private Sub cmbLocation_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbLocation.SelectedIndexChanged
        dtr = objDO.ReadRecord("Select Name from Location where Code=" & objDO.SafeSQL(cmbLocation.Text))
        If dtr.Read = True Then
            txtLocationName.Text = dtr("Name").ToString
        End If
        dtr.Close()
        LoadData("SELECT DISTINCT GoodsInvn.ItemNo, Item.Description, GoodsInvn.UOM, SUM(GoodsInvn.Qty) AS QTY FROM GoodsInvn INNER JOIN Item ON GoodsInvn.ItemNo = Item.ItemNo where Location=" & objDO.SafeSQL(cmbLocation.Text) & " group by GoodsInvn.ItemNo, Item.Description, UOM")
    End Sub

    Private Sub LoadData(ByVal SQL As String)
        Dim dbAdapter As New SqlDataAdapter
        Dim dbDataSet As DataSet
        Try
            'dgvInventory.Rows.Clear()
            dbAdapter = objDO.GetDataAdapter(SQL)
            dbDataSet = New DataSet
            dbAdapter.Fill(dbDataSet, "GoodsInvn")
            dgvInventory.DataSource = dbDataSet.Tables("GoodsInvn")

            dgvInventory.Columns(0).HeaderText = "Item No"
            dgvInventory.Columns(1).HeaderText = "Item Name"
            dgvInventory.Columns(2).HeaderText = "UOM"
            dgvInventory.Columns(3).HeaderText = "Quantity"

            dgvInventory.Columns(0).Width = 120
            dgvInventory.Columns(1).Width = 200
            dgvInventory.Columns(2).Width = 120
            dgvInventory.Columns(3).Width = 120
            dgvInventory.ReadOnly = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


        'Dim lastrow As DataRow = Nothing
        'Dim cnt As Integer = 0
        'Dim dtr1 As SqlDataReader

        'dgvInventory.Rows.Clear()
        'dtr1 = objDO.ReadRecord(SQL)
        'While dtr1.Read = True
        '    If lastrow Is Nothing Then
        '        dgvInventory.Rows.Add()
        '        dgvInventory(0, cnt).Value = dtr1("ItemNo")
        '        dgvInventory(1, cnt).Value = dtr1("Description")
        '        dgvInventory(2, cnt).Value = dtr1("UOM")
        '        dgvInventory(3, cnt).Value = dtr1("Qty")
        '        cnt = cnt + 1
        '    Else
        '        Exit While
        '    End If
        'End While
        'dtr1.Close()
        'dgvInventory.ReadOnly = True
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
        If SQL = "" Then
            LoadData("SELECT DISTINCT GoodsInvn.ItemNo, Item.Description, GoodsInvn.UOM, SUM(GoodsInvn.Qty) AS QTY FROM GoodsInvn INNER JOIN Item ON GoodsInvn.ItemNo = Item.ItemNo where Location=" & objDO.SafeSQL(cmbLocation.Text) & " group by GoodsInvn.ItemNo, Item.Description, UOM")
        Else
            'Dim iCnt As Integer = InStr(SQL, "=")
            'Dim sSearchField As String = Microsoft.VisualBasic.Left(SQL, iCnt - 1)

            'If sSearchField.Trim() = "ItemNo" Then sSearchField = "GoodsInvn.ItemNo"
            'If sSearchField.Trim() = "Description" Then sSearchField = "Item.Description"
            'If sSearchField.Trim() = "UOM" Then sSearchField = "GoodsInvn.UOM"
            'If sSearchField.Trim() = "Qty" Then sSearchField = "GoodsInvn.Qty"

            'If sSearchField = "GoodsInvn.Qty" Then
            '    SQL = sSearchField & " = " & Microsoft.VisualBasic.Right(SQL, Len(SQL) - iCnt).Trim()
            'Else
            '    SQL = sSearchField & " = '" & Microsoft.VisualBasic.Right(SQL, Len(SQL) - iCnt).Trim() & "'"
            'End If
            LoadData("SELECT DISTINCT GoodsInvn.ItemNo, Item.Description, GoodsInvn.UOM, SUM(GoodsInvn.Qty) AS QTY FROM GoodsInvn INNER JOIN Item ON GoodsInvn.ItemNo = Item.ItemNo where Location=" & objDO.SafeSQL(cmbLocation.Text) & " and " & SQL & " group by GoodsInvn.ItemNo, Item.Description, UOM")
        End If
    End Sub

    Public Event SearchField(ByVal FormName As String, ByVal FieldName As String, ByVal FieldType As String, ByVal CurValue As String) Implements SalesInterface.MobileSales.ISalesBase.SearchField

    'Localization
    Private Sub Localization()
        Try
            cInfo = New CultureInfo(sLang)
            Thread.CurrentThread.CurrentCulture = cInfo
            Thread.CurrentThread.CurrentUICulture = cInfo
            rMgr = New ResourceManager("SalesPromo.SalesPromo", [Assembly].GetExecutingAssembly())

            Me.Text = rMgr.GetString("Inventory")
            dgvInventory.Columns(0).HeaderText = rMgr.GetString("Col_DispOrd_ItemNo")
            dgvInventory.Columns(1).HeaderText = rMgr.GetString("Col_Inv_ItemName")
            dgvInventory.Columns(2).HeaderText = rMgr.GetString("Col_Inv_UOM")
            dgvInventory.Columns(3).HeaderText = rMgr.GetString("Col_Inv_Quantity")

            '  Inv_lblLocation.Text = rMgr.GetString("Inv_lblLocation")
            btnSave.Text = rMgr.GetString("btnSave")
            btnNew.Text = rMgr.GetString("btnNew")
            btnCancel.Text = rMgr.GetString("btnCancel")
            btnModify.Text = rMgr.GetString("btnModify")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub SetCulture(ByVal CultureName As String) Implements SalesInterface.MobileSales.ISalesBase.SetCulture
        sLang = CultureName
        Localization()
    End Sub


    Private Sub dgvInventory_CellMouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvInventory.CellMouseDown
        If e.RowIndex < 0 Then Exit Sub
        If e.ColumnIndex < 0 Then Exit Sub
        Dim dcl As DataGridViewColumn
        dcl = dgvInventory.Columns(e.ColumnIndex)
        Try
            Dim sField As String
            Dim sFieldType As String
            sField = dcl.Name
            'sFieldType = dcl.Name.GetType().ToString()
            If sField.Trim() = "ItemNo" Then sField = "GoodsInvn.ItemNo"
            If sField.Trim() = "Description" Then sField = "Item.Description"
            If sField.Trim() = "UOM" Then sField = "GoodsInvn.UOM"
            If sField.Trim() = "QTY" Then sField = "GoodsInvn.Qty"

            sFieldType = "String"
            If sField <> "" Then RaiseEvent SearchField(Me.ProductName & "." & Me.Name, sField, sFieldType, dgvInventory.Item(e.ColumnIndex, e.RowIndex).Value.ToString)
        Catch ex As Exception
        End Try
    End Sub

   End Class