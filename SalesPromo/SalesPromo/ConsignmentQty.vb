Imports System
Imports System.Resources
Imports System.Threading
Imports System.Reflection
Imports System.Globalization

Imports DataInterface.IbizDO
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales

Public Class ConsignmentQty
    Implements ISalesBase


    Dim sCustNo As String = ""
    Dim bCustNoChanging As Boolean = False
    Dim bCustNoCorrection As Boolean = False
    Private rInd As Integer

    Dim iRowNo As Integer = -1
    Dim myDataAdapter As New SqlDataAdapter
    Dim myDataView As DataView
    Dim myDataSet As DataSet

    Private objDO As New DataInterface.IbizDO

    Private Sub LoadRow()
        If myDataView.Count <= 0 Then Exit Sub
        bCustNoChanging = True
        sCustNo = myDataView(iRowNo).Item("CustNo").ToString
        txtCustNo.Text = sCustNo
        bCustNoChanging = False
    End Sub

    Private Sub LoadDataSet()
        myDataSet = New DataSet
        myDataView = New DataView
        myDataAdapter = objDO.GetDataAdapter("Select distinct CustNo from ConsignmentQty order by CustNo")
        myDataAdapter.Fill(myDataSet, "ConsignmentQty")
        If myDataSet.Tables.Count > 0 Then
            myDataView = myDataSet.Tables(0).DefaultView
            iRowNo = 0
            LoadRow()
        End If
    End Sub

    Public Function GetListViewForm() As String Implements SalesInterface.MobileSales.ISalesBase.GetListViewForm
        Return ""
    End Function

    Public Sub ListViewClick() Implements SalesInterface.MobileSales.ISalesBase.ListViewClick

    End Sub

    Public Event LoadDataForm(ByVal Location As String, ByVal LoadType As String, ByVal ParentLoadType As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.LoadDataForm

    Public Event LoadForm(ByVal Location As String, ByVal LoadType As String) Implements SalesInterface.MobileSales.ISalesBase.LoadForm

    Public Sub MoveFirstClick() Implements SalesInterface.MobileSales.ISalesBase.MoveFirstClick
        If myDataView.Count > 0 Then
            iRowNo = 0
            bCustNoChanging = True
            LoadRow()
            bCustNoChanging = False
            DgvCustName.Visible = False
        End If
    End Sub

    Public Sub MoveLastClick() Implements SalesInterface.MobileSales.ISalesBase.MoveLastClick
        If myDataView.Count > 0 Then
            iRowNo = myDataView.Count - 1
            bCustNoChanging = True
            LoadRow()
            bCustNoChanging = False
            DgvCustName.Visible = False
        End If
    End Sub

    Public Sub MoveNextClick() Implements SalesInterface.MobileSales.ISalesBase.MoveNextClick
        If myDataView.Count > 0 And iRowNo < myDataView.Count - 1 Then
            iRowNo += 1
            bCustNoChanging = True
            LoadRow()
            bCustNoChanging = False
            DgvCustName.Visible = False
        End If
    End Sub

    Public Sub MovePositionClick(ByVal Position As Long) Implements SalesInterface.MobileSales.ISalesBase.MovePositionClick

    End Sub

    Public Sub MovePreviousClick() Implements SalesInterface.MobileSales.ISalesBase.MovePreviousClick
        If myDataView.Count > 0 And iRowNo > 0 Then
            iRowNo -= 1
            bCustNoChanging = True
            LoadRow()
            bCustNoChanging = False
            DgvCustName.Visible = False
        End If
    End Sub

    Public Sub Print() Implements SalesInterface.MobileSales.ISalesBase.Print

    End Sub

    Public Sub Print(ByVal PageNo As Integer) Implements SalesInterface.MobileSales.ISalesBase.Print

    End Sub

    Public Event ResultData(ByVal ChildLoadType As String, ByVal Value As String) Implements SalesInterface.MobileSales.ISalesBase.ResultData

    Public Sub ReturnedData(ByVal Value As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.ReturnedData
        If ResultTo = "dgvItems" Then
            While InStr(Value, Chr(1)) > 0
                Dim bFound As Boolean = False
                For idx As Integer = 0 To dgvItems.Rows.Count - 1
                    If dgvItems.Item(0, idx).Value <> "" Then
                        If dgvItems.Item(0, idx).Value = Mid(Value, 1, InStr(Value, Chr(1)) - 1) Then
                            If InStr(Value, Chr(1)) > 0 Then
                                Value = Mid(Value, InStr(Value, Chr(1)) + 1)
                            End If
                            bFound = True
                            Exit For
                        End If
                    End If
                Next
                If bFound = False Then
                    If dgvItems.Rows.Count - 1 = rInd Then dgvItems.Rows.Add()
                    dgvItems.Item(0, rInd).Value = Mid(Value, 1, InStr(Value, Chr(1)) - 1)
                    Value = Mid(Value, InStr(Value, Chr(1)) + 1)
                    rInd = rInd + 1
                End If
            End While

            For idx As Integer = 0 To dgvItems.Rows.Count - 1
                If dgvItems.Item(0, idx).Value <> "" Then
                    If dgvItems.Item(0, idx).Value = Value Then
                        Exit Sub
                    End If
                End If
            Next
            If dgvItems.Rows.Count - 1 = rInd Then dgvItems.Rows.Add()
            dgvItems.Item(0, rInd).Value = Value
            rInd = rInd + 1
        ElseIf ResultTo = "txtCustNo" Then
            bCustNoChanging = True
            sCustNo = Value
            txtCustNo.Text = Value
            bCustNoChanging = False
            DgvCustName.Visible = False
        End If
    End Sub

    Public Sub ReturnSearch(ByVal SQL As String) Implements SalesInterface.MobileSales.ISalesBase.ReturnSearch

    End Sub

    Public Event SearchField(ByVal FormName As String, ByVal FieldName As String, ByVal FieldType As String, ByVal CurValue As String) Implements SalesInterface.MobileSales.ISalesBase.SearchField

    Public Sub SetCulture(ByVal CultureName As String) Implements SalesInterface.MobileSales.ISalesBase.SetCulture

    End Sub

    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim cnt As Integer
        Dim strSql As String
        Try
            objDO.ExecuteSQL("Delete from ConsignmentQty where CustNo = " & objDO.SafeSQL(txtCustNo.Text))
            For cnt = 0 To dgvItems.Rows.Count() - 1
                If Not dgvItems.Item(0, cnt).Value Is Nothing Then
                    If dgvItems.Item(0, cnt).Value <> "" Then
                        If dgvItems.Item(3, cnt).Value <> "" Then
                            If CDbl(dgvItems.Item(3, cnt).Value) > 0 Then
                                strSql = "Insert into ConsignmentQty(CustNo, ItemNo, UOM, Qty) Values (" & objDO.SafeSQL(txtCustNo.Text) & " , " & objDO.SafeSQL(dgvItems(0, cnt).Value) & " , " & objDO.SafeSQL(dgvItems(2, cnt).Value) & " , " & dgvItems(3, cnt).Value & ")"
                                objDO.ExecuteSQL(strSql)
                            End If
                        End If
                    End If
                End If
            Next
            MessageBox.Show("Saved Successfully...", "Consignment Qty", MessageBoxButtons.OK)
            LoadDataSet()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub SpecialItems_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        objDO.DisconnectDB()
    End Sub

    Private Sub SpecialItems_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objDO.ConnectDB()
        LoadDataSet()
    End Sub

    Private Sub dgvItems_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItems.CellDoubleClick
        If e.RowIndex < 0 Then Exit Sub
        If e.ColumnIndex < 0 Then Exit Sub
        rInd = e.RowIndex
        If e.ColumnIndex = 0 Then
            RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\DataList.dll", "DataList.ItemListMultiple", "SalesPromo.ConsignmentQty", "dgvItems", 0, 0)
        End If
    End Sub

    Private Sub dgvItems_CellMouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvItems.CellMouseDown
        If e.ColumnIndex < 0 Or e.RowIndex < 0 Then Exit Sub
        If e.ColumnIndex = 2 Or e.ColumnIndex = 3 Then
            dgvItems.CurrentCell = dgvItems.Item(e.ColumnIndex, e.RowIndex)
            dgvItems.CurrentCell.ReadOnly = False
            dgvItems.BeginEdit(True)
        Else
            dgvItems.CurrentCell = dgvItems.Item(e.ColumnIndex, e.RowIndex)
            dgvItems.CurrentCell.ReadOnly = True
            dgvItems.BeginEdit(False)
        End If
    End Sub

    Private Sub dgvItems_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItems.CellValueChanged
        Try
            If e.RowIndex < 0 Then Exit Sub
            If e.ColumnIndex < 0 Then Exit Sub
            If e.ColumnIndex = 0 And dgvItems.Item(0, e.RowIndex).Value <> "" Then
                Dim dtr As SqlDataReader
                dtr = objDO.ReadRecord("Select * from Item where ItemNo = " & objDO.SafeSQL(dgvItems.Item(0, e.RowIndex).Value))
                If dtr.Read = True Then
                    dgvItems.Item(1, e.RowIndex).Value = dtr("Description").ToString
                    dgvItems.Item(2, e.RowIndex).Value = ""
                    Dim cb As DataGridViewComboBoxCell
                    cb = dgvItems.Item(2, e.RowIndex)
                    cb.Items.Clear()
                    Dim rs As SqlDataReader
                    rs = objDO.ReadRecordAnother("Select * from Uom where ItemNo = " & objDO.SafeSQL(dgvItems.Item(0, e.RowIndex).Value))
                    While rs.Read()
                        cb.Items.Add(rs("Uom").ToString)
                    End While
                    rs.Close()
                    rs.Dispose()
                    dgvItems.Item(2, e.RowIndex).Value = dtr("BaseUOM").ToString
                    dgvItems.Item(3, e.RowIndex).Value = "0"
                End If
                dtr.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub txtCustName_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustName.KeyDown
        If DgvCustName.SelectedCells.Count = 0 And DgvCustName.Rows.Count > 0 Then
            DgvCustName.Rows(0).Selected = True
        Else
            If e.KeyCode = Keys.Up Then
                If DgvCustName.SelectedCells(0).RowIndex <> 0 Then
                    DgvCustName.Item(1, DgvCustName.SelectedCells(0).RowIndex - 1).Selected = True
                End If
            ElseIf e.KeyCode = Keys.Down Then
                If DgvCustName.SelectedCells(0).RowIndex <> DgvCustName.RowCount - 1 Then
                    DgvCustName.Item(1, DgvCustName.SelectedCells(0).RowIndex + 1).Selected = True
                End If
            ElseIf e.KeyCode = Keys.Enter Then
                DgvCustName_Enter(Nothing, Nothing)
            End If
        End If
    End Sub

    Private Sub DgvCustName_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvCustName.Enter
        If DgvCustName.Rows.Count > 0 Then
            If DgvCustName.SelectedCells.Count > 0 Then
                DgvCustName.Visible = False
                If DgvCustName.SelectedCells.Count = 1 Then
                    sCustNo = DgvCustName.Item(0, DgvCustName.SelectedCells(0).RowIndex).Value
                Else
                    sCustNo = DgvCustName.Item(0, DgvCustName.SelectedCells(1).RowIndex).Value
                End If
                bCustNoChanging = True
                txtCustNo.Text = sCustNo
                txtCustNo_TextChanged(Nothing, Nothing)
                bCustNoChanging = False
            End If
        End If
    End Sub

    Private Sub txtCustName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCustName.TextChanged
        If bCustNoChanging = True Then Exit Sub
        DgvCustName.Visible = True
        DgvCustName.Rows.Clear()
        Dim dtr As SqlDataReader
        Dim str As String = "Select * from Customer where Active=1 and CustName like '%" & txtCustName.Text.Replace("'", "''") & "%'"
        dtr = objDO.ReadRecord(str)
        While dtr.Read
            Dim row As String() = New String() {dtr("CustNo").ToString, dtr("CustName").ToString}
            DgvCustName.Rows.Add(row)
        End While
        dtr.Close()
        dtr.Close()
        DgvCustName.BringToFront()
    End Sub

    Private Sub txtCustNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCustNo.TextChanged
        If bCustNoCorrection = True Then Exit Sub
        If bCustNoChanging = False Then
            bCustNoCorrection = True
            txtCustNo.Text = sCustNo
            bCustNoCorrection = False
            Exit Sub
        End If

        Dim rs As SqlDataReader
        rs = objDO.ReadRecord("Select * from Customer where CustNo = " & objDO.SafeSQL(txtCustNo.Text))
        If rs.Read() Then
            If Not IsDBNull(rs("CustName")) Then
                If txtCustName.Text <> rs("CustName").ToString Then
                    txtCustName.Text = rs("CustName").ToString
                End If
            End If
        End If
        rs.Close()
        dgvItems.Rows.Clear()
        rs = objDO.ReadRecord("Select ConsignmentQty.*, Item.Description from ConsignmentQty, Item where ConsignmentQty.ItemNo = Item.ItemNo and CustNo = " & objDO.SafeSQL(txtCustNo.Text))
        Dim cnt As Integer = 0
        While rs.Read = True
            Dim row As String() = New String() _
            {rs("ItemNo").ToString, rs("Description").ToString, "", rs("Qty")}
            dgvItems.Rows.Add(row)
            dgvItems.Item(2, cnt).Value = ""
            Dim cb As New DataGridViewComboBoxCell
            cb = dgvItems.Item(2, cnt)
            cb.Items.Clear()
            Dim dtr As SqlDataReader
            dtr = objDO.ReadRecordAnother("Select Uom from UOM where ItemNo = " & objDO.SafeSQL(rs("ItemNo").ToString))
            While dtr.Read
                cb.Items.Add(dtr("Uom"))
            End While
            dtr.Close()
            dtr.Dispose()
            dgvItems.Item(2, cnt).Value = rs("UOM")
            cnt = cnt + 1
        End While
        rs.Close()
        rs.Dispose()
    End Sub

    Private Sub btnCust_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCust.Click
        RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\DataList.dll", "DataList.Customer", "SalesPromo.ConsignmentQty", "txtCustNo", 0, 0)
    End Sub

    Private Sub txtCustNo_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCustNo.KeyPress
        e.Handled = True
    End Sub

    End Class