Imports System
Imports System.Resources
Imports System.Threading
Imports System.Reflection
Imports System.Globalization

Imports DataInterface.IbizDO
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Public Class SpecialItems
    Implements ISalesBase

    Private rInd As Integer
    Private objDO As New DataInterface.IbizDO
    Dim IsNewRecord As Boolean = False




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
        If ResultTo = "dgvSpecial" Then
            While InStr(Value, Chr(1)) > 0
                If dgvSpecial.Rows.Count - 1 = rInd Then dgvSpecial.Rows.Add()
                dgvSpecial.Item(0, rInd).Value = Mid(Value, 1, InStr(Value, Chr(1)) - 1)
                Value = Mid(Value, InStr(Value, Chr(1)) + 1)
                rInd = rInd + 1
            End While
            If dgvSpecial.Rows.Count - 1 = rInd Then dgvSpecial.Rows.Add()
            dgvSpecial.Item(0, rInd).Value = Value
            rInd = rInd + 1
        End If
    End Sub

    Public Sub ReturnSearch(ByVal SQL As String) Implements SalesInterface.MobileSales.ISalesBase.ReturnSearch

    End Sub

    Public Event SearchField(ByVal FormName As String, ByVal FieldName As String, ByVal FieldType As String, ByVal CurValue As String) Implements SalesInterface.MobileSales.ISalesBase.SearchField

    Public Sub SetCulture(ByVal CultureName As String) Implements SalesInterface.MobileSales.ISalesBase.SetCulture

    End Sub

    Public Sub LoadSpecialItems()
        Dim rs As SqlDataReader
        rs = objDO.ReadRecord("Select ItemNo, ItemName, UOM from SpecialItem")
        dgvSpecial.Rows.Clear()
        Dim cnt As Integer = 0
        While rs.Read
            Dim row As String() = New String() _
                                {rs("ItemNo").ToString, rs("ItemName").ToString, rs("UOM").ToString}
            dgvSpecial.Rows.Add(row)
            cnt = cnt + 1
        End While
        rs.Close()
    End Sub
    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim cnt As Integer
        Dim strSql As String
        Try
            objDO.ExecuteSQL("Delete from SpecialItem")
            For cnt = 0 To dgvSpecial.Rows.Count() - 2
                strSql = "Insert into SpecialItem (ItemNo, ItemName, UOM) Values (" & objDO.SafeSQL(dgvSpecial(0, cnt).Value) & " , " & objDO.SafeSQL(dgvSpecial(1, cnt).Value) & " , " & objDO.SafeSQL(dgvSpecial(2, cnt).Value) & ")"
                objDO.ExecuteSQL(strSql)
            Next
            MessageBox.Show("Special Items Saved Successfully...", "New Record", MessageBoxButtons.OK)
            LoadSpecialItems()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub SpecialItems_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        objDO.DisconnectDB()
    End Sub

    Private Sub SpecialItems_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objDO.ConnectDB()
        LoadSpecialItems()
    End Sub

    Private Sub dgvSpecial_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSpecial.CellDoubleClick
        If e.RowIndex < 0 Then Exit Sub
        If e.ColumnIndex < 0 Then Exit Sub
        rInd = e.RowIndex
        If e.ColumnIndex = 0 Then
            IsNewRecord = True
            RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\DataList.dll", "DataList.ItemListMultiple", "SalesPromo.SpecialItems", "dgvSpecial", 0, 0)
        End If
    End Sub

    Private Sub dgvSpecial_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSpecial.CellValueChanged
        Try
            If e.RowIndex < 0 Then Exit Sub
            If e.ColumnIndex < 0 Then Exit Sub
            If e.ColumnIndex = 0 And dgvSpecial.Item(0, e.RowIndex).Value <> "" Then
                Dim dtr As SqlDataReader
                dtr = objDO.ReadRecord("Select * from Item where ItemNo=" & objDO.SafeSQL(dgvSpecial.Item(0, e.RowIndex).Value))
                If dtr.Read = True Then
                    dgvSpecial.Item(1, e.RowIndex).Value = dtr("Description").ToString
                    dgvSpecial.Item(2, e.RowIndex).Value = dtr("BaseUOM").ToString
                End If
                dtr.Close()
                IsNewRecord = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class