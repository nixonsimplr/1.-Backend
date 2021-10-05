Imports System
Imports System.Resources
Imports System.Threading
Imports System.Reflection
Imports System.Globalization
Imports DataInterface.IbizDO
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Public Class ItemPrAdjustments
    Implements ISalesBase

    Private objDO As New DataInterface.IbizDO

    Private rInd As Integer
    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim i As Integer
        Try
            For i = 0 To dgvCustomerPrice.Rows.Count - 2
                If dgvCustomerPrice.Item(0, i).Value = "" Then Continue For
                objDO.ExecuteSQL("Update ItemPr set UnitPrice=UnitPrice+" & CDbl(dgvCustomerPrice(3, i).Value) & " Where ItemNo=" & objDO.SafeSQL(dgvCustomerPrice(0, i).Value).ToString)
            Next
            MsgBox("Price Adjusted Successfully")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
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
        If ResultTo = "dgvCustPrice" Then
            While InStr(Value, Chr(1)) > 0
                If dgvCustomerPrice.Rows.Count - 1 = rInd Then dgvCustomerPrice.Rows.Add()
                dgvCustomerPrice.Item(0, rInd).Value = Mid(Value, 1, InStr(Value, Chr(1)) - 1)
                Value = Mid(Value, InStr(Value, Chr(1)) + 1)
                rInd = rInd + 1
            End While
            If dgvCustomerPrice.Rows.Count - 1 = rInd Then dgvCustomerPrice.Rows.Add()
            dgvCustomerPrice.Item(0, rInd).Value = Value
            rInd = rInd + 1
        End If
    End Sub

    Public Sub ReturnSearch(ByVal SQL As String) Implements SalesInterface.MobileSales.ISalesBase.ReturnSearch

    End Sub

    Public Event SearchField(ByVal FormName As String, ByVal FieldName As String, ByVal FieldType As String, ByVal CurValue As String) Implements SalesInterface.MobileSales.ISalesBase.SearchField

    Public Sub SetCulture(ByVal CultureName As String) Implements SalesInterface.MobileSales.ISalesBase.SetCulture

    End Sub

    Private Sub dgvCustomerPrice_CellMouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvCustomerPrice.CellMouseDown
        If e.RowIndex < 0 Then Exit Sub
        If e.ColumnIndex < 0 Then Exit Sub
        rInd = e.RowIndex
        If e.ColumnIndex = 0 Then
            If dgvCustomerPrice.Item(0, e.RowIndex).Value = "" Then
                RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\DataList.dll", "DataList.ItemListMultiple", "SalesPromo.ItemPrAdjustments", "dgvCustPrice", 0, 0)
            End If
        End If
    End Sub

    Private Sub dgvCustomerPrice_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCustomerPrice.CellValueChanged
        Try
            If e.RowIndex < 0 Then Exit Sub
            If e.ColumnIndex < 0 Then Exit Sub
            If e.ColumnIndex = 0 And dgvCustomerPrice.Item(0, e.RowIndex).Value <> "" Then
                Dim dtr As SqlDataReader
                dtr = objDO.ReadRecord("Select Item.ItemNo, Item.Description, Item.BaseUOM from Item, ItemPr where Item.ItemNo = ItemPr.ItemNo and Item.ItemNo=" & objDO.SafeSQL(dgvCustomerPrice.Item(0, e.RowIndex).Value))
                If dtr.Read = True Then
                    dgvCustomerPrice.Item(1, e.RowIndex).Value = dtr("Description").ToString
                    Dim cb1 As New DataGridViewTextBoxCell
                    dgvCustomerPrice.Item(2, e.RowIndex) = cb1
                    dgvCustomerPrice.Item(2, e.RowIndex).Value = dtr("BaseUOM")
                    dgvCustomerPrice.Item(3, e.RowIndex).Value = "0" 'dtr("UnitPrice").ToString
                Else
                    MsgBox("Item not found in Item Price")
                End If
                dtr.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

   End Class