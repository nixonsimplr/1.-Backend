Imports DataInterface.IbizDO
Imports SalesInterface.MobileSales
Imports System.Data.SqlClient
Public Class LoyaltyRedemptionByItem
    Implements ISalesBase


    Private objDO As New DataInterface.IbizDO
    Private IsNewRecord As Boolean
    Private rInd As Integer
    Private Sub SaveLoyalty()
        objDO.ExecuteSQL("Delete from LoyaltyRedByItem")
        Dim cnt As Integer
        Dim strSql As String
        Try
            For cnt = 0 To dgvCustomerPrice.Rows.Count() - 2
                If dgvCustomerPrice.Item(0, cnt).Value = "" Then Exit For
                strSql = "Insert into LoyaltyRedByItem(ItemNo, UOM, Points, DTG) Values (" & objDO.SafeSQL(dgvCustomerPrice.Item(0, cnt).Value.ToString) & "," & objDO.SafeSQL(dgvCustomerPrice.Item(2, cnt).Value.ToString) & "," & dgvCustomerPrice.Item(3, cnt).Value & "," & objDO.SafeSQL(Format(DateTime.Now, "yyyyMMdd HH:mm:ss")) & ")"
                'MsgBox(strSql)
                objDO.ExecuteSQL(strSql)
            Next
            MessageBox.Show("Loyalty Redemption By Item Modified", "New Record", MessageBoxButtons.OK)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
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


    Private Sub GetRedemption()
        Try
            Dim sCode As String = ""
            Dim rs As SqlDataReader
            sCode = "Select LoyaltyRedByItem.ItemNo, Description, UOM, LoyaltyRedByItem.Points from Item, LoyaltyRedByItem where Item.ItemNo=LoyaltyRedByItem.ItemNo"
            rs = objDO.ReadRecord(sCode)
            dgvCustomerPrice.Rows.Clear()
            Dim cnt As Integer = 0
            While rs.Read
                Dim row As String() = New String() _
                {rs("ItemNo").ToString, rs("Description").ToString, "", _
                 rs("Points").ToString}
                dgvCustomerPrice.Rows.Add(row)
                Dim cb As New DataGridViewTextBoxCell
                dgvCustomerPrice.Item(2, cnt) = cb
                dgvCustomerPrice.Item(2, cnt).Value = rs("UOM")
                cnt = cnt + 1
            End While
            rs.Close()
            IsNewRecord = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        SaveLoyalty()
        IsNewRecord = False
    End Sub
    Private Sub dgvCustomerPrice_CellMouseDown1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvCustomerPrice.CellMouseDown
        If e.RowIndex < 0 Then Exit Sub
        If e.ColumnIndex < 0 Then Exit Sub
        rInd = e.RowIndex
        If e.ColumnIndex = 0 Then
            If dgvCustomerPrice.Item(0, e.RowIndex).Value = "" Then
                'If rbCustPrice.Checked = True Then
                '    'IsNewRecord = True
                '    RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\DataList.dll", "DataList.CustProd", "Customer.CustomerPrice", "dgvCustPrice", 0, 0)
                'Else
                '    RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\DataList.dll", "DataList.Itempr", "Customer.CustomerPrice", "dgvCustPrice", 0, 0)
                'End If
                RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\DataList.dll", "DataList.ItemListMultiple", "SalesPromo.LoyaltyRedemptionByItem", "dgvCustPrice", 0, 0)
            End If
        End If
    End Sub

    Private Sub dgvCustomerPrice_CellValueChanged1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCustomerPrice.CellValueChanged
        Try
            If e.RowIndex < 0 Then Exit Sub
            If e.ColumnIndex < 0 Then Exit Sub
            If e.ColumnIndex = 0 And dgvCustomerPrice.Item(0, e.RowIndex).Value <> "" Then
                Dim dtr As SqlDataReader
                dtr = objDO.ReadRecord("Select * from Item where ItemNo=" & objDO.SafeSQL(dgvCustomerPrice.Item(0, e.RowIndex).Value))
                If dtr.Read = True Then
                    dgvCustomerPrice.Item(1, e.RowIndex).Value = dtr("Description").ToString
                    Dim cb1 As New DataGridViewTextBoxCell
                    dgvCustomerPrice.Item(2, e.RowIndex) = cb1
                    dgvCustomerPrice.Item(2, e.RowIndex).Value = dtr("BaseUOM")
                End If
                dtr.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        'dgvCustomerPrice.Rows.RemoveAt(rInd)
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        dgvCustomerPrice.Rows.Clear()
    End Sub

    Private Sub dgvCustomerPrice_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCustomerPrice.CellContentClick
    End Sub

    Private Sub LoyaltyRedemptionByItem_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        GetRedemption()
    End Sub

  End Class