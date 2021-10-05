Imports DataInterface.IbizDO
Imports SalesInterface.MobileSales
Imports System.Data.SqlClient
Public Class GoodsInventory
    Implements ISalesBase

    Private objDO As New DataInterface.IbizDO
    Private IsNewRecord As Boolean
    Private rInd As Integer

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

    Public Sub GetInventory(ByVal sItemNo As String)
        Try
            Dim sCode As String = ""
            Dim rs As SqlDataReader
            sCode = "Select Location, GoodsInvn.ItemNo, Description, UOM, Qty from Item, GoodsInvn where Item.ItemNo=GoodsInvn.ItemNo and GoodsInvn.ItemNo=" & objDO.SafeSQL(sItemNo)
            rs = objDO.ReadRecord(sCode)
            dgvCustomerPrice.Rows.Clear()
            Dim cnt As Integer = 0
            While rs.Read
                Dim row As String() = New String() _
                {rs("Location").ToString, rs("ItemNo").ToString, rs("Description").ToString, "", _
                 rs("Qty").ToString}
                dgvCustomerPrice.Rows.Add(row)
                Dim cb As New DataGridViewTextBoxCell
                dgvCustomerPrice.Item(3, cnt) = cb
                dgvCustomerPrice.Item(3, cnt).Value = rs("UOM")
                cnt = cnt + 1
            End While
            rs.Close()
            IsNewRecord = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub GoodsInventory_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        GetInventory("")
    End Sub

   End Class