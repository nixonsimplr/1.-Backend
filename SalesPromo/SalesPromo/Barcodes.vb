Imports DataInterface.IbizDO
Imports SalesInterface.MobileSales
Imports System.Data
Imports System.Data.SqlClient

Public Class Barcodes
    Implements ISalesBase

    Private objDO As New DataInterface.IbizDO
    Dim bloading As Boolean = False
    Private Sub btnItemNo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnItemNo.Click
        RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\DataList.dll", "DataList.ItemList", "SalesPromo.Barcodes", "txtCode", 0, 0)
    End Sub

    Public Function GetListViewForm() As String Implements SalesInterface.MobileSales.ISalesBase.GetListViewForm
        Return ""
    End Function

    Public Sub ListViewClick() Implements SalesInterface.MobileSales.ISalesBase.ListViewClick

    End Sub


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
        If ResultTo = "txtCode" Then
            txtCode.Text = Value
        End If
    End Sub

    Public Sub ReturnSearch(ByVal SQL As String) Implements SalesInterface.MobileSales.ISalesBase.ReturnSearch

    End Sub

    Public Event SearchField(ByVal FormName As String, ByVal FieldName As String, ByVal FieldType As String, ByVal CurValue As String) Implements SalesInterface.MobileSales.ISalesBase.SearchField

    Public Sub SetCulture(ByVal CultureName As String) Implements SalesInterface.MobileSales.ISalesBase.SetCulture

    End Sub

    Private Sub txtCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCode.KeyPress
        e.Handled = True
    End Sub

    Private Sub txtCode_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCode.TextChanged
        Try
            Dim sCode As String = ""
            Dim dtr As SqlDataReader

            dtr = objDO.ReadRecord("Select Description from Item where ItemNo =" & objDO.SafeSQL(txtCode.Text))
            If dtr.Read = True Then
                txtDescription.Text = dtr(0).ToString
            End If
            dtr.Close()
            bLoading = True
            dtr = objDO.ReadRecord("Select * from UOM where ItemNo =" & objDO.SafeSQL(txtCode.Text))
            cmbUOM.Items.Clear()
            While dtr.Read = True
                cmbUOM.Items.Add(dtr("UOM").ToString)
            End While
            dtr.Close()
            If cmbUOM.Items.Count > 0 Then
                cmbUOM.SelectedIndex = 0
            End If
            bLoading = False
            LoadBarcodes()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub LoadBarcodes()
        If bloading = True Then Exit Sub
        Dim rs As SqlDataReader
        rs = objDO.ReadRecord("Select * from Barcodes where ItemNo=" & objDO.SafeSQL(txtCode.Text) & " and UOM=" & objDO.SafeSQL(cmbUOM.Text))
        dgvCustomerPrice.Rows.Clear()
        While rs.Read
            Dim row As String() = New String() _
            {rs("Barcode").ToString}
            dgvCustomerPrice.Rows.Add(row)
        End While
        rs.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim cnt As Integer
        Dim strSql As String
        Try
            objDO.ExecuteSQL("Delete from Barcodes where ItemNo=" & objDO.SafeSQL(txtCode.Text) & " and UOM=" & objDO.SafeSQL(cmbUOM.Text))
            For cnt = 0 To dgvCustomerPrice.Rows.Count() - 2
                strSql = "Insert into Barcodes (ItemNo, Barcode, UOM) Values (" & objDO.SafeSQL(txtCode.Text) & " , " & objDO.SafeSQL(dgvCustomerPrice(0, cnt).Value) & " , " & objDO.SafeSQL(cmbUOM.Text) & ")"
                objDO.ExecuteSQL(strSql)
            Next
            MessageBox.Show("Barcode Saved Successfully", "New Record", MessageBoxButtons.OK)
            'LoadSpecialItems()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Event LoadDataForm(ByVal Location As String, ByVal LoadType As String, ByVal ParentLoadType As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.LoadDataForm

    Public Event LoadForm(ByVal Location As String, ByVal LoadType As String) Implements SalesInterface.MobileSales.ISalesBase.LoadForm

    Private Sub cmbUOM_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbUOM.SelectedIndexChanged
        LoadBarcodes()
    End Sub
End Class