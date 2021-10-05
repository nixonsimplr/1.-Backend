Imports System
Imports System.Resources
Imports System.Threading
Imports System.Reflection
Imports System.Globalization

Imports System.Data.SqlClient
Imports DataInterface.IbizDO
Imports SalesInterface.MobileSales

Public Class DisplayOrder
    Implements ISalesBase



#Region "Localization Variants"
    Private rMgr As ResourceManager
    Private cInfo As CultureInfo
    Private sLang As String
#End Region

    Public sItemNo, sDesc, sShortDesc As String
    Public iIndex, iCnt As Integer
    Private objDo As New DataInterface.IbizDO
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            Dim cnt As Integer
            For cnt = 0 To dgvItem.Rows.Count() - 1
                objDo.ExecuteSQL("Update Item set DisplayNo=" & cnt + 1 & " where ItemNo = " & objDo.SafeSQL(dgvItem.Item(0, cnt).Value))
            Next
            MessageBox.Show(rMgr.GetString("Msg_OrdMody"), rMgr.GetString("Msg_OrdMody_Cap"), MessageBoxButtons.OK)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        ' MsgBox(dgvItem.SelectedRows.Item(0).Index)
    End Sub

    Private Sub btnUP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUP.Click
        If dgvItem.SelectedRows.Item(0).Index > 0 Then
            iIndex = dgvItem.SelectedRows.Item(0).Index
            sItemNo = dgvItem.Item(0, iIndex - 1).Value.ToString
            sDesc = dgvItem.Item(1, iIndex - 1).Value.ToString
            sShortDesc = dgvItem.Item(2, iIndex - 1).Value.ToString
            dgvItem.Item(0, iIndex - 1).Value = dgvItem.Item(0, iIndex).Value
            dgvItem.Item(1, iIndex - 1).Value = dgvItem.Item(1, iIndex).Value
            dgvItem.Item(2, iIndex - 1).Value = dgvItem.Item(2, iIndex).Value
            dgvItem.Item(0, iIndex).Value = sItemNo
            dgvItem.Item(1, iIndex).Value = sDesc
            dgvItem.Item(2, iIndex).Value = sShortDesc
            dgvItem.Rows(iIndex - 1).Selected = True
        End If
    End Sub

    Private Sub DisplayOrder_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'SalesDataSet3.Item' table. You can move, or remove it, as needed.
        Me.ItemTableAdapter.Fill(Me.SalesDataSet3.Item)
        iCnt = dgvItem.Rows.Count

    End Sub

    Private Sub dgvItemTrans_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
        'If e.RowIndex >= 1 And e.RowIndex < iCnt Then
        '    iIndex = e.RowIndex
        'End If
    End Sub

    Private Sub btnDOWN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDOWN.Click
        iIndex = dgvItem.SelectedRows.Item(0).Index
        If iIndex < iCnt - 1 Then
            sItemNo = dgvItem.Item(0, iIndex + 1).Value.ToString
            sDesc = dgvItem.Item(1, iIndex + 1).Value.ToString
            sShortDesc = dgvItem.Item(2, iIndex + 1).Value.ToString
            dgvItem.Item(0, iIndex + 1).Value = dgvItem.Item(0, iIndex).Value
            dgvItem.Item(1, iIndex + 1).Value = dgvItem.Item(1, iIndex).Value
            dgvItem.Item(2, iIndex + 1).Value = dgvItem.Item(2, iIndex).Value
            dgvItem.Item(0, iIndex).Value = sItemNo
            dgvItem.Item(1, iIndex).Value = sDesc
            dgvItem.Item(2, iIndex).Value = sShortDesc
            dgvItem.Rows(iIndex + 1).Selected = True
        End If
    End Sub

    Private Sub PnlButton_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PnlButton.Paint

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

    'Localization
    Private Sub Localization()
        Try
            cInfo = New CultureInfo(sLang)
            Thread.CurrentThread.CurrentCulture = cInfo
            Thread.CurrentThread.CurrentUICulture = cInfo
            rMgr = New ResourceManager("SalesPromo.SalesPromo", [Assembly].GetExecutingAssembly())

            Me.Text = rMgr.GetString("DisplayOrder")
            dgvItem.Columns("ItemNoDataGridViewTextBoxColumn").HeaderText = rMgr.GetString("Col_DispOrd_ItemNo")
            dgvItem.Columns("ShortDesc").HeaderText = rMgr.GetString("Col_DispOrd_ShortDesc")
            dgvItem.Columns("DescriptionDataGridViewTextBoxColumn").HeaderText = rMgr.GetString("Col_DispOrd_Desc")
            btnSave.Text = rMgr.GetString("btnSave")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub SetCulture(ByVal CultureName As String) Implements SalesInterface.MobileSales.ISalesBase.SetCulture
        sLang = CultureName
        Localization()
    End Sub

   End Class