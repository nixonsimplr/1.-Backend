Imports System
Imports System.Resources
Imports System.Threading
Imports System.Reflection
Imports System.Globalization

Imports DataInterface.IbizDO
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales

Public Class ItemTransaction
    Implements ISalesBase



#Region "Localization Variants"
    Private rMgr As ResourceManager
    Private cInfo As CultureInfo
    Private sLang As String
#End Region

    Private myAdapter As SqlDataAdapter
    Private myDataSet As DataSet
    Private myDataView As DataView
    Private RowNo As Integer
    Private objDO As New DataInterface.IbizDO
    Private IsNewRecord As Boolean
    Private IsModify As Boolean
    Public sItemId As String
    Private sMDT As String = ""
    Dim iRtnNo As Int32
    Dim rs As SqlDataReader
    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Dim sRetNo As String = ""
        dgvItemTrans.ReadOnly = False
        btnNew.Enabled = False
        btnSave.Enabled = True
        clear()
        Dim rs As SqlDataReader
        rs = objDO.ReadRecord("Select MDTNo from System")
        If rs.Read = True Then
            sMDT = rs("MDTNo")
        End If
        rs.Close()
        rs = objDO.ReadRecord("Select * from MDT where MdtNo = " & objDO.SafeSQL(sMDT) & "")
        If rs.Read = True Then
            Dim sPre As String = rs("PreDoNo")
            Dim iLen As Int32 = rs("LenDoNo")
            iRtnNo = rs("LastDoNo") + 1
            sRetNo = sPre & StrDup(iLen - Len(sPre) - Len(CStr(iRtnNo)), "0") & CStr(iRtnNo)
        End If
        rs.Close()
        txtDocNo.Text = sRetNo
    End Sub
    Public Sub clear()
        dgvItemTrans.Rows.Clear()
        txtDocNo.Text = String.Empty
        txtLocation.Text = String.Empty
        dtpDocumentDate.Value = Date.Now
        If cmbType.Items.Count > 0 Then
            cmbType.SelectedIndex = 0
        End If
    End Sub

    Private Sub ItemTransaction_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        objDO.DisconnectDB()
    End Sub

    Private Sub ItemTransaction_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objDO.ConnectionString = My.Settings.ConnectionString
        objDO.ConnectDB()
        GetItemTrans()
        btnSave.Enabled = False
        PnlButton.Visible = False
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        SaveItemTran()
        btnNew.Enabled = True
    End Sub
    Private Sub SaveItemTran()
        If txtDocNo.Text = String.Empty Then
            MsgBox(rMgr.GetString("Msg_DocNo"))
            Return
        End If
        If cmbType.Text = String.Empty Then
            MsgBox(rMgr.GetString("Msg_DocType"))
            Return
        End If
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
            strSql = "Insert into ItemTrans (DocNo, DocDt, DocType, Location) Values (" & objDO.SafeSQL(txtDocNo.Text) & ", " & objDO.SafeSQL(Format(dtpDocumentDate.Value, "dd-MMM-yyyy")) & ", " & objDO.SafeSQL(cmbType.Text) & ", " & objDO.SafeSQL(txtLocation.Text) & ")"
            objDO.ExecuteSQL(strSql)
            For cnt = 0 To dgvItemTrans.Rows.Count() - 2
                strSql = "Insert into ItemTranDet (DocNo, ItemNo, UOM, QTY) Values (" & objDO.SafeSQL(txtDocNo.Text) & ", " & objDO.SafeSQL(dgvItemTrans(0, cnt).Value) & " , " & objDO.SafeSQL(dgvItemTrans(2, cnt).Value) & " , " & (dgvItemTrans(3, cnt).Value) & ")"
                objDO.ExecuteSQL(strSql)
            Next
            MessageBox.Show(rMgr.GetString("Msg_ItemTran"), rMgr.GetString("Msg_Inventry_Cap"), MessageBoxButtons.OK)
            objDO.ExecuteSQL("Update MDT set LastDoNo=" & iRtnNo & " where MdtNo = " & objDO.SafeSQL(sMDT))
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try

        clear()
    End Sub


    Private Sub btnLocation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLocation.Click
        'Dim frm As New LocationList
        'frm.frm = Me
        'frm.ShowDialog()
        'txtLocation.Text = frm.GetLocation()
    End Sub

    Private Sub dgvItemTrans_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvItemTrans.CellMouseClick
        '  If IsNewRecord = True Or IsModify = True Then
        'If e.RowIndex < 0 Then Exit Sub
        'If e.ColumnIndex < 0 Then Exit Sub
        'If e.ColumnIndex = 0 Then
        '    Dim frm As New ItemDataList
        '    'frm.frm3 = Me
        '    'frm.ShowDialog()
        '    ' If frm.ItemId <> " " Then
        '    dgvItemTrans.Item(e.ColumnIndex, e.RowIndex).Value = frm.ItemId
        '    dgvItemTrans.Item(1, e.RowIndex).Value = frm.Desc
        '    dgvItemTrans.Item(2, e.RowIndex).Value = frm.uom
        'End If
        ''End If
        ' End If
    End Sub

    Private Sub btnDocNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDocNo.Click
        'Dim lastrow As DataRow = Nothing
        'Dim cnt As Integer = 0
        'clear()
        'Dim frm As New DocList
        'frm.frm = Me
        'frm.ShowDialog()
        'Dim strsql As String
        'Dim dtr1 As SqlDataReader
        'strsql = "(Select * from ItemTranDet where DocNo='" & txtDocNo.Text & "')"
        'dtr1 = objDO.ReadRecord(strsql)

        'While dtr1.Read = True

        '    If lastrow Is Nothing Then
        '        '    Dim row As String() = New String() _
        '        '{dtr1("ItemId").ToString, dtr1("Description").ToString, "", dtr1("minqty").ToString, dtr1("Maxqty").ToString}
        '        dgvItemTrans.Rows.Add()
        '        dgvItemTrans(0, cnt).Value = dtr1("ItemNo")
        '        'dgvOrdItem(1, cnt).Value = dtr1("ItemName")
        '        dgvItemTrans(2, cnt).Value = dtr1("UOM")
        '        dgvItemTrans(3, cnt).Value = dtr1("QTY")
        '        cnt = cnt + 1
        '    Else
        '        Exit While
        '    End If
        'End While
        'dtr1.Close()
        'For cnt = 0 To dgvItemTrans.Rows.Count - 2
        '    dtr1 = objDO.ReadRecord("select Description from Item where itemNo='" & dgvItemTrans.Item(0, cnt).Value & "'")
        '    While dtr1.Read
        '        dgvItemTrans.Item(1, cnt).Value = dtr1("Description")
        '    End While
        '    dtr1.Close()
        'Next
        'dgvItemTrans.ReadOnly = True
    End Sub

    Public Function GetListViewForm() As String Implements SalesInterface.MobileSales.ISalesBase.GetListViewForm
        Return Windows.Forms.Application.StartupPath & "\SalesPromo.dll,SalesPromo.ItemTransList"
    End Function

    Public Sub ListViewClick() Implements SalesInterface.MobileSales.ISalesBase.ListViewClick

    End Sub

    Public Event LoadForm(ByVal Location As String, ByVal LoadType As String) Implements SalesInterface.MobileSales.ISalesBase.LoadForm

    Public Sub MoveFirstClick() Implements SalesInterface.MobileSales.ISalesBase.MoveFirstClick
        If IsNewRecord = True Then
            If MsgBox(rMgr.GetString("Msg_RecordNotSaved"), MsgBoxStyle.YesNo, rMgr.GetString("Msg_RecordNotSaved_Cap")) = MsgBoxResult.No Then
                Exit Sub
            End If
        End If
        RowNo = 0
        LoadRow()
    End Sub

    Public Sub MoveLastClick() Implements SalesInterface.MobileSales.ISalesBase.MoveLastClick
        If IsNewRecord = True Then
            If MsgBox(rMgr.GetString("Msg_RecordNotSaved"), MsgBoxStyle.YesNo, rMgr.GetString("Msg_RecordNotSaved_Cap")) = MsgBoxResult.No Then
                Exit Sub
            End If
        End If
        RowNo = myDataView.Count - 1
        LoadRow()
    End Sub

    Public Sub MoveNextClick() Implements SalesInterface.MobileSales.ISalesBase.MoveNextClick
        If IsNewRecord = True Then
            If MsgBox(rMgr.GetString("Msg_RecordNotSaved"), MsgBoxStyle.YesNo, rMgr.GetString("Msg_RecordNotSaved_Cap")) = MsgBoxResult.No Then
                Exit Sub
            End If
        End If
        If RowNo >= myDataView.Count - 1 Then Exit Sub
        RowNo = RowNo + 1
        LoadRow()
    End Sub

    Public Sub MovePositionClick(ByVal Position As Long) Implements SalesInterface.MobileSales.ISalesBase.MovePositionClick

    End Sub

    Public Sub MovePreviousClick() Implements SalesInterface.MobileSales.ISalesBase.MovePreviousClick
        If IsNewRecord = True Then
            If MsgBox(rMgr.GetString("Msg_RecordNotSaved"), MsgBoxStyle.YesNo, rMgr.GetString("Msg_RecordNotSaved_Cap")) = MsgBoxResult.No Then
                Exit Sub
            End If
        End If
        If RowNo <= 0 Then Exit Sub
        RowNo = RowNo - 1
        LoadRow()
    End Sub

    Public Sub Print() Implements SalesInterface.MobileSales.ISalesBase.Print

    End Sub

    Public Sub Print(ByVal PageNo As Integer) Implements SalesInterface.MobileSales.ISalesBase.Print

    End Sub

    Private Sub GetItemTrans()
        myAdapter = objDO.GetDataAdapter("SELECT * from ItemTrans")
        myDataSet = New DataSet
        myAdapter.Fill(myDataSet, "ItemTrans")
        myDataView = New DataView(myDataSet.Tables("ItemTrans"))
        If myDataView.Count = 0 Then
            RowNo = -1
        Else
            RowNo = 0
        End If

        LoadRow()
    End Sub
    Private Sub LoadRow()
        If RowNo < 0 Or RowNo >= myDataView.Count Then Exit Sub
        txtDocNo.Text = myDataView(RowNo).Item("DocNo").ToString
        dtpDocumentDate.Value = myDataView(RowNo).Item("DocDt").ToString
        cmbType.Text = myDataView(RowNo).Item("DocType").ToString
        txtLocation.Text = myDataView(RowNo).Item("Location").ToString
    End Sub

    Private Sub txtDocNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDocNo.TextChanged
        If IsNewRecord = True Then Exit Sub
        rs = objDO.ReadRecord("SELECT ItemTrans.ItemID, ItemTrans.Uom, ItemTrans.QTY, Item.Description FROM ItemTrans INNER JOIN Item ON ItemTrans.ItemID = Item.ItemNo where ItemTrans.DocNo=" & objDO.SafeSQL(txtDocNo.Text))
        dgvItemTrans.Rows.Clear()
        While rs.Read
            Dim row As String() = New String() _
            {rs("ItemID").ToString, rs("Description").ToString, rs("UOM").ToString, _
             rs("QTY").ToString}
            dgvItemTrans.Rows.Add(row)
        End While
        rs.Close()
    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub btnModify_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnModify.Click

    End Sub

    Public Event LoadDataForm(ByVal Location As String, ByVal LoadType As String, ByVal ParentLoadType As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.LoadDataForm


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

            Me.Text = rMgr.GetString("ItemTransaction")
            dgvItemTrans.Columns("ItemID").HeaderText = rMgr.GetString("Col_DispOrd_ItemNo")
            dgvItemTrans.Columns("ItemName").HeaderText = rMgr.GetString("Col_Inv_ItemName")
            dgvItemTrans.Columns("BaseUOM").HeaderText = rMgr.GetString("Col_Inv_UOM")
            dgvItemTrans.Columns("Qty").HeaderText = rMgr.GetString("Col_Inv_Quantity")

            Item_lblDocNo.Text = rMgr.GetString("Item_lblDocNo")
            Item_lblDocDate.Text = rMgr.GetString("Item_lblDocDate")
            Item_lblLocation.Text = rMgr.GetString("Inv_lblLocation")
            Item_lblType.Text = rMgr.GetString("Item_lblType")

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

   End Class