Imports System
Imports System.Resources
Imports System.Threading
Imports System.Reflection
Imports System.Globalization

Imports DataInterface.IbizDO
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
'Imports DataList

Public Class SalesTargetQty
    Implements ISalesBase


#Region "Localization Variants"
    Private rMgr As ResourceManager
    Private cInfo As CultureInfo
    Private sLang As String
#End Region
    Private arrAgentNo As New ArrayList()
    Private aAgent As New ArrayList()
    Private myAdapter As SqlDataAdapter
    Private myDataSet As DataSet
    Private myDataView As DataView
    Private RowNo As Integer
    Private objDO As New DataInterface.IbizDO
    Public sCustNo As String
    Public sCustName As String
    Private IsNewRecord As Boolean = False
    Private IsModify As Boolean = True
    Public sItemId As String
    Private sMDT As String = ""
    Dim iExNo As Int32
    Dim rs As SqlDataReader
    Private rInd As Integer

    Private Sub btnCustNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\DataList.dll", "DataList.Customer", "GoodsExchange.GoodsExchange", "txtCustNo", 0, 0)
    End Sub

    Public Sub LoadDescription()
        Dim dtr As SqlDataReader
        dtr = objDO.ReadRecord("Select Code, Name as  Description from SalesAgent")
        cmbAgent.DataSource = Nothing
        aAgent.Clear()
        aAgent.Add(New ComboValues("ALL", "ALL"))
        While dtr.Read()
            aAgent.Add(New ComboValues(dtr("Code").ToString, dtr("Description").ToString))
        End While
        dtr.Close()
        cmbAgent.DataSource = aAgent
        cmbAgent.DisplayMember = "Desc"
        cmbAgent.ValueMember = "Code"
        cmbAgent.SelectedIndex = 0
    End Sub
    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub btnAgent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\DataList.dll", "DataList.Agent", "GoodsExchange.GoodsExchange", "txtSalesPersonCode", 0, 0)
    End Sub


    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim i, j As Integer
        Dim iCnt As Integer = 0
        For i = 0 To dgvTarget.Rows.Count - 2
            If CDbl(dgvTarget.Item(4, i).Value) > 0 Then
                iCnt += 1
            End If
        Next
        Dim StrSQL As String
        If MsgBox("You are about to set Target Qty for " & iCnt & " Products." & vbCrLf & "Do you want to continue?", MsgBoxStyle.YesNo, "Sales Target") = MsgBoxResult.Yes Then
            arrAgentNo = New ArrayList
            If cmbAgent.SelectedValue = "ALL" Then
                Dim dtr As SqlDataReader
                dtr = objDO.ReadRecord("Select Distinct AgentID from MDT")
                While dtr.Read()
                    arrAgentNo.Add(dtr("AgentID").ToString)
                End While
                dtr.Close()
                objDO.ExecuteSQL("Delete from SalesTargetQty")
                For j = 0 To arrAgentNo.Count - 1
                    For i = 0 To dgvTarget.Rows.Count - 2
                        If dgvTarget.Item(0, i).Value = True Then
                            'If CDbl(dgvTarget.Item(4, i).Value) > 0 Then
                            StrSQL = "Insert into SalesTargetQty (AgentID, Month, ItemNo, UOM, TargetQty, ActualQty) Values (" & objDO.SafeSQL(arrAgentNo(j).ToString) & ", '', " & objDO.SafeSQL(dgvTarget(1, i).Value) & ", " & objDO.SafeSQL(dgvTarget(3, i).Value) & " , " & CDbl(dgvTarget(4, i).Value) & ",0)"
                            objDO.ExecuteSQL(StrSQL)
                        End If
                    Next
                Next
            Else
                For i = 0 To dgvTarget.Rows.Count - 2
                    If dgvTarget.Item(0, i).Value = True Then
                        'If CDbl(dgvTarget.Item(4, i).Value) > 0 Then
                        If IsRecExists(dgvTarget(1, i).Value, cmbAgent.SelectedValue) = True Then
                            StrSQL = "Update SalesTargetQty Set TargetQty=" & CDbl(dgvTarget(4, i).Value) & " Where AgentID=" & objDO.SafeSQL(cmbAgent.SelectedValue) & " and ItemNo= " & objDO.SafeSQL(dgvTarget(1, i).Value)
                            objDO.ExecuteSQL(StrSQL)
                        Else
                            StrSQL = "Insert into SalesTargetQty (AgentID, Month, ItemNo, UOM, TargetQty, ActualQty) Values (" & objDO.SafeSQL(cmbAgent.SelectedValue) & ", '', " & objDO.SafeSQL(dgvTarget(1, i).Value) & ", " & objDO.SafeSQL(dgvTarget(3, i).Value) & " , " & CDbl(dgvTarget(4, i).Value) & ",0)"
                            objDO.ExecuteSQL(StrSQL)
                        End If
                    End If
                Next
            End If
        End If
        MsgBox("Saved Successfully")
    End Sub
    Private Function IsRecExists(ByVal sItemNo As String, ByVal sAgent As String) As Boolean
        Dim dtr As SqlDataReader
        Dim bAns As Boolean = False
        dtr = objDO.ReadRecord("Select * from SalesTargetQty where ItemNo = " & objDO.SafeSQL(sItemNo) & " and AgentID = " & objDO.SafeSQL(sAgent))
        If dtr.Read Then
            bAns = True
        End If
        dtr.Close()
        Return bAns
    End Function

    Private Sub GoodsExchange_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        objDO.DisconnectDB()
    End Sub

    Private Sub GoodsExchange_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objDO.ConnectDB()
        LoadDescription()
        IsNewRecord = True
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub


    Public Function GetListViewForm() As String Implements SalesInterface.MobileSales.ISalesBase.GetListViewForm
        Return ""
    End Function

    Public Sub ListViewClick() Implements SalesInterface.MobileSales.ISalesBase.ListViewClick

    End Sub

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

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub


    Private Sub dgvExchange_CellMouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvTarget.CellMouseDown
        If e.RowIndex < 0 Then Exit Sub
        If e.ColumnIndex < 0 Then Exit Sub
        rInd = e.RowIndex
        If IsNewRecord = True Then
            If e.ColumnIndex = 1 Then
                RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\DataList.dll", "DataList.ItemList", "GoodsExchange.GoodsExchange", "dgvExchange", 0, 0)
            End If
        End If
    End Sub

    Private Sub dgvExchange_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvTarget.CellValueChanged
        ''Try
        ''    'If IsNewRecord = True Then
        ''    If e.ColumnIndex = 1 And dgvTarget.Item(1, e.RowIndex).Value <> "" Then
        ''        Dim dtr As SqlDataReader
        ''        dtr = objDO.ReadRecord("Select * from Item where ItemNo=" & objDO.SafeSQL(dgvTarget.Item(1, e.RowIndex).Value))
        ''        If dtr.Read = True Then
        ''            dgvTarget.Item(1, e.RowIndex).Value = dtr("Description").ToString
        ''            '            dgvExchange.Item(2, e.RowIndex).Value = dtr("BaseUOM").ToString
        ''        End If
        ''        dtr.Close()
        ''        Dim cb As New DataGridViewComboBoxCell
        ''        cb = dgvTarget.Item(3, e.RowIndex)
        ''        cb.Items.Clear()
        ''        Dim bCur As Boolean = False
        ''        dtr = objDO.ReadRecord("Select Uom from UOM where ItemNo=" & objDO.SafeSQL(dgvTarget.Item(1, e.RowIndex).Value))
        ''        While dtr.Read
        ''            cb.Items.Add(dtr("Uom"))
        ''            bCur = True
        ''        End While
        ''        dtr.Close()
        ''    End If
        ''    If e.ColumnIndex = 5 Then
        ''        If dgvTarget.Item(3, e.RowIndex).Value = "" Then
        ''            'MessageBox.Show(rMgr.GetString("Msg_SeltUOM"), rMgr.GetString("Msg_SeltUOM_Caption"), MessageBoxButtons.OK)
        ''            MessageBox.Show("Select UOM", "Information", MessageBoxButtons.OK)
        ''            '  dgvOrdItem.Item(4, e.RowIndex).Value = ""
        ''            Exit Sub
        ''        End If
        ''    End If
        ''    'End If
        ''Catch ex As Exception
        ''    MsgBox(ex.Message)
        ''End Try
    End Sub

    Public Event LoadDataForm(ByVal Location As String, ByVal LoadType As String, ByVal ParentLoadType As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.LoadDataForm

    Public Event ResultData(ByVal ChildLoadType As String, ByVal Value As String) Implements SalesInterface.MobileSales.ISalesBase.ResultData

    Public Sub ReturnedData(ByVal Value As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.ReturnedData

        ''  Dim cnt As Integer
        If ResultTo = "txtSalesPersonCode" Then
        End If
        If ResultTo = "txtCustNo" Then
        End If
        If ResultTo = "txtExchangeNo" Then
        End If
        If ResultTo = "dgvExchange" Then
            dgvTarget.Item(1, rInd).Value = Value
        End If
    End Sub


    '=====================================================================================
    'Search Function
    Public Sub ReturnSearch(ByVal SQL As String) Implements SalesInterface.MobileSales.ISalesBase.ReturnSearch

    End Sub

    Public Event SearchField(ByVal FormName As String, ByVal FieldName As String, ByVal FieldType As String, ByVal CurValue As String) Implements SalesInterface.MobileSales.ISalesBase.SearchField

    Private Sub GetTextBox()
        Dim ctl As Object
        Dim tb As TextBox
        For Each ctl In Me.Controls
            If (ctl.GetType() Is GetType(TabControl)) Then
                GetTabControl(ctl)
            End If
            If (ctl.GetType() Is GetType(Panel)) Then
                GetPanel(ctl)
            End If
            If (ctl.GetType() Is GetType(TextBox)) Then
                tb = ctl
                AddHandler tb.Enter, AddressOf TextBox_Enter
            End If
        Next
    End Sub

    Private Sub GetTabControl(ByVal obj As Object)
        Dim ctl As Object
        Dim tb As TextBox
        For Each ctl In obj.Controls
            If (ctl.GetType() Is GetType(Panel)) Then
                GetPanel(ctl)
            End If
            If (ctl.GetType() Is GetType(TabPage)) Then
                GetTabPage(ctl)
            End If
            If (ctl.GetType() Is GetType(TextBox)) Then
                tb = ctl
                AddHandler tb.Enter, AddressOf TextBox_Enter
            End If
        Next
    End Sub

    Private Sub GetTabPage(ByVal obj As Object)
        Dim ctl As Object
        Dim tb As TextBox
        For Each ctl In obj.Controls
            If (ctl.GetType() Is GetType(Panel)) Then
                GetPanel(ctl)
            End If
            If (ctl.GetType() Is GetType(TextBox)) Then
                tb = ctl
                AddHandler tb.Enter, AddressOf TextBox_Enter
            End If
        Next
    End Sub

    Private Sub GetPanel(ByVal obj As Object)
        Dim ctl As Object
        Dim tb As TextBox
        For Each ctl In obj.Controls
            If (ctl.GetType() Is GetType(TextBox)) Then
                tb = ctl
                AddHandler tb.Enter, AddressOf TextBox_Enter
            End If
        Next
    End Sub

    Private Sub TextBox_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim tb As TextBox = sender
        Try
            Dim sField As String
            Dim sFieldType As String
            sField = tb.DataBindings.Item(0).BindingMemberInfo.BindingField.ToString
            sFieldType = tb.DataBindings.Item(0).BindingMemberInfo.BindingField.GetType.Name
            If sField <> "" Then RaiseEvent SearchField(Me.ProductName & "." & Me.Name, sField, sFieldType, tb.Text)
        Catch ex As Exception
        End Try
    End Sub

    'Localization
    'Private Sub Localization()
    '    Try
    '        cInfo = New CultureInfo(sLang)
    '        Thread.CurrentThread.CurrentCulture = cInfo
    '        Thread.CurrentThread.CurrentUICulture = cInfo
    '        rMgr = New ResourceManager("GoodsExchange.GoodsExchange", [Assembly].GetExecutingAssembly())

    '        Me.Text = rMgr.GetString("GoodsExchange")
    '        lblSalePersCode.Text = rMgr.GetString("lblSalePersCode")
    '        OrdNoLabel1.Text = rMgr.GetString("Col_ExhList_ExhNo")
    '        OrdDtLabel.Text = rMgr.GetString("Col_ExhList_ExhDate")
    '        CustIdLabel1.Text = rMgr.GetString("Col_ExhList_CustNo")
    '        Label1.Text = rMgr.GetString("Col_ExhList_CustName")

    '        btnNew.Text = rMgr.GetString("btnNew")
    '        btnSave.Text = rMgr.GetString("GoodsChkIn_btnSave")
    '        btnDelete.Text = rMgr.GetString("btnDelete")
    '        btnCancel.Text = rMgr.GetString("btnCancel")
    '        btnModify.Text = rMgr.GetString("btnModify")

    '        dgvTarget.Columns("ItemNo").HeaderText = rMgr.GetString("Col_GdChkIn_ItemNo")
    '        dgvTarget.Columns("ItemName").HeaderText = rMgr.GetString("Col_GdChkIn_ItemName")
    '        dgvTarget.Columns("UOM").HeaderText = rMgr.GetString("Col_GdExh_UOM")
    '        dgvTarget.Columns("Qy").HeaderText = rMgr.GetString("Col_GdChkIn_Qty")
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message)
    '    End Try
    'End Sub

    Public Sub SetCulture(ByVal CultureName As String) Implements SalesInterface.MobileSales.ISalesBase.SetCulture
        sLang = CultureName
        '   Localization()
    End Sub

    Private Sub Label1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub cmbAgent_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbAgent.SelectedIndexChanged

        If IsNewRecord = False Then Exit Sub

        Dim rs As SqlDataReader
        rs = objDO.ReadRecord("Select ItemNo, Description, BaseUOM from Item where Active=1 and ToPDA=1")
        dgvTarget.Rows.Clear()
        Dim cnt As Integer = 0
        While rs.Read
            Dim row As String() = New String() _
                                {False, rs("ItemNo").ToString, rs("Description").ToString, "", 0}
            dgvTarget.Rows.Add(row)
            Dim cb As New DataGridViewTextBoxCell
            dgvTarget.Item(3, cnt) = cb
            dgvTarget.Item(3, cnt).Value = rs("BaseUom")
            cnt = cnt + 1
        End While
        rs.Close()
        Dim i As Integer
        For i = 0 To dgvTarget.Rows.Count - 2
            rs = objDO.ReadRecord("SELECT TargetQty FROM  SalesTargetQty Where AgentID=" & objDO.SafeSQL(cmbAgent.SelectedValue) & " and ItemNo= " & objDO.SafeSQL(dgvTarget(1, i).Value))
            If rs.Read = True Then
                If rs(0) Is System.DBNull.Value = False Then
                    If rs("TargetQty") > 0 Then
                        dgvTarget.Item(0, i).Value = True
                    End If
                    dgvTarget.Item(4, i).Value = rs("TargetQty")
                End If
            End If
            rs.Close()
        Next
    End Sub

    Private Sub chkSelAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSelAll.CheckedChanged
        Dim i As Integer
        If chkSelAll.Checked = True Then
            For i = 0 To dgvTarget.Rows.Count - 2
                dgvTarget.Item(0, i).Value = True
            Next
        Else
            For i = 0 To dgvTarget.Rows.Count - 2
                dgvTarget.Item(0, i).Value = False
            Next
        End If
    End Sub

  End Class