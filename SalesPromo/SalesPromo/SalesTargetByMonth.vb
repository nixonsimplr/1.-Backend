Imports System
Imports System.Resources
Imports System.Threading
Imports System.Reflection
Imports System.Globalization

Imports DataInterface.IbizDO
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
'Imports DataList

Public Class SalesTargetByMonth
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
    Public sItemId As String
    Private sMDT As String = ""
    Dim iExNo As Int32
    Dim rs As SqlDataReader
    Private rInd As Integer
    Dim bAgentLoading As Boolean = False

    Private Sub btnCustNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\DataList.dll", "DataList.Customer", "GoodsExchange.GoodsExchange", "txtCustNo", 0, 0)
    End Sub

    Public Sub LoadDescription()
        bAgentLoading = True
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
        bAgentLoading = False
        cmbAgent_SelectedIndexChanged(Nothing, Nothing)
    End Sub
    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub btnAgent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\DataList.dll", "DataList.Agent", "GoodsExchange.GoodsExchange", "txtSalesPersonCode", 0, 0)
    End Sub


    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        For idx As Integer = 0 To dgvTarget.Rows.Count - 1
            If IsNumeric(dgvTarget.Item(1, idx).Value) = False Or IsNumeric(dgvTarget.Item(2, idx).Value) = False Then
                MessageBox.Show("Incorrect values, please enter again")
                Exit Sub
            End If
        Next

        Dim i, j As Integer
        Dim strSql As String = ""
        arrAgentNo = New ArrayList

        If cmbAgent.Text = "ALL" Then
            Dim dtr As SqlDataReader
            dtr = objDO.ReadRecord("Select Distinct Code from SalesAgent")
            While dtr.Read()
                arrAgentNo.Add(dtr("Code").ToString)
            End While
            dtr.Close()
            dtr.Dispose()

            For j = 0 To arrAgentNo.Count - 1
                If RecordExists(arrAgentNo(j).ToString) = False Then
                    For i = 0 To dgvTarget.Rows.Count - 1
                        strSql = "Insert into SalesTargetByMonth (AgentID, Month, Target, Actual) Values (" & objDO.SafeSQL(arrAgentNo(j).ToString) & ", " & objDO.SafeSQL(dgvTarget(0, i).Value) & ", " & dgvTarget(1, i).Value & ", " & dgvTarget(2, i).Value & ")"
                        objDO.ExecuteSQL(strSql)
                    Next
                Else
                    For i = 0 To dgvTarget.Rows.Count - 1
                        strSql = "Update SalesTargetByMonth set Target = " & dgvTarget(1, i).Value & " where AgentId = " & objDO.SafeSQL(arrAgentNo(j).ToString) & " and Month = " & objDO.SafeSQL(dgvTarget(0, i).Value)
                        objDO.ExecuteSQL(strSql)
                    Next
                End If
            Next
        Else
            If RecordExists(cmbAgent.SelectedValue) = False Then
                For i = 0 To dgvTarget.Rows.Count - 1
                    strSql = "Insert into SalesTargetByMonth (AgentID, Month, Target, Actual) Values (" & objDO.SafeSQL(cmbAgent.SelectedValue) & ", " & objDO.SafeSQL(dgvTarget(0, i).Value) & ", " & objDO.SafeSQL(dgvTarget(1, i).Value) & ", " & objDO.SafeSQL(dgvTarget(2, i).Value) & ")"
                    objDO.ExecuteSQL(strSql)
                Next
            Else
                For i = 0 To dgvTarget.Rows.Count - 1
                    strSql = "Update SalesTargetByMonth set Target = " & dgvTarget(1, i).Value & " where AgentId = " & objDO.SafeSQL(cmbAgent.SelectedValue) & " and Month = " & objDO.SafeSQL(dgvTarget(0, i).Value)
                    objDO.ExecuteSQL(strSql)
                Next
            End If
        End If
        MsgBox("Saved Successfully")
    End Sub

    Private Function RecordExists(ByVal sAgentId As String) As Boolean
        Dim dtr As SqlDataReader
        dtr = objDO.ReadRecord("Select * from SalesTargetByMonth where AgentID = " & objDO.SafeSQL(sAgentId))
        If dtr.Read() Then
            RecordExists = True
        Else
            RecordExists = False
        End If
        dtr.Close()
        dtr.Dispose()

    End Function

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

    Private Sub dgvTarget_CellMouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvTarget.CellMouseDown
        If e.RowIndex < 0 Then Exit Sub
        If e.ColumnIndex < 0 Then Exit Sub
        If e.ColumnIndex = 1 Then
            dgvTarget.CurrentCell = dgvTarget.Item(e.ColumnIndex, e.RowIndex)
            dgvTarget.Item(e.ColumnIndex, e.RowIndex).ReadOnly = False
            dgvTarget.BeginEdit(True)
        Else
            dgvTarget.CurrentCell = dgvTarget.Item(e.ColumnIndex, e.RowIndex)
            dgvTarget.Item(e.ColumnIndex, e.RowIndex).ReadOnly = True
            dgvTarget.BeginEdit(False)
        End If
    End Sub


    Public Event LoadDataForm(ByVal Location As String, ByVal LoadType As String, ByVal ParentLoadType As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.LoadDataForm

    Public Event ResultData(ByVal ChildLoadType As String, ByVal Value As String) Implements SalesInterface.MobileSales.ISalesBase.ResultData

    Public Sub ReturnedData(ByVal Value As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.ReturnedData

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


    Private Sub cmbAgent_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbAgent.SelectedIndexChanged
        If bAgentLoading = True Then Exit Sub
        If cmbAgent.SelectedValue Is Nothing Then Exit Sub
        If cmbAgent.SelectedValue = "" Then Exit Sub

        Dim sAgent As String = cmbAgent.SelectedValue
        If cmbAgent.Text = "ALL" Then
            sAgent = "Admin"
        End If

        Dim rs As SqlDataReader
        rs = objDO.ReadRecord("Select * from SalesTargetByMonth where AgentID = " & objDO.SafeSQL(sAgent))
        dgvTarget.Rows.Clear()
        Dim bRows As Boolean = False
        While rs.Read
            bRows = True
            Dim row As String() = New String() _
                                {rs("Month").ToString, Format(rs("Target"), "0.00"), Format(rs("Actual"), "0.00")}
            dgvTarget.Rows.Add(row)
        End While
        rs.Close()
        rs.Dispose()

        If bRows = False Then
            Dim row As String() = New String() _
            {"January", "0.00", "0.00"}
            dgvTarget.Rows.Add(row)

            row = New String() _
            {"February", "0.00", "0.00"}
            dgvTarget.Rows.Add(row)

            row = New String() _
            {"March", "0.00", "0.00"}
            dgvTarget.Rows.Add(row)

            row = New String() _
            {"April", "0.00", "0.00"}
            dgvTarget.Rows.Add(row)

            row = New String() _
            {"May", "0.00", "0.00"}
            dgvTarget.Rows.Add(row)

            row = New String() _
            {"June", "0.00", "0.00"}
            dgvTarget.Rows.Add(row)

            row = New String() _
            {"July", "0.00", "0.00"}
            dgvTarget.Rows.Add(row)

            row = New String() _
            {"August", "0.00", "0.00"}
            dgvTarget.Rows.Add(row)

            row = New String() _
            {"September", "0.00", "0.00"}
            dgvTarget.Rows.Add(row)

            row = New String() _
            {"October", "0.00", "0.00"}
            dgvTarget.Rows.Add(row)

            row = New String() _
            {"November", "0.00", "0.00"}
            dgvTarget.Rows.Add(row)

            row = New String() _
            {"December", "0.00", "0.00"}
            dgvTarget.Rows.Add(row)

        End If
    End Sub

    Private Sub dgvTarget_CellValueChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvTarget.CellValueChanged
        If e.ColumnIndex < 0 Or e.RowIndex < 0 Then Exit Sub
        If IsNumeric(dgvTarget.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
            dgvTarget.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDbl(dgvTarget.Item(e.ColumnIndex, e.RowIndex).Value), "0.00")
        Else
            MsgBox("Incorrect values, please enter correct values")
        End If
    End Sub

   End Class