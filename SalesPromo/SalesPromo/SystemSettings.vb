Imports System
Imports System.Resources
Imports System.Threading
Imports System.Reflection
Imports System.Globalization
Imports System.IO
Imports System.Xml
Imports System.Text
Imports System.Security.Cryptography
Imports Microsoft.VisualBasic
Imports DataInterface.IbizDO
Imports System.Data
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales

Public Class SystemSettings
    Implements ISalesBase



#Region "Localization Variants"
    Private rMgr As ResourceManager
    Private cInfo As CultureInfo
    Private sLang As String
#End Region

    Private IsNewRecord As Boolean = False
    Private myAdapter As SqlDataAdapter
    Private myDataSet As DataSet
    Private myDataView As DataView
    Private RowNo As Integer
    Private objDo As New DataInterface.IbizDO

    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.txtMaxInv.Text = String.Empty Then
            MessageBox.Show("Please Enter Max Invoice", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.txtMaxInv.Select()
            Return
        End If
        'If Me.txtComName.Text = String.Empty Then
        '    MessageBox.Show("Please Enter Company Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
        '    Me.txtComName.Select()
        '    Return
        'End If
        If Me.txtGST.Text = String.Empty Then
            MessageBox.Show("Please Enter GST", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.txtGST.Select()
            Return
        End If
        If Me.txtSchAmt.Text = String.Empty Then
            MessageBox.Show("Please Enter Schedule Admin Charges", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.txtSchAmt.Select()
            Return
        End If
        If Me.txtHisNo.Text = String.Empty Then
            MessageBox.Show("Please Enter History No", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.txtSchAmt.Select()
            Return
        End If
        If Me.txtDefItemTerms.Text = String.Empty Then
            MessageBox.Show("Please Enter Default Item Terms", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.txtDefItemTerms.Select()
            Return
        End If
        Dim bLGST As Integer = 0
        If chkLGST.Checked = True Then
            bLGST = 1
        Else
            bLGST = 0
        End If
        Try
            Dim uptQry As String = "Update System set MDTNo=" & objDo.SafeSQL(cmbMDT.Text) & "," & "MaxInvPrint=" & txtMaxInv.Text & "," & "GST=" & txtGST.Text & "," & "CompanyName=" & objDo.SafeSQL(txtComName.Text) & "," & "Header1=" & objDo.SafeSQL(txtHeader1.Text) & "," & "Header2=" & objDo.SafeSQL(txtHeader2.Text) & "," & "Header3=" & objDo.SafeSQL(txtHeader3.Text) & "," & "Header4=" & objDo.SafeSQL(txtHeader4.Text) & "," & "Tail1=" & objDo.SafeSQL(txtTail1.Text) & "," & "Tail2=" & objDo.SafeSQL(txtTail2.Text) & "," & "Tail3=" & objDo.SafeSQL(txtTail3.Text) & "," & "Tail4=" & objDo.SafeSQL(txtTail4.Text) & "," & "HistoryUnit=" & objDo.SafeSQL(cmbHisUnit.Text) & "," & "HistoryNo=" & txtHisNo.Text & "," & "BankAccount=" & objDo.SafeSQL(txtBankAC.Text) & "," & "ExchangeLocation=" & objDo.SafeSQL(txtBadLocation.Text) & "," & "MembershipItemNo=" & objDo.SafeSQL(txtSchAmt.Text) & "," & "TradeDealPriceGroup=" & objDo.SafeSQL(txtPrGroup.Text) & "," & "DefPriceGroup=" & objDo.SafeSQL(txtDefPrGroup.Text) & "," & "DefSalesType=" & objDo.SafeSQL(cmbSalesType.Text) & "," & "DefaultItemTerms=" & objDo.SafeSQL(txtDefItemTerms.Text) & "," & "LoyaltyGST=" & bLGST & "," & "AutoMember=" & txtAutoMember.Text & "," & "MaxDiscount=0, " & "PrinterName=" & objDo.SafeSQL(txtPrinterName.Text.ToString) & "," & "DefCategory=" & objDo.SafeSQL(txtDefCategory.Text.ToString) & ", ExchangePer=" & txtExchangePer.Text.ToString
            objDo.ExecuteSQL(uptQry)
            MessageBox.Show("System Settings Updated Successfully...")
            '  MessageBox.Show(rMgr.GetString("Msg_UpdatesSuccess"), rMgr.GetString("Msg_UpdatesSuccess_Cap"), MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub GetUser()
        myAdapter = objDo.GetDataAdapter("select * from System")
        myDataSet = New DataSet
        myAdapter.Fill(myDataSet, "System")
        myDataView = New DataView(myDataSet.Tables("System"))
        If myDataView.Count = 0 Then
            RowNo = -1
        Else
            RowNo = 0
        End If
        LoadRow()
    End Sub
    Private Sub LoadRow()
        If RowNo < 0 Or RowNo >= myDataView.Count Then Exit Sub
        cmbMDT.Text = myDataView(RowNo).Item("MDTNO").ToString
        txtGST.Text = myDataView(RowNo).Item("GST").ToString
        txtMaxInv.Text = myDataView(RowNo).Item("MaxInvPrint").ToString
        txtComName.Text = myDataView(RowNo).Item("CompanyName").ToString
        txtHeader1.Text = myDataView(RowNo).Item("Header1").ToString
        txtHeader2.Text = myDataView(RowNo).Item("Header2").ToString
        txtHeader3.Text = myDataView(RowNo).Item("Header3").ToString
        txtHeader4.Text = myDataView(RowNo).Item("Header4").ToString
        txtTail1.Text = myDataView(RowNo).Item("Tail1").ToString
        txtTail2.Text = myDataView(RowNo).Item("Tail2").ToString
        txtTail3.Text = myDataView(RowNo).Item("Tail3").ToString
        txtTail4.Text = myDataView(RowNo).Item("Tail4").ToString
        cmbHisUnit.Text = myDataView(RowNo).Item("HistoryUnit").ToString
        txtHisNo.Text = myDataView(RowNo).Item("HistoryNo").ToString
        txtBankAC.Text = myDataView(RowNo).Item("BankAccount").ToString
        txtBadLocation.Text = myDataView(RowNo).Item("ExchangeLocation").ToString
        txtSchAmt.Text = myDataView(RowNo).Item("MembershipItemNo").ToString
        'txtAccessLevel.Text = myDataView(RowNo).Item("Access").ToString
        'chkActive.Checked = myDataView(RowNo).Item("Active")
        txtPrGroup.Text = myDataView(RowNo).Item("TradeDealPriceGroup").ToString
        txtDefPrGroup.Text = myDataView(RowNo).Item("DefPriceGroup").ToString
        cmbSalesType.Text = myDataView(RowNo).Item("DefSalesType").ToString
        txtAutoMember.Text = myDataView(RowNo).Item("AutoMember").ToString
        txtDefItemTerms.Text = myDataView(RowNo).Item("DefaultItemTerms").ToString
        txtPrinterName.Text = myDataView(RowNo).Item("PrinterName").ToString
        txtMaxDiscount.Text = myDataView(RowNo).Item("LastImExDate").ToString
        chkLGST.Checked = CBool(myDataView(RowNo).Item("LoyaltyGST"))
        txtExchangePer.Text = myDataView(RowNo).Item("ExchangePer").ToString
        txtDefCategory.Text = myDataView(RowNo).Item("DefCategory").ToString
    End Sub

    Private Sub SystemSettings_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        objDo.DisconnectDB()
    End Sub

    Private Sub SystemSettings_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objDo.ConnectionString = My.Settings.ConnectionString
        objDo.ConnectDB()
        GetUser()
        Me.txtGST.Select()
        LoadMDT()
        cmbHisUnit.SelectedIndex = 0
    End Sub
    Public Sub LoadMDT()
        Dim dtr As SqlDataReader
        dtr = objDo.ReadRecord("Select MDTNO from MDT")
        Do While dtr.Read = True
            If dtr("MDTNO") <> "" Then
                cmbMDT.Items.Add(dtr(0))
            End If
        Loop
        dtr.Close()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Clear()
        GetUser()
    End Sub
    Private Sub Clear()
        txtGST.Select()
        txtGST.Text = ""
        txtMaxInv.Text = ""
        txtComName.Text = ""
        txtHeader1.Text = ""
        txtHeader2.Text = ""
        txtHeader3.Text = ""
        txtHeader4.Text = ""
        txtTail1.Text = ""
        txtTail2.Text = ""
        txtTail3.Text = ""
        txtTail4.Text = ""
        'txtHisUnit.Text = ""
        txtHisNo.Text = ""
        txtBankAC.Text = ""
        '   cmbSalesType.SelectedIndex = 0
        txtPrinterName.Text = ""
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
        If ResultTo = "txtMemItemNo" Then
            txtSchAmt.Text = Value
        End If
        If ResultTo = "txtPrGroup" Then
            txtPrGroup.Text = Value
        End If
        If ResultTo = "txtDefPrGroup" Then
            txtDefPrGroup.Text = Value
        End If
        If ResultTo = "txtDefItemTerms" Then
            txtDefItemTerms.Text = Value
        End If
    End Sub

    Private Sub pnlLogin_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles pnlLogin.Paint

    End Sub

    Public Sub ReturnSearch(ByVal SQL As String) Implements SalesInterface.MobileSales.ISalesBase.ReturnSearch

    End Sub

    Public Event SearchField(ByVal FormName As String, ByVal FieldName As String, ByVal FieldType As String, ByVal CurValue As String) Implements SalesInterface.MobileSales.ISalesBase.SearchField

    Private Sub btnMemItemNo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMemItemNo.Click
        RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\DataList.dll", "DataList.ItemList", "SalesPromo.SystemSettings", "txtMemItemNo", 0, 0)
    End Sub

    'Localization
    Private Sub Localization()
        Try
            cInfo = New CultureInfo(sLang)
            Thread.CurrentThread.CurrentCulture = cInfo
            Thread.CurrentThread.CurrentUICulture = cInfo
            rMgr = New ResourceManager("SalesPromo.SalesPromo", [Assembly].GetExecutingAssembly())
            Me.Text = rMgr.GetString("SystemSettings")
            SysSet_lblGST.Text = rMgr.GetString("SysSet_lblGST")
            SysSet_lblMdtNo.Text = rMgr.GetString("SysSet_lblMdtNo")
            SysSet_lblMax.Text = rMgr.GetString("SysSet_lblMax")
            SysSet_lblCompName.Text = rMgr.GetString("SysSet_lblCompName")
            SysSet_lblHeader1.Text = rMgr.GetString("SysSet_lblHeader1") & "1 . . . . . ."
            SysSet_lblHeader2.Text = rMgr.GetString("SysSet_lblHeader1") & "2 . . . . . ."
            SysSet_lblHeader3.Text = rMgr.GetString("SysSet_lblHeader1") & "3 . . . . . ."
            SysSet_lblHeader4.Text = rMgr.GetString("SysSet_lblHeader1") & "4 . . . . . ."

            SysSet_lblFooter1.Text = rMgr.GetString("SystemSettings") & "1 . . . . . ."
            SysSet_lblFooter2.Text = rMgr.GetString("SystemSettings") & "2 . . . . . ."
            SysSet_lblFooter3.Text = rMgr.GetString("SystemSettings") & "3 . . . . . ."
            SysSet_lblFooter4.Text = rMgr.GetString("SystemSettings") & "4 . . . . . ."
            SysSet_lblHistUnit.Text = rMgr.GetString("SysSet_lblHistUnit")
            SysSet_lblHistNo.Text = rMgr.GetString("SysSet_lblHistNo")
            SysSet_lblBankAcnt.Text = rMgr.GetString("SysSet_lblBankAcnt")
            SysSet_lblMemItemNo.Text = rMgr.GetString("SysSet_lblMemItemNo")
            btnSave.Text = rMgr.GetString("btnSave")
            btnCancel.Text = rMgr.GetString("btnCancel")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub SetCulture(ByVal CultureName As String) Implements SalesInterface.MobileSales.ISalesBase.SetCulture
        'sLang = CultureName
        'Localization()
    End Sub

    Private Sub btnPrGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrGroup.Click
        RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\DataList.dll", "DataList.PriceGroupList", "SalesPromo.SystemSettings", "txtPrGroup", 0, 0)
    End Sub

    Private Sub btnDefPrGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDefPrGroup.Click
        RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\DataList.dll", "DataList.PriceGroupList", "SalesPromo.SystemSettings", "txtDefPrGroup", 0, 0)
    End Sub

    Private Sub txtMaxDiscount_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMaxDiscount.TextChanged

    End Sub

    Private Sub btnDefTerms_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDefTerms.Click
        RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\DataList.dll", "DataList.PaymentTerms", "SalesPromo.SystemSettings", "txtDefItemTerms", 0, 0)
    End Sub

    Private Sub cmbHisUnit_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbHisUnit.SelectedIndexChanged

    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub LoadSalesType()
        Dim dtr As SqlDataReader
        cmbSalesType.Items.Clear()
        dtr = objDo.ReadRecord("Select Code from SalesType")
        Do While dtr.Read = True
            If dtr("Code") <> "" Then
                cmbSalesType.Items.Add(dtr(0))
            End If
        Loop
        dtr.Close()
        dtr.Dispose()
    End Sub

    Private Sub Label12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

   End Class