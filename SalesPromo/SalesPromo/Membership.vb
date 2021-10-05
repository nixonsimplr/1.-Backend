Imports System
Imports System.Resources
Imports System.Globalization
Imports System.Threading
Imports System.Reflection
Imports System.Data.SqlClient
Imports DataInterface.IbizDO
Imports SalesInterface.MobileSales

Public Class Membership
    Implements ISalesBase


#Region "Localization Variants"
    Private rMgr As ResourceManager
    Private cInfo As CultureInfo
    Private sLang As String
#End Region
    Private objDO As New DataInterface.IbizDO
    Private myAdapter As SqlDataAdapter
    Private myDataSet As DataSet
    Private myDataView As DataView
    Private sMDT As String = ""
    Private iMemNo As Integer
    Private IsNewRecord As Boolean
    Private IsModify As Boolean = True
    Private sItemNo As String = ""
    Private sInvNo As String = ""
    Private sRcptNo As String = ""
    Public bPayment As Boolean = False
    Public sChqNo As String = ""
    Public sBank As String = ""
    Public dChqDate As Date
    Public PayMethod As String = ""
    '=====================================================================================
    'Search Function
    Public Sub ReturnSearch(ByVal SQL As String) Implements SalesInterface.MobileSales.ISalesBase.ReturnSearch
        If SQL = "" Then
            LoadData("SELECT * FROM Customer where Active=1 and CustType='Member' Order by CustNo")
        Else
            LoadData("SELECT * FROM Customer where Active=1 and CustType='Member' and " & SQL & " Order by CustNo")
        End If
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
    'End Search
    '=====================================================================================

    Private Sub LoadData(ByVal sSQL As String)
        myAdapter = objDO.GetDataAdapter(sSQL)
        myDataSet = New DataSet
        myAdapter.Fill(myDataSet, "Member")
        myDataView = New DataView(myDataSet.Tables("Member"))
        ' MemberBindingSource.DataSource = myDataView
    End Sub

    Private Sub DataForm_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        objDO.DisconnectDB()
    End Sub

    Private Sub DataForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'TODO: This line of code loads data into the 'MemberDataSet.Member' table. You can move, or remove it, as needed.
        'Me.MemberTableAdapter.Fill(Me.MemberDataSet.Member)
        objDO.ConnectionString = My.Settings.ConnectionString
        objDO.ConnectDB()
        txtExpiryDate.ReadOnly = True

        LoadData("SELECT * FROM Customer where Active=1 and CustType='Member' Order by CustNo")
        GetTextBox()
    End Sub

    'Localization
    Private Sub Localization()
        Try
            cInfo = New CultureInfo(sLang)
            Thread.CurrentThread.CurrentCulture = cInfo
            Thread.CurrentThread.CurrentUICulture = cInfo
            rMgr = New ResourceManager("Customer.Customer", [Assembly].GetExecutingAssembly())

            Me.Text = rMgr.GetString("Membership")
            Cust_Memb_lblMemNo.Text = rMgr.GetString("Cust_Memb_lblMemNo")
            Cust_Memb_lblMName.Text = rMgr.GetString("Cust_Memb_lblMName")
            Cust_Memb_lblAddr.Text = rMgr.GetString("Cust_DFrm_lblAddress") & ". . . . . . . . . . ."
            Cust_Memb_lblAddr2.Text = rMgr.GetString("Cust_DFrm_lblAddress") & " 2. . . . . . . . . ."
            Cust_Memb_lblAddr3.Text = rMgr.GetString("Cust_DFrm_lblAddress") & " 3. . . . . . . . . ."

            Cust_Memb_lblAddr4.Text = rMgr.GetString("Cust_DFrm_lblAddress") & " 4. . . . . . . . . ."
            Cust_Memb_lblPtCd.Text = rMgr.GetString("Cust_DFrm_lblPost")
            Cust_Memb_lblCntCd.Text = rMgr.GetString("Cust_DFrm_lblCountry")
            Cust_Memb_lblPhNo.Text = rMgr.GetString("Cust_Memb_lblPhNo")
            Cust_Memb_lblWeb.Text = rMgr.GetString("Col_DataList_WebSite") & ". . . . . . . . . . ."
            Cust_Memb_lblExpDt.Text = rMgr.GetString("Col_DataList_ExpiryDate") & ". . . . . . . . . . ."
            Cust_Memb_lblMemTp.Text = rMgr.GetString("Col_DataList_MemType") & ". . . . . . . . . . ."
            Cust_Memb_lblChNm.Text = rMgr.GetString("Cust_DFrm_lblChinese")
            Cust_Memb_lblBaln.Text = rMgr.GetString("Cust_DFrm_lblBalance")
            Cust_Memb_lblCrLtm.Text = rMgr.GetString("Cust_DFrm_lblCredLmt")
            Cust_Memb_lblCatg.Text = rMgr.GetString("Col_DataList_Category") & ". . . . . . . . . . ."
            Cust_Memb_lblPhoto.Text = rMgr.GetString("Col_DataList_Photo") & ". . . . . . . . . . ."

            Cust_Memb_btnCancel.Text = rMgr.GetString("Cust_DFrm_btnCancel")
            Cust_Memb_btnDelete.Text = rMgr.GetString("Cust_DFrm_btnDelete")
            Cust_Memb_btnNew.Text = rMgr.GetString("Cust_DFrm_btnNew")
            Cust_Memb_btnRenew.Text = rMgr.GetString("Cust_Memb_btnRenew")
            Cust_Memb_btnSave.Text = rMgr.GetString("Cust_DFrm_btnSave")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub ListViewClick() Implements SalesInterface.MobileSales.ISalesBase.ListViewClick
        RaiseEvent LoadForm(Windows.Forms.Application.StartupPath & "\Customer.dll", "Customer.MembershipList")
    End Sub

    Public Sub MoveFirstClick() Implements SalesInterface.MobileSales.ISalesBase.MoveFirstClick
        '   MemberBindingSource.MoveFirst()
    End Sub

    Public Sub MoveLastClick() Implements SalesInterface.MobileSales.ISalesBase.MoveLastClick
        '  MemberBindingSource.MoveLast()
    End Sub

    Public Sub MoveNextClick() Implements SalesInterface.MobileSales.ISalesBase.MoveNextClick
        ' MemberBindingSource.MoveNext()
    End Sub

    Public Sub MovePositionClick(ByVal Position As Long) Implements SalesInterface.MobileSales.ISalesBase.MovePositionClick

    End Sub

    Public Sub MovePreviousClick() Implements SalesInterface.MobileSales.ISalesBase.MovePreviousClick
        'MemberBindingSource.MovePrevious()
    End Sub

    Public Sub Print() Implements SalesInterface.MobileSales.ISalesBase.Print

    End Sub

    Public Sub Print(ByVal PageNo As Integer) Implements SalesInterface.MobileSales.ISalesBase.Print

    End Sub

    Public Event LoadForm(ByVal Location As String, ByVal LoadType As String) Implements SalesInterface.MobileSales.ISalesBase.LoadForm

    Public Function GetListViewForm() As String Implements SalesInterface.MobileSales.ISalesBase.GetListViewForm
        RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\Customer.dll", "Customer.MembershipList", "Customer.Membership", "txtMemberNo", 0, 0)
        Return Windows.Forms.Application.StartupPath & "\Customer.dll,Customer.MembershipList"
    End Function

    Private Sub Clear()
        txtMemberNo.Text = String.Empty
        txtName.Text = String.Empty
        txtChinese.Text = String.Empty
        txtAdd.Text = String.Empty
        txtAdd2.Text = String.Empty
        txtAdd3.Text = String.Empty
        txtAdd4.Text = String.Empty
        txtPost.Text = String.Empty
        txtCity.Text = String.Empty
        txtCCode.Text = String.Empty
        txtCountry.Text = String.Empty
        txtPhone.Text = String.Empty
        txtEmail.Text = String.Empty
        txtBalance.Text = String.Empty
        txtCL.Text = String.Empty
        txtCategory.Text = String.Empty
        txtMobileNo.Text = String.Empty
        txtMemberType.Text = String.Empty
        txtCategory.Text = String.Empty
        txtEmail.Text = String.Empty
        txtWebsite.Text = String.Empty
        txtPhoto.Text = String.Empty
        txtFaxNo.Text = String.Empty
        txtExpiryDate.Text = ""
        txtName.Select()
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cust_Memb_btnNew.Click
        NewCustomer()
    End Sub

    Private Sub NewCustomer()
        Dim sMemNo As String = ""
        Dim rs As SqlDataReader

        Clear()
        rs = objDO.ReadRecord("Select MDTNo from System")
        If rs.Read = True Then
            sMDT = rs("MDTNo")
        End If
        rs.Close()

        rs = objDO.ReadRecord("Select * from MDT where MdtNo = " & objDO.SafeSQL(sMDT) & "")
        If rs.Read = True Then
            Dim sPre As String = rs("PreCustNo")
            Dim iLen As Int32 = rs("LenCustNo")
            iMemNo = rs("LastCustNo") + 1
            sMemNo = sPre & StrDup(iLen - Len(sPre) - Len(CStr(iMemNo)), "0") & CStr(iMemNo)
        End If
        rs.Close()

        IsNewRecord = True
        txtMemberNo.Text = sMemNo

        Dim dtr As SqlDataReader
        Dim sPeriodType As String = ""
        Dim sPayTerms As String = ""
        Dim sPeriod As Integer

        dtr = objDO.ReadRecord("Select Period, PeriodType, PayTerms from MemberType where Code=" & objDO.SafeSQL(txtMemberType.Text))
        If dtr.Read = True Then
            sPeriodType = dtr("PeriodType").ToString
            sPeriod = CInt(dtr("Period"))
            sPayTerms = dtr("PayTerms").ToString
        End If
        dtr.Close()

        If sPeriodType = "Year" Then
            txtExpiryDate.Text = Format(DateAdd(DateInterval.Year, sPeriod, Date.Now), "dd-MMM-yyyy")
        End If
        If sPeriodType = "Month" Then
            txtExpiryDate.Text = Format(DateAdd(DateInterval.Month, sPeriod, Date.Now), "dd-MMM-yyyy")
        End If
        If sPeriodType = "Day" Then
            txtExpiryDate.Text = Format(DateAdd(DateInterval.Day, sPeriod, Date.Now), "dd-MMM-yyyy")
        End If
    End Sub

    Private Sub SaveCustomer()
        'Localization
        If txtMemberNo.Text = String.Empty Then
            MsgBox(rMgr.GetString("Message_Empty_MemNo"))
            Me.txtMemberNo.Select()
            Return
        End If

        If txtName.Text = String.Empty Then
            MsgBox(rMgr.GetString("Message_Empty_MemName"))
            Me.txtName.Select()
            Return
        End If

        Dim strSql As String
        Try
            If txtBalance.Text = String.Empty Then txtBalance.Text = CDbl("0")
            If txtCategory.Text = String.Empty Then txtCategory.Text = CDbl("0")
            If txtCL.Text = String.Empty Then txtCL.Text = CDbl("0")
            If txtPhoto.Text = String.Empty Then txtPhoto.Text = Application.StartupPath & "\blank.jpg"
            strSql = "Insert into Customer(CustNo, CustName, ChineseName, Address, Address2, Address3, Address4, PostCode, City, CountryCode, Phone, Balance, CreditLimit, FaxNo, Email, Website, category, MemberType, Photo, Active, CustType, Contactperson, ProvisionalBalance, ZoneCode, ICPartner, PriceGroup, PaymentTerms, PaymentMethod, SalesAgent, ShipAgent, Dimension1, Dimension2, InvDisGroup, Location, CurrencyCode, CustPostGroup, ExpiryDate, Exported, DTG, MDTNo) Values (" & objDO.SafeSQL(txtMemberNo.Text) & ", " & objDO.SafeSQL(txtName.Text) & ", " & objDO.SafeSQL(txtChinese.Text) & ", " & objDO.SafeSQL(txtAdd.Text) & ", " & objDO.SafeSQL(txtAdd2.Text) & ", " & objDO.SafeSQL(txtAdd3.Text) & ", " & objDO.SafeSQL(txtAdd4.Text) & ", " & objDO.SafeSQL(txtPost.Text) & ", " & objDO.SafeSQL(txtCity.Text) & ", " & objDO.SafeSQL(txtCCode.Text) & ", " & objDO.SafeSQL(txtPhone.Text) & ", " & txtBalance.Text & ", " & txtCL.Text & ", " & objDO.SafeSQL(txtFaxNo.Text) & ", " & objDO.SafeSQL(txtEmail.Text) & ", " & objDO.SafeSQL(txtWebsite.Text) & ", " & objDO.SafeSQL(txtCategory.Text) & ", " & objDO.SafeSQL(txtMemberType.Text) & ", " & objDO.SafeSQL(txtPhoto.Text) & ",1,'Member','',0,'','','','','','','','','','','','',''," & objDO.SafeSQL(Format(CDate(txtExpiryDate.Text), "dd-MMM-yyyy")) & ",0," & objDO.SafeSQL(Format(Date.Now, "yyyyMMdd 00:00:00")).ToString & "," & objDO.SafeSQL(sMDT) & ")"
            objDO.ExecuteSQL(strSql)
            objDO.ExecuteSQL("Update MDT set LastCustNo=" & iMemNo & " where MdtNo = " & objDO.SafeSQL(sMDT))
            Clear()
            IsNewRecord = False
            LoadData("SELECT * FROM Customer where Active=1 and CustType='Member' Order by CustNo")
            '     MemberBindingSource.MoveFirst()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub ModifyCustomer()
        Dim strsql As String
        Try
            'Localization
            If txtMemberNo.Text = String.Empty Then
                MsgBox(rMgr.GetString("Message_Empty_Mfy_MemNo"), MsgBoxStyle.OkOnly, "Information")
                Exit Sub
            End If

            If txtBalance.Text = String.Empty Then txtBalance.Text = CDbl("0")
            If txtCategory.Text = String.Empty Then txtCategory.Text = CDbl("0")
            If txtCL.Text = String.Empty Then txtCL.Text = CDbl("0")
            If txtPhoto.Text = String.Empty Then txtPhoto.Text = Application.StartupPath & "\blank.jpg"
            strsql = "Update Customer set CustName=" & objDO.SafeSQL(txtName.Text).ToString & " , ChineseName=" & objDO.SafeSQL(txtChinese.Text).ToString & ", Address=" & objDO.SafeSQL(txtAdd.Text).ToString & ", Address2=" & objDO.SafeSQL(txtAdd2.Text).ToString & ", Address3=" & objDO.SafeSQL(txtAdd3.Text).ToString & ", Address4=" & objDO.SafeSQL(txtAdd4.Text).ToString & ", PostCode=" & objDO.SafeSQL(txtPost.Text).ToString & ", City=" & objDO.SafeSQL(txtCity.Text).ToString & ", CountryCode=" & objDO.SafeSQL(txtCountry.Text).ToString & ", Phone=" & objDO.SafeSQL(txtPhone.Text).ToString & ", Balance=" & txtBalance.Text & ", CreditLimit=" & txtCL.Text & ", FaxNo=" & objDO.SafeSQL(txtMobileNo.Text).ToString & ", Email=" & objDO.SafeSQL(txtEmail.Text).ToString & ", Website=" & objDO.SafeSQL(txtWebsite.Text).ToString & ", Category=" & objDO.SafeSQL(txtCategory.Text).ToString & ", MemberType=" & objDO.SafeSQL(txtMemberType.Text).ToString & ",ExpiryDate=" & objDO.SafeSQL(Format(CDate(txtExpiryDate.Text), "dd-MMM-yyyy")) & ", Photo=" & objDO.SafeSQL(txtPhoto.Text).ToString & ", Active=1, Exported=0 where CustNo = " & objDO.SafeSQL(txtMemberNo.Text)
            objDO.ExecuteSQL(strsql)

            MessageBox.Show(rMgr.GetString("Message_MemModified"), rMgr.GetString("Message_MemMdfy_Caption"), MessageBoxButtons.OK)
            LoadData("SELECT * FROM Customer where Active=1 and CustType='Member' Order by CustNo")
            ' MemberBindingSource.MoveFirst()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub


    Private Sub btnSave_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cust_Memb_btnSave.Click
        If IsNewRecord = True Then
            Dim dtr As SqlDataReader
            Dim sDesc As String = ""
            Dim sPayTerms As String = ""
            Dim dPrice As Double = 0
            dtr = objDO.ReadRecord("Select * from Item, System where System.MembershipItemNo=Item.ItemNo")
            If dtr.Read = True Then
                sItemNo = dtr("ItemNo").ToString
                sDesc = dtr("Description").ToString
                dPrice = dtr("UnitPrice")
            End If
            dtr.Close()
            Dim frm As New frmPayment
            frm.frm = Me
            frm.txtAmt.ReadOnly = True
            frm.txtAmt.Text = Format(dPrice, "0.00")
            frm.ShowDialog()
            If bPayment = False Then Exit Sub
            SaveCustomer()
            Dim rs As SqlDataReader
            rs = objDO.ReadRecord("Select MDTNo from System")
            If rs.Read = True Then
                sMDT = rs("MDTNo")
            End If
            rs.Close()
            rs = objDO.ReadRecord("Select * from MDT where MdtNo = " & objDO.SafeSQL(sMDT) & "")
            If rs.Read = True Then
                Dim sPre As String = rs("PreInvNo")
                Dim iLen As Int32 = rs("LenInvNo")
                iMemNo = rs("LastInvNo") + 1
                sInvNo = sPre & StrDup(iLen - Len(sPre) - Len(CStr(iMemNo)), "0") & CStr(iMemNo)
                sPre = rs("PreRcptNo")
                iLen = rs("LenRcptNo")
                iMemNo = rs("LastRcptNo") + 1
                sRcptNo = sPre & StrDup(iLen - Len(sPre) - Len(CStr(iMemNo)), "0") & CStr(iMemNo)
            End If
            rs.Close()
            objDO.ExecuteSQL("Insert into Invoice (InvNo, InvDt, DoNO, DoDT, OrdNo, CustId, PONo, AgentId, Discount, SubTotal, GSTAmt, TotalAmt, PayTerms, CurCode, CurExRate, PaidAmt, DTG, MDTNo ) Values (" & objDO.SafeSQL(sInvNo) & ",'" & Format(Date.Now, "dd-MMM-yyyy") & "'," & objDO.SafeSQL(sInvNo) & ",'" & Format(Date.Now, "yyyyMMdd 00:00:00") & "', " & objDO.SafeSQL("") & ", " & objDO.SafeSQL(txtMemberNo.Text) & ", " & objDO.SafeSQL("") & ", " & objDO.SafeSQL("") & ",0, 0, 0, " & dPrice & ", " & objDO.SafeSQL(sPayTerms) & ", " & objDO.SafeSQL("") & ",0,0" & "," & objDO.SafeSQL(Format(Date.Now, "yyyyMMdd 00:00:00")).ToString & "," & objDO.SafeSQL(sMDT) & ")")
            objDO.ExecuteSQL("Insert into InvItem (InvNo, ItemNo, Description, Uom, Qty, Price, SubAmt, DeliQty,Location, [LineNo]) Values (" & objDO.SafeSQL(sInvNo) & ", " & objDO.SafeSQL(sItemNo) & ", " & objDO.SafeSQL(sDesc) & ",'CARD' ,1, " & dPrice & "," & dPrice & ", 1 , '',1)")
            objDO.ExecuteSQL("Update MDT set LastInvNo = LastInvNo + 1, LastRcptNo = LastRcptNo + 1 where MDTNo = " & objDO.SafeSQL(sMDT))
            objDO.ExecuteSQL("Insert into Receipt (RcptNo, RcptDt, CustId, AgentId, PayMethod, ChqNo, Amount, CurCode, CurExRate, chqDt, BankName, Exported, Void, DTG, MDTNo) Values (" & objDO.SafeSQL(sRcptNo) & ",'" & Format(Date.Now, "yyyyMMdd 00:00:00") & "'," & objDO.SafeSQL(txtMemberNo.Text) & ", " & objDO.SafeSQL("") & ", " & objDO.SafeSQL(PayMethod) & ", " & objDO.SafeSQL(sChqNo) & ", " & dPrice.ToString & ", " & objDO.SafeSQL("") & ", 1," & objDO.SafeSQL(dChqDate) & "," & objDO.SafeSQL(sBank) & ",0,0" & "," & objDO.SafeSQL(Format(Date.Now, "yyyyMMdd 00:00:00")).ToString & "," & objDO.SafeSQL(sMDT) & ")")
            objDO.ExecuteSQL("Insert into RcptItem (RcptNo, InvNo, AmtPaid, ActAmt) Values (" & objDO.SafeSQL(sRcptNo) & ", " & objDO.SafeSQL(sInvNo) & ", " & dPrice.ToString & ", " & dPrice.ToString & ")")
            objDO.ExecuteSQL("Update Invoice set PaidAmt = TotalAmt where InvNo = " & objDO.SafeSQL(sInvNo))
            LoadReport(sInvNo)

            'Localization
            MessageBox.Show(rMgr.GetString("Message_MemAdded"), rMgr.GetString("Message_MemAdd_Caption"), MessageBoxButtons.OK)
        Else
            ModifyCustomer()
        End If

        ' IsModify = False
        'MemberBindingSource.MoveFirst()
    End Sub

    Private Sub btnDelete_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cust_Memb_btnDelete.Click
        Dim StrSql As String
        Try
            StrSql = "Delete from Customer where CustNo=" & objDO.SafeSQL(txtMemberNo.Text)
            objDO.ExecuteSQL(StrSql)

            'Localization
            MessageBox.Show(rMgr.GetString("Message_MemDel"), rMgr.GetString("Message_MemDel_Caption"), MessageBoxButtons.OK)
            '   objDO.ExecuteSQL("Update MDT set LastExNo=" & iExNo & " where MdtNo = " & objDO.SafeSQL(sMDT))
            Clear()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        '  MemberBindingSource.MoveFirst()
    End Sub

    Private Sub btnCancel_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cust_Memb_btnCancel.Click
        IsNewRecord = False
        Clear()
        myAdapter = objDO.GetDataAdapter("SELECT * FROM Customer where Active=1 and CustType='Member' order by CustNo")
        myDataSet = New DataSet
        myAdapter.Fill(myDataSet, "Member")
        myDataView = New DataView(myDataSet.Tables("Member"))
        '    MemberBindingSource.DataSource = myDataView
        '   MemberBindingSource.MoveFirst()
        'btnNew.Enabled = True
        'btnSave.Enabled = False
    End Sub

    Private Sub txtNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If IsNewRecord = True Then Exit Sub
        myAdapter = objDO.GetDataAdapter("SELECT * FROM Customer where Active=1 and CustType='Member' Order by CustNo")
        myDataSet = New DataSet
        myAdapter.Fill(myDataSet, "Member")
        myDataView = New DataView(myDataSet.Tables("Member"))
        'MemberBindingSource.DataSource = myDataView
    End Sub

    Public Event LoadDataForm(ByVal Location As String, ByVal LoadType As String, ByVal ParentLoadType As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.LoadDataForm

    Public Event ResultData(ByVal ChildLoadType As String, ByVal Value As String) Implements SalesInterface.MobileSales.ISalesBase.ResultData

    Public Sub ReturnedData(ByVal Value As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.ReturnedData
        If ResultTo = "txtMemberNo" Then
            LoadData("SELECT * FROM Customer where Active=1 and CustType='Member' and CustNo='" & Value & "'")
            GetTextBox()
        End If
        'If ResultTo = "txtMemberNo" Then
        'txtMemberNo.Text = Value
        If ResultTo = "txtMemberType" Then
            txtMemberType.Text = Value
        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPhoto.TextChanged
        'pbMemberPhoto.ImageLocation(txtPhoto.Text)
        If txtPhoto.Text = "" Then
            Exit Sub
        Else
            pbMemberPhoto.Image = New System.Drawing.Bitmap(txtPhoto.Text)
        End If
    End Sub

    Private Sub btnPromo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
        With OFDPhoto
            .Filter = "Image Files (*.jpg)|*.jpg|" & "All files|*.*"
            If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                txtPhoto.Text = .FileName
            End If
        End With
    End Sub

    Private Sub btnMemberType_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMemberType.Click
        RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\DataList.dll", "DataList.MemberType", "Customer.Membership", "txtMemberType", 0, 0)
    End Sub
    Private Sub btnRenew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cust_Memb_btnRenew.Click
        'Clear()
        Dim sPeriodType As String = ""
        Dim sDesc As String = ""
        Dim sPeriod As Integer
        Dim sPayTerms As String = ""
        Dim dPrice As Double = 0
        Dim IsCon As Boolean = False

        If txtMemberNo.Text = String.Empty Then
            'Localization
            MsgBox(rMgr.GetString("Message_Empty_MemRenew"))
            Me.txtMemberNo.Select()
            Return
        End If

        Dim dtr As SqlDataReader
        dtr = objDO.ReadRecord("Select ExpiryDate from Customer where CustNo=" & objDO.SafeSQL(txtMemberNo.Text))
        If dtr.Read = True Then
            If MessageBox.Show(rMgr.GetString("Msg_Member1") & DateDiff(DateInterval.Day, Date.Now, dtr("ExpiryDate")) & rMgr.GetString("Msg_Member2") & vbCrLf & rMgr.GetString("Msg_Member3"), rMgr.GetString("Msg_Member4"), MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                IsCon = True
            End If
        End If
        dtr.Close()
        If IsCon = True Then
            dtr = objDO.ReadRecord("Select * from Item, System where System.MembershipItemNo=Item.ItemNo")
            If dtr.Read = True Then
                sItemNo = dtr("ItemNo").ToString
                sDesc = dtr("Description").ToString
                dPrice = dtr("UnitPrice")
            End If
            dtr.Close()
            Dim frm As New frmPayment
            frm.frm = Me
            frm.txtAmt.ReadOnly = True
            frm.txtAmt.Text = Format(dPrice, "0.00")
            frm.ShowDialog()
            If bPayment = False Then Exit Sub

            dtr = objDO.ReadRecord("Select Period, PeriodType, PayTerms from MemberType where Code=" & objDO.SafeSQL(txtMemberType.Text))
            If dtr.Read = True Then
                sPeriodType = dtr("PeriodType").ToString
                sPeriod = CInt(dtr("Period"))
                sPayTerms = dtr("PayTerms").ToString
            End If
            dtr.Close()
            If sPeriodType = "Year" Then
                txtExpiryDate.Text = Format(DateAdd(DateInterval.Year, sPeriod, Date.Now), "dd-MMM-yyyy")
            End If
            If sPeriodType = "Month" Then
                txtExpiryDate.Text = Format(DateAdd(DateInterval.Month, sPeriod, Date.Now), "dd-MMM-yyyy")
            End If
            If sPeriodType = "Day" Then
                txtExpiryDate.Text = Format(DateAdd(DateInterval.Day, sPeriod, Date.Now), "dd-MMM-yyyy")
            End If
            objDO.ExecuteSQL("Update Customer set ExpiryDate=" & objDO.SafeSQL(CDate(txtExpiryDate.Text)) & " where CustNo = " & objDO.SafeSQL(txtMemberNo.Text))
            Dim rs As SqlDataReader
            rs = objDO.ReadRecord("Select MDTNo from System")
            If rs.Read = True Then
                sMDT = rs("MDTNo")
            End If
            rs.Close()

            rs = objDO.ReadRecord("Select * from MDT where MdtNo = " & objDO.SafeSQL(sMDT) & "")
            If rs.Read = True Then
                Dim sPre As String = rs("PreInvNo")
                Dim iLen As Int32 = rs("LenInvNo")
                iMemNo = rs("LastInvNo") + 1
                sInvNo = sPre & StrDup(iLen - Len(sPre) - Len(CStr(iMemNo)), "0") & CStr(iMemNo)
                sPre = rs("PreRcptNo")
                iLen = rs("LenRcptNo")
                iMemNo = rs("LastRcptNo") + 1
                sRcptNo = sPre & StrDup(iLen - Len(sPre) - Len(CStr(iMemNo)), "0") & CStr(iMemNo)
            End If
            rs.Close()
            objDO.ExecuteSQL("Insert into Invoice (InvNo, InvDt, DoNO, DoDT, OrdNo, CustId, PONo, AgentId, Discount, SubTotal, GSTAmt, TotalAmt, PayTerms, CurCode, CurExRate, PaidAmt) Values (" & objDO.SafeSQL(sInvNo) & ",'" & Format(Date.Now, "dd-MMM-yyyy") & "'," & objDO.SafeSQL(sInvNo) & ",'" & Format(Date.Now, "dd-MMM-yyyy") & "', " & objDO.SafeSQL("") & ", " & objDO.SafeSQL(txtMemberNo.Text) & ", " & objDO.SafeSQL("") & ", " & objDO.SafeSQL("") & ",0, 0, 0, " & dPrice & ", " & objDO.SafeSQL(sPayTerms) & ", " & objDO.SafeSQL("") & ",0,0)")
            objDO.ExecuteSQL("Insert into InvItem (InvNo, ItemNo, Description, Uom, Qty, Price, SubAmt, DeliQty,Location, [LineNo]) Values (" & objDO.SafeSQL(sInvNo) & ", " & objDO.SafeSQL(sItemNo) & ", " & objDO.SafeSQL(sDesc) & ",'CARD' ,1, " & dPrice & "," & dPrice & ", 1 , '',1)")
            objDO.ExecuteSQL("Update MDT set LastInvNo = LastInvNo + 1, LastRcptNo = LastRcptNo + 1 where MDTNo = " & objDO.SafeSQL(sMDT))
            objDO.ExecuteSQL("Insert into Receipt (RcptNo, RcptDt, CustId, AgentId, PayMethod, ChqNo, Amount, CurCode, CurExRate, chqDt, BankName, Exported, Void) Values (" & objDO.SafeSQL(sRcptNo) & ",'" & Format(Date.Now, "dd-MMM-yyyy") & "'," & objDO.SafeSQL(txtMemberNo.Text) & ", " & objDO.SafeSQL("") & ", " & objDO.SafeSQL(PayMethod) & ", " & objDO.SafeSQL(sChqNo) & ", " & dPrice.ToString & ", " & objDO.SafeSQL("") & ", 1," & objDO.SafeSQL(dChqDate) & "," & objDO.SafeSQL(sBank) & ",0,0)")
            objDO.ExecuteSQL("Insert into RcptItem (RcptNo, InvNo, AmtPaid, ActAmt) Values (" & objDO.SafeSQL(sRcptNo) & ", " & objDO.SafeSQL(sInvNo) & ", " & dPrice.ToString & ", " & dPrice.ToString & ")")
            objDO.ExecuteSQL("Update Invoice set PaidAmt = TotalAmt where InvNo = " & objDO.SafeSQL(sInvNo))
            LoadReport(sInvNo)
            MessageBox.Show(rMgr.GetString("Message_MemRenewed"), rMgr.GetString("Message_MemRenewed_Info"), MessageBoxButtons.OK)
        Else
            MessageBox.Show(rMgr.GetString("Message_MemNotRenewed"), rMgr.GetString("Message_MemNotRenewed_Info"), MessageBoxButtons.OK)
        End If
    End Sub

    Public Sub LoadReport(ByVal InvNo As String)
        objDO.ExecuteSQL("Delete from InvReport")
        Dim strSql As String = "SELECT Item.Description, InvItem.[LineNo] as Line, InvItem.ItemNo, Customer.CustNo, Customer.CustName, Customer.Address, Customer.Address2, Customer.Address3, Customer.Address4, Customer.Phone, Customer.ContactPerson, Customer.FaxNo, Invoice.DoNo, Invoice.InvNo AS Expr1, Invoice.InvDt, Invoice.PoNo, Invoice.AgentId, Invoice.CustRefNo, Invoice.PayTerms, InvItem.Qty, InvItem.Price, InvItem.SubAmt, Invoice.SubTotal, Invoice.GST, Invoice.GstAmt, Invoice.TotalAmt, Item.AssemblyBOM FROM Customer INNER JOIN Invoice ON Customer.CustNo = Invoice.CustId INNER JOIN InvItem ON Invoice.InvNo = InvItem.InvNo INNER JOIN Item ON InvItem.ItemNo = Item.ItemNo and Invoice.InvNo=" & objDO.SafeSQL(InvNo)
        Dim rs As SqlDataReader
        rs = objDO.ReadRecord(strSql)
        While rs.Read = True
            objDO.ExecuteSQLAnother("Insert into InvReport(InvNo, InvDate, CustNo, AgentID, CustName, Add1, Add2, Add3, Add4, Phone, Fax, Attn, DoNo, RefNo, PONo, Terms, Line, ItemNo, Description, Qty, price, SubAmt, SubTotal, GSTAmt, TotalAmt) values (" & objDO.SafeSQL(rs("Expr1").ToString) & "," & objDO.SafeSQL(Format(CDate(rs("InvDt")), "dd-MMM-yyyy")) & "," & objDO.SafeSQL(rs("CustNo").ToString) & "," & objDO.SafeSQL(rs("AgentID").ToString) & "," & objDO.SafeSQL(rs("CustName").ToString) & "," & objDO.SafeSQL(rs("Address").ToString) & "," & objDO.SafeSQL(rs("Address2").ToString) & "," & objDO.SafeSQL(rs("Address3").ToString) & "," & objDO.SafeSQL(rs("Address4").ToString) & "," & objDO.SafeSQL(rs("Phone").ToString) & "," & objDO.SafeSQL(rs("FaxNo").ToString) & "," & objDO.SafeSQL(rs("ContactPerson").ToString) & "," & objDO.SafeSQL(rs("DONo").ToString) & "," & objDO.SafeSQL(rs("CustRefNo").ToString) & "," & objDO.SafeSQL(rs("PONo").ToString) & "," & objDO.SafeSQL(rs("PayTerms").ToString) & "," & rs("Line").ToString & "," & objDO.SafeSQL(rs("ItemNo").ToString) & "," & objDO.SafeSQL(rs("Description").ToString) & "," & objDO.SafeSQL(rs("Qty").ToString) & "," & objDO.SafeSQL(Format(rs("Price"), "0.00")) & "," & objDO.SafeSQL(Format(rs("SubAmt"), "0.00")) & "," & objDO.SafeSQL(Format(rs("SubTotal"), "0.00")) & "," & objDO.SafeSQL(Format(rs("GstAmt"), "0.00")) & "," & objDO.SafeSQL(Format(rs("TotalAmt"), "0.00")) & ")")
            If CBool(rs("AssemblyBOM")) = True Then
                Dim rsAss As SqlDataReader
                Dim sSpace As String = "  "
                rsAss = objDO.ReadRecordAnother("Select * from AssemblyBOM where ParentItemNo = " & objDO.SafeSQL(rs("ItemNo")))
                While rsAss.Read
                    objDO.ExecuteSQLAnother("Insert into InvReport(InvNo, InvDate, CustNo, AgentID, CustName, Add1, Add2, Add3, Add4, Phone, Fax, Attn, DoNo, RefNo, PONo, Terms, ItemNo, Description, Qty, price, SubAmt, SubTotal, GSTAmt, TotalAmt) values (" & objDO.SafeSQL(rs("Expr1").ToString) & "," & objDO.SafeSQL(Format(CDate(rs("InvDt")), "dd-MMM-yyyy")) & "," & objDO.SafeSQL(rs("CustNo").ToString) & "," & objDO.SafeSQL(rs("AgentID").ToString) & "," & objDO.SafeSQL(rs("CustName").ToString) & "," & objDO.SafeSQL(rs("Address").ToString) & "," & objDO.SafeSQL(rs("Address2").ToString) & "," & objDO.SafeSQL(rs("Address3").ToString) & "," & objDO.SafeSQL(rs("Address4").ToString) & "," & objDO.SafeSQL(rs("Phone").ToString) & "," & objDO.SafeSQL(rs("FaxNo").ToString) & "," & objDO.SafeSQL(rs("ContactPerson").ToString) & "," & objDO.SafeSQL(rs("DONo").ToString) & "," & objDO.SafeSQL(rs("CustRefNo").ToString) & "," & objDO.SafeSQL(rs("PONo").ToString) & "," & objDO.SafeSQL(rs("PayTerms").ToString) & "," & objDO.SafeSQL("") & "," & objDO.SafeSQL(sSpace & rsAss("Qty").ToString & " x " & rsAss("Description").ToString) & ", '', '', '', '', '', '')")
                End While
                rsAss.Close()
            End If
        End While
        rs.Close()
        'ExecuteReport("Select * from InvReport", "EverHomeInvoiceRep")
        strSql = "Select * from InvReport"
        ' Dim strSql As String = "SELECT Invoice.InvNo, Invoice.InvDt, Customer.CustNo, Customer.CustName, Invoice.PayTerms, Invoice.AgentId, InvItem.ItemNo, Item.Description, InvItem.UOM, InvItem.Qty, InvItem.Price, InvItem.SubAmt, InvItem.GstAmt, Invoice.TotalAmt FROM  Item INNER JOIN InvItem ON Item.ItemNo = InvItem.ItemNo INNER JOIN Invoice ON InvItem.InvNo = Invoice.InvNo INNER JOIN Customer ON Invoice.CustId = Customer.CustNo where Invoice.InvDt Between " & objDO.SafeSql(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & objDO.SafeSql(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strTerms & strPay
        'ExecuteReport(strSql, "DailySalesRep")
        Dim RptName As String = "EverHomeInvoiceRep"
        Dim DA As New SqlDataAdapter(strSql, My.Settings.ConnectionString)
        Dim DS As New DataSet
        DA.Fill(DS)
        ' Dim ConInfo As New CrystalDecisions.Shared.TableLogOnInfo
        ' ConInfo.ConnectionInfo.IntegratedSecurity = True
        'ConInfo.ConnectionInfo.UserID = objDataBase.UserName
        'ConInfo.ConnectionInfo.Password = objDataBase.Password
        'ConInfo.ConnectionInfo.ServerName = ".\SQLEXPRESS"
        'ConInfo.ConnectionInfo.DatabaseName = "Sales"
        Dim strReportPath As String = Application.StartupPath & "\" & RptName & ".rpt"
        If Not IO.File.Exists(strReportPath) Then
            Throw (New Exception("Unable to locate report file:" & vbCrLf & strReportPath))
        End If
        'Dim rptDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        'rptDocument.Load(strReportPath)
        ' rptDocument.SetDataSource(DS.Tables(0))
        Dim frm As New ViewReport
        'frm.crvReport.ShowRefreshButton = False
        'frm.crvReport.ShowCloseButton = False
        'frm.crvReport.ShowGroupTreeButton = False
        'frm.crvReport.ReportSource = rptDocument
        'frm.Show()
    End Sub

    Private Sub txtMemberType_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMemberType.TextChanged
        Dim dtr As SqlDataReader
        Dim sPeriodType As String = ""
        Dim sDesc As String = ""
        Dim sPeriod As Integer
        Dim sPayTerms As String = ""
        If txtMemberType.Text = "" Then Exit Sub
        dtr = objDO.ReadRecord("Select Period, PeriodType, PayTerms from MemberType where Code=" & objDO.SafeSQL(txtMemberType.Text))
        If dtr.Read = True Then
            sPeriodType = dtr("PeriodType").ToString
            sPeriod = CInt(dtr("Period"))
            sPayTerms = dtr("PayTerms").ToString
        End If
        dtr.Close()
        If sPeriodType = "Year" Then
            txtExpiryDate.Text = Format(DateAdd(DateInterval.Year, sPeriod, Date.Now), "dd-MMM-yyyy")
        End If
        If sPeriodType = "Month" Then
            txtExpiryDate.Text = Format(DateAdd(DateInterval.Month, sPeriod, Date.Now), "dd-MMM-yyyy")
        End If
        If sPeriodType = "Day" Then
            txtExpiryDate.Text = Format(DateAdd(DateInterval.Day, sPeriod, Date.Now), "dd-MMM-yyyy")
        End If
    End Sub

    Public Sub SetCulture(ByVal CultureName As String) Implements SalesInterface.MobileSales.ISalesBase.SetCulture
        sLang = CultureName
        Localization()
    End Sub

    Private Sub txtMemberNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMemberNo.TextChanged

    End Sub

 End Class
