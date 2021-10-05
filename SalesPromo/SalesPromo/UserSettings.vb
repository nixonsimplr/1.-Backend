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

Public Class UserSettings
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
    Private strUserCode As String = String.Empty
    Private strUserName As String = String.Empty
    Private strPassword As String = String.Empty
    Private strEnc As String
    Private objDo As New DataInterface.IbizDO
    Dim chkvalue As Integer = 0
    Dim key As New DESCryptoServiceProvider()
    Dim buffer As Byte()
    Private Sub UserSettings_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        objDo.DisconnectDB()
    End Sub
    Private Sub UserSettings_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'SalesDataSet1.SalesAgent' table. You can move, or remove it, as needed.
        Me.SalesAgentTableAdapter.Fill(Me.SalesDataSet1.SalesAgent)
        objDo.ConnectionString = My.Settings.ConnectionString
        objDo.ConnectDB()
        GetUser()
        Me.txtUserCode.Select()
        LoadCommissionCode()
        'IsNewRecord = True
    End Sub
    Private Sub LoadCommissionCode()
        Dim rs As SqlDataReader
        rs = objDo.ReadRecord("SELECT Code FROM CommissionMaster")
        cmbCommCode.Items.Clear()
        While rs.Read
            cmbCommCode.Items.Add(rs("Code").ToString)
        End While
        rs.Close()
    End Sub
    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim iAccess As Integer
        Dim IsUpt As Boolean = False
        Dim dtr1 As SqlDataReader
        dtr1 = objDo.ReadRecord("Select Code from SalesAgent where Code=" & objDo.SafeSQL(txtUserCode.Text))
        If dtr1.Read = True Then
            IsUpt = True
        Else
            IsUpt = False
        End If
        dtr1.Close()
        If Me.txtUserCode.Text <> String.Empty Then
            Me.strUserCode = Me.txtUserCode.Text
        Else
            MessageBox.Show("Please Enter UserCode", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.txtUserCode.Select()
            Return
        End If
        If Me.txtUserName.Text <> String.Empty Then
            Me.strUserName = Me.txtUserName.Text
        Else
            MessageBox.Show("Please Enter UserName", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.txtUserName.Select()
            Return
        End If
        If Me.txtPassword.Text <> String.Empty Then
            Me.strPassword = Me.txtPassword.Text
        Else
            MessageBox.Show("Please Enter Password", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.txtPassword.Select()
            Return
        End If
        If Me.txtRePassword.Text = String.Empty Then
            MessageBox.Show("Please Enter Re-enter Password", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.txtRePassword.Select()
            Return
        End If
        If Me.cmbUserType.Text = "" Then
            MessageBox.Show("Please Select the User Type", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.cmbUserType.Select()
            Return
        End If
        If StrComp(txtPassword.Text, txtRePassword.Text, CompareMethod.Binary) <> 0 Then
            MessageBox.Show("Re-enter Password does not match", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.txtRePassword.Select()
            Return
        End If
        Dim dSTarget As Double
        If txtSalesTarget.Text = "" Then
            dSTarget = 0
        Else
            dSTarget = CDbl(txtSalesTarget.Text)
        End If
        If cmbUserType.Text = "MDT User" Then
            iAccess = 1
        Else
            iAccess = 2
        End If
        strEnc = EncryptText(txtPassword.Text)
        Try
            If IsUpt = True Then
                objDo.ExecuteSQL("Update SalesAgent set Access=" & iAccess & "," & "Name=" & objDo.SafeSQL(txtUserName.Text) & "," & "Password=" & objDo.SafeSQL(txtPassword.Text) & "," & "UserID=" & objDo.SafeSQL(txtSName.Text) & "," & "CommissionCode=" & objDo.SafeSQL(cmbCommCode.Text) & "," & "Active=" & chkvalue & "," & "SalesTarget=" & dSTarget & " where code=" & objDo.SafeSQL(txtUserCode.Text))
                MessageBox.Show("Sales Agent Modified", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                objDo.ExecuteSQL("Insert into SalesAgent(Code, Name, CommissionCode, Password, Access, Active, UserID, SalesTarget, CurMonTarget, LastInvDate) values (" & objDo.SafeSQL(txtUserCode.Text) & "," & objDo.SafeSQL(txtUserName.Text) & "," & objDo.SafeSQL(cmbCommCode.Text) & "," & objDo.SafeSQL(txtPassword.Text) & "," & iAccess & "," & chkvalue & "," & objDo.SafeSQL(txtSName.Text) & "," & dSTarget & ",0," & objDo.SafeSQL(Format(Date.Now, "yyyyMMdd")) & ")")
                MessageBox.Show("Sales Agent Added", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        '  IsNewRecord = False
        'myDataView.FindRows("InvNo")
        'Me.Refresh()
        Clear()
        GetUser()
        RowNo = 0
        LoadRow()
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Clear()
        'Dim dtr As SqlDataReader
        'dtr = objDo.ReadRecord("Select Password from UserSettings where UserName=" & objDo.SafeSQL(txtUserName.Text))
        'While dtr.Read
        '    MsgBox(dtr(0))
        '    txtAccessLevel.Text = DecryptText(dtr(0))
        '    Exit While
        'End While
        'dtr.Close()
    End Sub
    Public Shared Function EncryptText(ByVal strText As String) As String
        Return Encrypt(strText, "~!@#$%*&")
    End Function
    Public Shared Function DecryptText(ByVal strText As String) As String
        Return Decrypt(strText, "~!@#$%*&")
    End Function
    Private Shared Function Encrypt(ByVal strText As String, ByVal strEncrKey _
             As String) As String
        Dim byKey() As Byte = {}
        Dim IV() As Byte = {&H12, &H34, &H56, &H78, &H90, &HAB, &HCD, &HEF}
        Try
            byKey = System.Text.Encoding.UTF8.GetBytes(Microsoft.VisualBasic.Left(strEncrKey, 8))
            Dim des As New DESCryptoServiceProvider()
            Dim inputByteArray() As Byte = Encoding.UTF8.GetBytes(strText)
            Dim ms As New MemoryStream()
            Dim cs As New CryptoStream(ms, des.CreateEncryptor(byKey, IV), _
                   CryptoStreamMode.Write)
            cs.Write(inputByteArray, 0, inputByteArray.Length)
            cs.FlushFinalBlock()
            Return Convert.ToBase64String(ms.ToArray())
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    'The function used to decrypt the text
    Private Shared Function Decrypt(ByVal strText As String, ByVal sDecrKey _
               As String) As String
        Dim byKey() As Byte = {}
        Dim IV() As Byte = {&H12, &H34, &H56, &H78, &H90, &HAB, &HCD, &HEF}
        Dim inputByteArray(strText.Length) As Byte
        Try
            byKey = System.Text.Encoding.UTF8.GetBytes(Microsoft.VisualBasic.Left(sDecrKey, 8))
            Dim des As New DESCryptoServiceProvider()
            inputByteArray = Convert.FromBase64String(strText)
            Dim ms As New MemoryStream()
            Dim cs As New CryptoStream(ms, des.CreateDecryptor(byKey, _
                 IV), CryptoStreamMode.Write)

            cs.Write(inputByteArray, 0, inputByteArray.Length)
            cs.FlushFinalBlock()
            Dim encoding As System.Text.Encoding = System.Text.Encoding.UTF8
            Return encoding.GetString(ms.ToArray())
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Private Sub chkActive_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkActive.CheckedChanged
        If chkActive.Checked = True Then
            chkvalue = 1
        Else
            chkvalue = 0
        End If
    End Sub
    Private Sub GetUser()
        myAdapter = objDo.GetDataAdapter("Select * from SalesAgent")
        myDataSet = New DataSet
        myAdapter.Fill(myDataSet, "SalesAgent")
        myDataView = New DataView(myDataSet.Tables("SalesAgent"))
        If myDataView.Count = 0 Then
            RowNo = -1
        Else
            RowNo = 0
        End If
        LoadRow()
    End Sub
    Private Sub LoadRow()
        If RowNo < 0 Or RowNo >= myDataView.Count Then Exit Sub
        txtUserCode.Text = myDataView(RowNo).Item("Code").ToString
        txtUserName.Text = myDataView(RowNo).Item("Name").ToString
        txtPassword.Text = myDataView(RowNo).Item("Password").ToString
        txtRePassword.Text = myDataView(RowNo).Item("Password").ToString
        txtSalesTarget.Text = myDataView(RowNo).Item("SalesTarget").ToString
        txtCurTarget.Text = myDataView(RowNo).Item("CurMonTarget").ToString
        If CInt(myDataView(RowNo).Item("Access")) = 1 Then
            cmbUserType.SelectedIndex = 0
        Else
            cmbUserType.SelectedIndex = 1
        End If
        txtSName.Text = myDataView(RowNo).Item("UserID").ToString
        chkActive.Checked = myDataView(RowNo).Item("Active")
    End Sub
    Private Sub pnlLogin_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles pnlLogin.Paint
    End Sub

    Public Function GetListViewForm() As String Implements SalesInterface.MobileSales.ISalesBase.GetListViewForm
        RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\SalesAgent.dll", "SalesAgent.DataList", "SalesPromo.UserSettings", "AgentId", 0, 0)
        'Return Windows.Forms.Application.StartupPath & "\SalesAgent.dll,SalesAgent.DataList"
    End Function

    Public Sub ListViewClick() Implements SalesInterface.MobileSales.ISalesBase.ListViewClick

    End Sub

    Public Event LoadForm(ByVal Location As String, ByVal LoadType As String) Implements SalesInterface.MobileSales.ISalesBase.LoadForm

    Public Sub MoveFirstClick() Implements SalesInterface.MobileSales.ISalesBase.MoveFirstClick
        RowNo = 0
        LoadRow()
    End Sub

    Public Sub MoveLastClick() Implements SalesInterface.MobileSales.ISalesBase.MoveLastClick
        RowNo = myDataView.Count - 1
        LoadRow()
    End Sub

    Public Sub MoveNextClick() Implements SalesInterface.MobileSales.ISalesBase.MoveNextClick
        If RowNo >= myDataView.Count - 1 Then Exit Sub
        RowNo = RowNo + 1
        LoadRow()
    End Sub

    Public Sub MovePositionClick(ByVal Position As Long) Implements SalesInterface.MobileSales.ISalesBase.MovePositionClick

    End Sub

    Public Sub MovePreviousClick() Implements SalesInterface.MobileSales.ISalesBase.MovePreviousClick
        If RowNo <= 0 Then Exit Sub
        RowNo = RowNo - 1
        LoadRow()
    End Sub

    Public Sub Print() Implements SalesInterface.MobileSales.ISalesBase.Print

    End Sub

    Public Sub Print(ByVal PageNo As Integer) Implements SalesInterface.MobileSales.ISalesBase.Print

    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub Clear()
        txtUserCode.Select()
        txtUserCode.Text = ""
        txtUserName.Text = ""
        txtPassword.Text = ""
        txtRePassword.Text = ""
        txtSName.Text = ""
        txtSalesTarget.Text = ""
        txtCurTarget.Text = ""
        chkActive.Checked = False
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Public Event LoadDataForm(ByVal Location As String, ByVal LoadType As String, ByVal ParentLoadType As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.LoadDataForm

    Public Event ResultData(ByVal ChildLoadType As String, ByVal Value As String) Implements SalesInterface.MobileSales.ISalesBase.ResultData

    Public Sub ReturnedData(ByVal Value As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.ReturnedData
        If ResultTo = "AgentId" Then
            If myDataSet.Tables.Count < 1 Then Exit Sub
            For idx As Integer = 0 To myDataSet.Tables(0).Rows.Count - 1
                If myDataSet.Tables(0).Rows(idx).Item("Code").ToString.ToUpper = Value.ToUpper Then
                    RowNo = idx
                    Exit For
                End If
            Next
            LoadRow()
            IsNewRecord = False
        End If
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
            Me.Text = rMgr.GetString("UserSettings")
            UserSet_lblPassword.Text = rMgr.GetString("UserSet_lblPassword")
            UserSet_lblRePassword.Text = rMgr.GetString("UserSet_lblRePassword")
            UserSet_lblUserType.Text = rMgr.GetString("UserSet_lblUserType")
            UserSet_lblAccessLevel.Text = rMgr.GetString("UserSet_lblAccessLevel")
            UserSet_lblUserName.Text = rMgr.GetString("UserSet_lblUserName")
            UserSet_lblUserCode.Text = rMgr.GetString("UserSet_lblUserCode")

            chkActive.Text = rMgr.GetString("chkActive")
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

    Private Sub btnNew_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Clear()
        IsNewRecord = True
    End Sub

    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim StrSql As String
        If MsgBox("Do you want to Delete Y/N", MsgBoxStyle.YesNo, "Delete?") = MsgBoxResult.Yes Then
            Try
                StrSql = "Delete from SalesAgent where Code=" & objDo.SafeSQL(txtUserCode.Text)
                objDo.ExecuteSQL(StrSql)
                MessageBox.Show("Sales Agent Deleted")
                '  MessageBox.Show(rMgr.GetString("Message_RecDel"), rMgr.GetString("Message_RecDel_Caption"), MessageBoxButtons.OK)
                '   objDO.ExecuteSQL("Update MDT set LastExNo=" & iExNo & " where MdtNo = " & objDO.SafeSQL(sMDT))
                Clear()
                GetUser()
                RowNo = 0
                LoadRow()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub

    Private Sub txtSalesTarget_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSalesTarget.TextChanged

    End Sub

  End Class