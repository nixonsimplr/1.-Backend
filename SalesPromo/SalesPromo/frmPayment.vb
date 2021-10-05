Imports System
Imports System.Resources
Imports System.Globalization
Imports System.Threading
Imports System.Reflection

Imports DataInterface.IbizDO
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales

Public Class frmPayment
    Implements ISalesBase



#Region "Localization Variants"
    Private rMgr As ResourceManager
    Private cInfo As CultureInfo
    Private sLang As String
#End Region

    Private iPayType As Integer = 0
    Private objDO As New DataInterface.IbizDO
    Public frm As Membership

    Private Sub btnPay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub PayMethodTextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub frmPayment_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        objDO.DisconnectDB()
    End Sub

    Private Sub frmPayment_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objDO.ConnectionString = My.Settings.ConnectionString
        objDO.ConnectDB()
        cmbPaymethod.Items.Clear()
        Try
            Dim dtr As SqlDataReader
            dtr = objDO.ReadRecord("Select Code from PayMethod")
            While dtr.Read
                cmbPaymethod.Items.Add(dtr("Code").ToString)
            End While
            dtr.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    'Localization
    Private Sub Localization()
        Try
            cInfo = New CultureInfo(sLang)
            Thread.CurrentThread.CurrentCulture = cInfo
            Thread.CurrentThread.CurrentUICulture = cInfo
            rMgr = New ResourceManager("Customer.Customer", [Assembly].GetExecutingAssembly())

            Me.Text = rMgr.GetString("frmPayment")
            Cust_Paym_lblAmount.Text = rMgr.GetString("Cust_Paym_lblAmount")
            Cust_Paym_lblPayMhd.Text = rMgr.GetString("Cust_DFrm_lblPayMthd")
            Cust_Paym_lblCCNo.Text = rMgr.GetString("Cust_Paym_lblCCNo")
            Cust_Paym_lblExpDat.Text = rMgr.GetString("Col_DataList_ExpiryDate")
            Cust_Paym_lblBankName.Text = rMgr.GetString("Cust_Paym_lblBankName")
            Cust_Paym_gbCCard.Text = rMgr.GetString("Cust_Paym_gbCCard")
            Cust_Paym_btnOK.Text = rMgr.GetString("Cust_Paym_btnOK")
            Cust_Paym_btnCancel.Text = rMgr.GetString("Cust_DFrm_btnCancel")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cust_Paym_btnCancel.Click
        frm.bPayment = False
        Me.Close()
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cust_Paym_btnOK.Click
        frm.PayMethod = cmbPaymethod.Text
        frm.bPayment = True
        If cmbPaymethod.Text = "" Then Exit Sub
        If iPayType = 1 Then
            frm.sChqNo = txtChqNo.Text
            frm.sBank = txtBankName.Text
            frm.dChqDate = dtpChqDate.Value
        ElseIf iPayType = 2 Then
            frm.sChqNo = txtCCardNo.Text
            frm.sBank = txtCBankName.Text
            frm.dChqDate = dtpExpiryDate.Value
        Else
            frm.sChqNo = ""
            frm.sBank = ""
            frm.dChqDate = Date.Now
        End If
        Me.Close()
    End Sub

    Private Sub cmbPaymethod_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbPaymethod.SelectedIndexChanged
        Cust_Paym_gbCCard.Visible = False
        gbCheque.Visible = False
        Try
            Dim rs As SqlDataReader
            rs = objDO.ReadRecord("Select PaymentType from PayMethod where Code=" & objDO.SafeSQL(cmbPaymethod.Text))
            If rs.Read = True Then
                If rs("PaymentType") = 1 Then
                    gbCheque.Visible = True
                ElseIf rs("PaymentType") = 2 Then
                    Cust_Paym_gbCCard.Visible = True
                End If
                iPayType = rs("PaymentType")
            End If
            rs.Close()
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

    End Sub

    Public Sub ReturnSearch(ByVal SQL As String) Implements SalesInterface.MobileSales.ISalesBase.ReturnSearch

    End Sub

    Public Event SearchField(ByVal FormName As String, ByVal FieldName As String, ByVal FieldType As String, ByVal CurValue As String) Implements SalesInterface.MobileSales.ISalesBase.SearchField

    Public Sub SetCulture(ByVal CultureName As String) Implements SalesInterface.MobileSales.ISalesBase.SetCulture
        sLang = CultureName
        Localization()
    End Sub
End Class