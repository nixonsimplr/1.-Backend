Imports System
Imports System.Resources
Imports System.Threading
Imports System.Reflection
Imports System.Globalization

Imports System.Data.SqlClient

Public Class Password

#Region "Localization Variants"
    Private rMgr As ResourceManager
    Private cInfo As CultureInfo
#End Region

    Private objDO As New DataInterface.IbizDO

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SServ_Cngpass_btnOK.Click
        Dim sActPswd As String = ""
        Dim dtr As SqlDataReader
        dtr = objDO.ReadRecord("Select * from MDT where MdtNo = 'Admin'")
        If dtr.Read() Then
            Dim rs As SqlDataReader
            rs = objDO.ReadRecordAnother("Select * from SalesAgent where Code = " & objDO.SafeSQL(dtr("AgentId")))
            If rs.Read() Then
                sActPswd = rs("Password")
            End If
            rs.Close()
            rs.Dispose()
        End If
        dtr.Close()
        dtr.Dispose()
        If txtPassword.Text = sActPswd Then
            Me.DialogResult = Windows.Forms.DialogResult.OK
        Else
            lblError.Visible = True
        End If
    End Sub

    Private Function IsPassValid(ByVal sUserName As String, ByVal sPass As String) As Boolean
        Dim rs As SqlDataReader
        rs = objDO.ReadRecord("Select code,password,access from SalesAgent where code = " & objDO.SafeSQL(sUserName) & " and Password=" & objDO.SafeSQL(sPass))
        IsPassValid = rs.Read
        rs.Close()
        Return IsPassValid
    End Function

    Private Sub ChangePassword_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Localization()
    End Sub

    Private Sub txtPassword_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPassword.TextChanged
        lblError.Visible = False
    End Sub

    Private Sub SServ_Cngpass_btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SServ_Cngpass_btnCancel.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
    End Sub

    Private Sub txtPassword_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPassword.KeyPress
        If e.KeyChar = Chr(13) Then
            btnOK_Click(Nothing, Nothing)
        End If
    End Sub

End Class