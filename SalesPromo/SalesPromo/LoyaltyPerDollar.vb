Imports DataInterface.IbizDO
Imports SalesInterface.MobileSales
Imports System.Data.SqlClient

Public Class LoyaltyPerDollar
    Implements ISalesBase


    Private objDO As New DataInterface.IbizDO
    Private IsNewRecord As Boolean
    Private rInd As Integer
    Private Sub SaveLoyalty()
        Try
            If txtCode.Text = String.Empty Then
                MsgBox("Sorry! Please select MemberType")
                Me.txtCode.Select()
                Return
            End If
            If txtPoints.Text = String.Empty Then
                MsgBox("Sorry! Please Enter Loyalty Points")
                Me.txtPoints.Select()
                Return
            End If
            objDO.ExecuteSQL("Delete from LoyaltyPerDollar where MemberType=" & objDO.SafeSQL(txtCode.Text))
            Dim strSql As String
            strSql = "Insert into LoyaltyPerDollar(MemberType, Points, DTG) Values (" & objDO.SafeSQL(txtCode.Text) & "," & txtPoints.Text & "," & objDO.SafeSQL(Format(DateTime.Now, "yyyyMMdd HH:mm:ss")) & ")"
            objDO.ExecuteSQL(strSql)
            MessageBox.Show("Loyalty Per Dollar Modified", "New Record", MessageBoxButtons.OK)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Public Sub ReturnedData(ByVal Value As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.ReturnedData
        If ResultTo = "txtCode" Then
            txtCode.Text = Value
        End If
    End Sub

    Private Sub txtCode_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCode.TextChanged
        Try
            Dim sCode As String = ""
            Dim rs As SqlDataReader
            sCode = "Select Membertype, Points from LoyaltyPerDollar where MemberType=" & objDO.SafeSQL(txtCode.Text)
            rs = objDO.ReadRecord(sCode)
            If rs.Read = True Then
                txtPoints.Text = rs("Points")
            End If
            rs.Close()
            IsNewRecord = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        SaveLoyalty()
        IsNewRecord = False
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

    Public Sub ReturnSearch(ByVal SQL As String) Implements SalesInterface.MobileSales.ISalesBase.ReturnSearch

    End Sub

    Public Event SearchField(ByVal FormName As String, ByVal FieldName As String, ByVal FieldType As String, ByVal CurValue As String) Implements SalesInterface.MobileSales.ISalesBase.SearchField

    Public Sub SetCulture(ByVal CultureName As String) Implements SalesInterface.MobileSales.ISalesBase.SetCulture

    End Sub

    Private Sub LoyaltyRedemptionPerDollar_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnCustNo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCustNo.Click
        clear()
        RaiseEvent LoadDataForm(Windows.Forms.Application.StartupPath & "\DataList.dll", "DataList.MemberType", "SalesPromo.LoyaltyPerDollar", "txtCode", 0, 0)
    End Sub
    Private Sub clear()
        txtCode.Text = ""
        txtPoints.Text = "0"
    End Sub

    Public Event LoadDataForm(ByVal Location As String, ByVal LoadType As String, ByVal ParentLoadType As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer) Implements SalesInterface.MobileSales.ISalesBase.LoadDataForm

    Public Event LoadForm(ByVal Location As String, ByVal LoadType As String) Implements SalesInterface.MobileSales.ISalesBase.LoadForm
End Class