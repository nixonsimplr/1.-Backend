Imports System
Imports System.Resources
Imports System.Threading
Imports System.Reflection
Imports System.Globalization
Imports DataInterface.IbizDO
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Public Class Commission
    Implements ISalesBase


    Dim bStart As Boolean = False
    Private objDO As New DataInterface.IbizDO
    Dim cCnt, rCnt As Integer
    Private Sub LoadCommissionCode()
        Dim rs As SqlDataReader
        rs = objDO.ReadRecord("SELECT Code FROM CommissionMaster where CommType='Sales Agent'")
        cmbAgent.Items.Clear()
        While rs.Read
            cmbAgent.Items.Add(rs("Code").ToString)
        End While
        rs.Close()
        bStart = True
        LoadGridView()
        cmbAgent.SelectedIndex = 0
    End Sub

    Private Sub LoadGridView()
        Dim i As Integer
        Dim rs As SqlDataReader

        dgvCommission.Rows.Clear()
        dgvCommission.Columns.Clear()

        'add column to gridview
        rs = objDO.ReadRecord("Select Code FROM CommissionMaster where CommType='Customer' order by Code")
        i = 0
        While rs.Read
            dgvCommission.Columns.Add(rs(0).ToString.Trim, rs(0).ToString.Trim)
            'dgvCommission.Columns.Add(i.ToString, i.ToString)
            'dgvCommission.Columns(i).AutoSizeMode() = DataGridViewAutoSizeColumnMode.AllCells
            dgvCommission.Columns(i).FillWeight = 1
            i = i + 1
        End While
        rs.Close()
        cCnt = i

        'add row to gridview
        rs = objDO.ReadRecord("SELECT Count(*) as rowCnt FROM CommissionMaster where CommType='Product'")
        If rs.Read = True Then
            rCnt = rs("rowCnt")
        Else
            rCnt = 0
        End If
        rs.Close()

        If cCnt <= 0 Then
            MsgBox("Please add Customer on Commission Master", MsgBoxStyle.Information)
        ElseIf rCnt > 0 Then
            rs = objDO.ReadRecord("Select Code FROM CommissionMaster where CommType='Product' order by Code")
            dgvCommission.Rows.Add(rCnt)
            For i = 0 To rCnt - 1
                rs.Read()
                dgvCommission.Rows(i).HeaderCell.Value = rs(0).ToString
            Next
            rs.Close()
        End If
    End Sub

    Public Sub LoadGridViewValue()
        Dim i As Integer
        Dim rs As SqlDataReader
        Dim sType As String = ""
        If IsAgentExist(cmbAgent.SelectedItem.ToString, "", "") = True Then
            If rbPrice.Checked = True Then
                sType = "Price"
            ElseIf rbPercentage.Checked = True Then
                sType = "Percentage"
            ElseIf sType = "" Then
                rs = objDO.ReadRecord("Select distinct PriceType from CommissionDet where SalesAgent = " & objDO.SafeSQL(cmbAgent.SelectedItem.ToString))
                If (rs.Read() = True) Then
                    If (rs("PriceType") = "Price") Then
                        rbPrice.Checked = True
                        sType = "Price"
                    Else
                        rbPercentage.Checked = True
                        sType = "Percentage"
                    End If
                End If
                rs.Close()
            End If
            'load data into gridview
            For i = 0 To dgvCommission.RowCount - 1
                For j As Integer = 0 To dgvCommission.ColumnCount - 1
                    dgvCommission.Item(j, i).Value = ""
                Next
            Next
            For i = 0 To dgvCommission.RowCount - 1
                rs = objDO.ReadRecord("Select Customer,Price from CommissionDet where SalesAgent = " & objDO.SafeSQL(cmbAgent.SelectedItem.ToString) & " and Product = " & objDO.SafeSQL(dgvCommission.Rows(i).HeaderCell.Value.ToString.Trim) & " and PriceType = " & objDO.SafeSQL(sType))
                While rs.Read
                    dgvCommission.Item(rs("Customer").ToString, i).Value = rs("Price").ToString
                End While
                rs.Close()
            Next
            dgvCommission.Refresh()
        End If
    End Sub

    Private Sub Commission_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        LoadCommissionCode()
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

    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        SaveCommissionDetail()
    End Sub

    Private Function IsAgentExist(ByVal sSalesAgent As String, ByVal sCustNo As String, ByVal sItemNo As String) As Boolean
        Dim rs As SqlDataReader
        If sCustNo = "" And sItemNo = "" Then
            rs = objDO.ReadRecord("Select SalesAgent FROM CommissionDet where SalesAgent=" & objDO.SafeSQL(sSalesAgent))
        Else
            rs = objDO.ReadRecord("Select SalesAgent FROM CommissionDet where SalesAgent=" & objDO.SafeSQL(sSalesAgent) & " and Customer = " & objDO.SafeSQL(sCustNo) & " and Product = " & objDO.SafeSQL(sItemNo))
        End If

        If rs.Read = True Then
            rs.Close()
            Return True
        End If
        rs.Close()
        Return False
    End Function

    Private Sub SaveCommissionDetail()
        Dim i, j As Integer
        Dim strSql As String
        Dim sPriceType As String
        If rbPrice.Checked = True Then
            sPriceType = "Price"
        ElseIf rbPercentage.Checked = True Then
            sPriceType = "Percentage"
        Else
            MsgBox("Price Type can not be empty")
            Exit Sub
        End If
        Try

            For i = 0 To dgvCommission.RowCount - 1
                For j = 0 To dgvCommission.ColumnCount - 1
                    If IsAgentExist(cmbAgent.SelectedItem.ToString, objDO.SafeSQL(cmbAgent.SelectedItem.ToString.Trim), objDO.SafeSQL(dgvCommission.Rows(i).HeaderCell.Value.ToString.Trim)) = False Then
                        strSql = "Insert into CommissionDet (SalesAgent, Customer, Product, PriceType, Price) values (" & objDO.SafeSQL(cmbAgent.SelectedItem.ToString.Trim) & "," & objDO.SafeSQL(dgvCommission.Columns(j).HeaderText.ToString.Trim) & "," & objDO.SafeSQL(dgvCommission.Rows(i).HeaderCell.Value.ToString.Trim) & "," & objDO.SafeSQL(sPriceType) & "," & CDbl(dgvCommission.Rows(i).Cells(j).Value.ToString) & ")"
                        objDO.ExecuteSQL(strSql)
                    Else
                        strSql = "Update CommissionDet set PriceType=" & objDO.SafeSQL(sPriceType) & ", Price=" & CDbl(dgvCommission.Rows(i).Cells(j).Value.ToString) & " where SalesAgent=" & objDO.SafeSQL(cmbAgent.SelectedItem.ToString.Trim) & " and Customer=" & objDO.SafeSQL(dgvCommission.Columns(j).HeaderText.ToString.Trim) & " and Product=" & objDO.SafeSQL(dgvCommission.Rows(i).HeaderCell.Value.ToString.Trim)
                        objDO.ExecuteSQL(strSql)
                    End If

                Next
            Next
            MsgBox("Commission Detail is updated successfully")

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try
    End Sub




    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        LoadGridViewValue()
    End Sub

 End Class