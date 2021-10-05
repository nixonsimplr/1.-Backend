Imports System.Data
Imports System.Data.SqlClient
Imports DataInterface.IbizDO
Imports SalesInterface.MobileSales

Public Class Commission
    Implements ISalesBase
    Dim ArList As New ArrayList
    Private Structure ArrItemPrice
        Dim ItemCode As String
        Dim Desc As String
    End Structure
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim strAgent As String
        Dim dDel As Double
        Dim rs1 As SqlDataReader
        Dim dtr As SqlDataReader
        Dim str As String
        If cmbAgent.Text = "ALL" Then
            strAgent = " "
        Else
            strAgent = "and Invoice.Agentid=" & "'" & cmbAgent.SelectedValue & "'"
        End If
        Dim strDel = "Delete from CommRep"
        ExecuteSQL(strDel)
        dtr = ReadRecord("SELECT Distinct ItemNo, Description, BaseUOM, Price, CommissionCode FROM Item where Active=1 and ToPDA=1 order by DisplayNo")
        While dtr.Read = True
            Dim aPr1 As ArrItemPrice
            aPr1.ItemCode = dtr("ItemNo").ToString
            aPr1.Desc = dtr("Description").ToString
            ArList.Add(aPr1)
        End While
        dtr.Close()
        Dim i As Integer
        For i = 0 To ArList.Count - 1
            Dim aPr As ArrItemPrice
            aPr = ArList(i)
            str = "SELECT SUM(InvItem.Qty) AS Delivered FROM InvItem INNER JOIN Invoice ON InvItem.InvNo = Invoice.InvNo WHERE (Invoice.Void = 0 and InvItem.ItemNo=" & SafeSQL(aPr.ItemCode) & " and Invoice.AgentID = " & SafeSQL(strAgent) & " and DateDiff(s, " & SafeSQL(Format(dtpStartDate, "yyyyMMdd HH:mm:ss")) & ", InvDt) >= 0 and DateDiff(s, " & SafeSQL(Format(dtpEndDate, "yyyyMMdd HH:mm:ss")) & ", InvDt) < 0 )"
            rs1 = ReadRecord(str)
            If rs1.Read = True Then
                If rs1(0) Is System.DBNull.Value = True Then
                    dDel = 0
                Else
                    dDel = rs1(0)
                End If
            End If
            rs1.Close()
        Next
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
End Class