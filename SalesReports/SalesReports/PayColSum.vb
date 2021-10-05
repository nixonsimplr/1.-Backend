Imports System.Data
Imports System.Data.SqlClient
Imports DataInterface.IbizDO
Imports SalesInterface.MobileSales

Public Class PayColSum
    Implements ISalesBase
    Private objDO As New DataInterface.IbizDO
    Dim strPay As String = " "
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim strSql As String = "select SalesAgent.Name as Name, Receipt.AgentID," & objDO.SafeSQL(Format(dtpDate.Value, "dd/MM/yyyy")) & " as Date, count(RcptItem.InvNo) as NoofInv, Sum(RcptItem.AmtPaid) as Amount from SalesAgent,Receipt,RcptItem where SalesAgent.Code=Receipt.AgentID and Receipt.RcptNo=RcptItem.RcptNo and RcptDt Between" & objDO.SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & objDO.SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & "group by AgentID, Name"
        ExecuteReport(strSql, "PayColSumRep")
    End Sub
    Public Sub loadCombo()
        Dim dtr As SqlDataReader
        dtr = objDO.ReadRecord("Select AgentID from MDT")
        cmbAgent.Items.Add("All")
        Do While dtr.Read = True
            cmbAgent.Items.Add(dtr(0))
        Loop
        dtr.Close()
        If cmbAgent.Items.Count > 0 Then
            cmbAgent.SelectedIndex = 0
        End If
        dtpDate.Value = Date.Now
    End Sub

    Private Sub cmbAgent_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAgent.SelectedIndexChanged
        If cmbAgent.Text = "All" Then
            strPay = " "
        Else
            strPay = "and Agentid=" & "'" & cmbAgent.Text & "'"
        End If
    End Sub

    Private Sub PayColSum_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        objDO.DisconnectDB()
    End Sub

    Private Sub PayColSum_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objDO.ConnectDB()
        loadCombo()
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged

    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Public Function GetListViewForm() As String Implements SalesInterface.MobileSales.ISalesBase.GetListViewForm

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