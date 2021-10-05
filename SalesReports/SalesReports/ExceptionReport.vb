Imports System.Data
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Public Class ExceptionReport
    Implements ISalesBase
    Private aAgent As New ArrayList()
    Dim strAgent, strType As String
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

    Private Sub ExceptionReport_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        DisconnectDB()
    End Sub

    Private Sub ExceptionReport_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ConnectDB()
        LoadSalesAgent()
        LoadDocType()
        dtpFromDate.Value = Date.Now
        dtpToDate.Value = Date.Now
    End Sub
    Public Sub LoadSalesAgent()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Code, Name from SalesAgent")
        cmbAgent.DataSource = Nothing
        aAgent.Clear()
        aAgent.Add(New ComboValues("ALL", "ALL"))
        While dtr.Read()
            aAgent.Add(New ComboValues(dtr("Code").ToString, dtr("Name").ToString))

            '    iSelIndex = iIndex
            'End If
            'iIndex = iIndex + 1
        End While
        dtr.Close()
        cmbAgent.DataSource = aAgent
        cmbAgent.DisplayMember = "Desc"
        cmbAgent.ValueMember = "Code"
        cmbAgent.SelectedIndex = 0
    End Sub

    Public Sub LoadDocType()
        cmbType.Items.Clear()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Distinct DocType From Exception order by DocType")
        cmbType.Items.Add("ALL")
        Do While dtr.Read = True
            If dtr("DocType") <> "" Then
                cmbType.Items.Add(dtr(0))
            End If
        Loop
        dtr.Close()
        If cmbType.Items.Count > 0 Then
            cmbType.SelectedIndex = 0
        End If
    End Sub

    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        If cmbAgent.Text = "ALL" Then
            strAgent = " "
        Else
            strAgent = " and Exception.Agentid=" & "'" & cmbAgent.SelectedValue & "'"
        End If
        If cmbType.Text = "ALL" Then
            strType = " "
        Else
            strType = " and Exception.DocType=" & "'" & cmbType.Text & "'"
        End If
        Dim strSql As String

        If cmbAgent.Text = "ALL" Then
            strSql = "SELECT Distinct CustID, DocNo, DocType, ItemId, ColType, DocDate, Remarks, '" & Format(dtpFromDate.Value, "dd/MM/yyyy") & "' as  FromDate, '" & Format(dtpToDate.Value, "dd/MM/yyyy") & "' as ToDate, 'ALL' as AgentID FROM  Exception where DocDate Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & strAgent & strType
        Else
            strSql = "SELECT Distinct CustID, DocNo, DocType, ItemId, ColType, DocDate, Remarks, '" & Format(dtpFromDate.Value, "dd/MM/yyyy") & "' as  FromDate, '" & Format(dtpToDate.Value, "dd/MM/yyyy") & "' as ToDate, AgentID  + ' / ' + SalesAgent.Name as AgentID FROM  Exception, SalesAgent where Exception.AgentID=SalesAgent.Code and DocDate Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & strAgent & strType
        End If
        ExecuteReport(strSql, "ExceptionRep")
    End Sub
End Class