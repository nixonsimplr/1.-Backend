Imports System.Data.SqlClient
Imports SalesInterface.MobileSales

Public Class GoodsExchange
    Implements ISalesBase
    Private aAgent As New ArrayList()
    Private aTerms As New ArrayList()
    Dim strPay As String = " "
    Dim strTerms As String
    Private Sub GoodsExchange_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ConnectDB()
        LoadAgent()
        LoadProduct()
    End Sub
    Public Sub LoadAgent()
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
    Public Sub LoadProduct()
        'Dim dtr As SqlDataReader
        'dtr = ReadRecord("Select Description from PayTerms")
        'Do While dtr.Read = True
        '    If dtr("Description") <> "" Then
        '        cmbTerms.Items.Add(dtr(0))
        '    End If
        'Loop
        'dtr.Close()
        'If cmbTerms.Items.Count > 0 Then
        '    cmbTerms.SelectedIndex = 0
        'End If

        'Dim iIndex As Integer = 0
        'Dim iSelIndex As Integer
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select ItemNo, ShortDesc from Item")
        cmbTerms.DataSource = Nothing
        aTerms.Clear()
        aTerms.Add(New ComboValues("ALL", "ALL"))
        While dtr.Read()
            aTerms.Add(New ComboValues(dtr("ItemNo").ToString, dtr("ShortDesc").ToString))

            '    iSelIndex = iIndex
            'End If
            'iIndex = iIndex + 1
        End While
        dtr.Close()
        cmbTerms.DataSource = aTerms
        cmbTerms.DisplayMember = "Desc"
        cmbTerms.ValueMember = "Code"
        cmbTerms.SelectedIndex = 0

    End Sub

    Private Sub btnReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        btnPrint.Enabled = False
        If cmbTerms.Text = "ALL" Then
            strTerms = ""
        Else
            strTerms = " and Item.ItemNo='" & cmbTerms.SelectedValue & "'"
        End If
        If cmbAgent.Text = "ALL" Then
            strPay = " "
        Else
            strPay = "and GoodsExchange.SalesPersonCode=" & "'" & cmbAgent.SelectedValue & "'"
        End If

        Dim strSql As String = "SELECT GoodsExchange.ExchangeNo, GoodsExchange.ExchangeDate, GoodsExchange.CustNo, SalesAgent.Name as SalesPersonCode, GoodsExchangeItem.ItemNo, GoodsExchangeItem.UOM, GoodsExchangeItem.Quantity, Customer.CustName, Item.Description FROM GoodsExchange INNER JOIN GoodsExchangeItem ON GoodsExchange.ExchangeNo = GoodsExchangeItem.ExchangeNo INNER JOIN Item ON GoodsExchangeItem.ItemNo = Item.ItemNo INNER JOIN Customer ON GoodsExchange.CustNo = Customer.CustNo INNER JOIN SalesAgent ON GoodsExchange.SalesPersonCode = SalesAgent.Code where GoodsExchange.ExchangeDate Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & " Union SELECT GoodsExchange.ExchangeNo, GoodsExchange.ExchangeDate, GoodsExchange.CustNo, SalesAgent.Name as SalesPersonCode, GoodsExchangeItem.ItemNo, GoodsExchangeItem.UOM, GoodsExchangeItem.Quantity, NewCust.CustName, Item.Description FROM GoodsExchange INNER JOIN GoodsExchangeItem ON GoodsExchange.ExchangeNo = GoodsExchangeItem.ExchangeNo INNER JOIN Item ON GoodsExchangeItem.ItemNo = Item.ItemNo INNER JOIN NewCust ON GoodsExchange.CustNo = NewCust.CustID INNER JOIN SalesAgent ON GoodsExchange.SalesPersonCode = SalesAgent.Code where GoodsExchange.ExchangeDate Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay
        ExecuteReport(strSql, "GoodsExRep")
        btnPrint.Enabled = True
    End Sub

    Private Sub cmbAgent_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAgent.SelectedIndexChanged

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