Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Public Class GoodsReturn
    Implements ISalesBase

    Private aAgent As New ArrayList()
    Private aTerms As New ArrayList()
    Dim strPay As String = " "
    Dim strTerms As String
    Private Sub GoodsReturn_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ConnectDB()
        LoadAgent()
        LoadProduct()
    End Sub
    Public Sub LoadAgent()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Distinct Code, Name from SalesAgent order by Name")
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

        'Dim strSql As String
        'If CmbGoodsRetNo.Text = "ALL" Then
        '    strSql = "SELECT GoodsReturn.ReturnNo, GoodsReturn.ReturnDate, GoodsReturn.CustNo, GoodsReturn.SalesPersonCode, GoodsReturn.CreditNoteNo, GoodsReturnItem.UOM, GoodsReturnItem.Quantity, GoodsReturnItem.ItemNo, Customer.CustName, Item.Description FROM GoodsReturn INNER JOIN GoodsReturnItem ON GoodsReturn.ReturnNo = GoodsReturnItem.ReturnNo INNER JOIN Item ON GoodsReturnItem.ItemNo = Item.ItemNo INNER JOIN Customer ON GoodsReturn.CustNo = Customer.CustNo where GoodsReturn.CustNo=" & SafeSQL(cmbCustNo.Text)
        'Else
        '    strSql = "SELECT GoodsReturn.ReturnNo, GoodsReturn.ReturnDate, GoodsReturn.CustNo, GoodsReturn.SalesPersonCode, GoodsReturn.CreditNoteNo, GoodsReturnItem.UOM, GoodsReturnItem.Quantity, GoodsReturnItem.ItemNo, Customer.CustName, Item.Description FROM GoodsReturn INNER JOIN GoodsReturnItem ON GoodsReturn.ReturnNo = GoodsReturnItem.ReturnNo INNER JOIN Item ON GoodsReturnItem.ItemNo = Item.ItemNo INNER JOIN Customer ON GoodsReturn.CustNo = Customer.CustNo where GoodsReturn.ReturnNo = " & SafeSQL(CmbGoodsRetNo.Text)
        'End If
        ''GoodsReturn.CustNo=" & SafeSQL(cmbCustNo.Text) & " and
        'ExecuteReport(strSql, "GoodsRetRep")
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
            strPay = "and GoodsReturn.SalesPersonCode=" & "'" & cmbAgent.SelectedValue & "'"
        End If
        Dim strSql As String = "SELECT GoodsReturn.ReturnNo, GoodsReturn.ReturnDate, GoodsReturn.CustNo, SalesAgent.Name as SalesPersonCode, GoodsReturn.CreditNoteNo, GoodsReturnItem.UOM, GoodsReturnItem.Quantity, GoodsReturnItem.ItemNo, Customer.CustName, Item.Description, GoodsReturn.CompanyName, Customer.CustNo, isnull(Reason.Description,'') as Reason, " & SafeSQL(Format(dtpFromDate.Value, "dd/MM/yyyy")) & " as FromDate, " & SafeSQL(Format(dtpToDate.Value, "dd/MM/yyyy")) & " as ToDate FROM  GoodsReturn INNER JOIN GoodsReturnItem ON GoodsReturn.ReturnNo = GoodsReturnItem.ReturnNo INNER JOIN Item ON Item.CompanyName = GoodsReturn.CompanyName and GoodsReturnItem.ItemNo = Item.ItemNo INNER JOIN Customer ON GoodsReturn.CompanyName = Customer.CompanyName and GoodsReturn.CustNo = Customer.CustNo INNER JOIN SalesAgent ON GoodsReturn.SalesPersonCode = SalesAgent.Code LEFT JOIN Reason on Reason.Code = GoodsReturnItem.Remarks where GoodsReturn.ReturnDate Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & strPay '& " Union SELECT GoodsReturn.ReturnNo, GoodsReturn.ReturnDate, GoodsReturn.CustNo, SalesAgent.Name as SalesPersonCode, GoodsReturn.CreditNoteNo, GoodsReturnItem.UOM, GoodsReturnItem.Quantity, GoodsReturnItem.ItemNo, NewCust.CustName, Item.Description, GoodsReturn.CompanyName, NewCust.CustID, isnull(Reason.Description,'') as Reason FROM  GoodsReturn INNER JOIN GoodsReturnItem ON GoodsReturn.ReturnNo = GoodsReturnItem.ReturnNo INNER JOIN Item ON GoodsReturnItem.ItemNo = Item.ItemNo INNER JOIN NewCust ON GoodsReturn.CustNo = NewCust.CustID INNER JOIN SalesAgent ON GoodsReturn.SalesPersonCode = SalesAgent.Code LEFT JOIN Reason ON Reason.Code = GoodsReturnItem.Remarks where GoodsReturn.ReturnDate Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay
        ExecuteReport(strSql, "GoodsRetRep")
        btnPrint.Enabled = True
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