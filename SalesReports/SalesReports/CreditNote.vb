Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Public Class CreditNote
    Implements ISalesBase
    Private aAgent As New ArrayList()
    Private aTerms As New ArrayList()
    Dim strPay As String = " "
    Dim strTerms As String
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
            strPay = "and CreditNote.SalesPersonCode=" & "'" & cmbAgent.SelectedValue & "'"
        End If
        Dim strSql As String = "SELECT CreditNote.CreditNoteNo, CreditNote.SubTotal as SubAmt, CreditNote.Gst as GSTAmt, CreditNote.TotalAmt as Amount, CreditNote.CreditDate, CreditNote.CustNo, SalesAgent.Name as SalesPersonCode, CreditNote.GoodsReturnNo, CreditNoteDet.UOM, CreditNoteDet.Qty, CreditNoteDet.Price, Amt as LineAmt, CreditNoteDet.ItemNo, Customer.CustName, Item.Description, CreditNoteDet.InvNo FROM  CreditNote INNER JOIN CreditNoteDet ON CreditNote.CreditNoteNo = CreditNoteDet.CreditNoteNo INNER JOIN Item ON CreditNoteDet.ItemNo = Item.ItemNo INNER JOIN Customer ON CreditNote.CustNo = Customer.CustNo INNER JOIN SalesAgent ON CreditNote.SalesPersonCode = SalesAgent.Code where (void=0 or void is null) and CreditNote.CreditDate Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay & " Union SELECT CreditNote.CreditNoteNo, CreditNote.SubTotal as SubAmt, CreditNote.Gst as GSTAmt, CreditNote.TotalAmt as Amount, CreditNote.CreditDate, CreditNote.CustNo, SalesAgent.Name as SalesPersonCode, CreditNote.GoodsReturnNo, CreditNoteDet.UOM, CreditNoteDet.Qty, CreditNoteDet.Price, Amt as LineAmt, CreditNoteDet.ItemNo, NewCust.CustName, Item.Description, '' as InvNo FROM  CreditNote INNER JOIN CreditNoteDet ON CreditNote.CreditNoteNo = CreditNoteDet.CreditNoteNo INNER JOIN Item ON CreditNoteDet.ItemNo = Item.ItemNo INNER JOIN NewCust ON CreditNote.CustNo = NewCust.CustID INNER JOIN SalesAgent ON CreditNote.SalesPersonCode = SalesAgent.Code where (void=0 or void is null) and CreditNote.CreditDate Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strPay
        ExecuteReport(strSql, "CreditNoteRep")
        btnPrint.Enabled = True
    End Sub

    Private Sub CreditNote_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ConnectDB()
        LoadAgent()
        '  LoadProduct()
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