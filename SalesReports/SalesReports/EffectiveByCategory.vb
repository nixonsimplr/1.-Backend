Imports System.Data
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Public Class EffectiveByCategory
    Implements ISalesBase
    Private aCat As New ArrayList()
    Private aAgent As New ArrayList()
    Dim strAgent As String = " "
    Dim strCat As String
    Dim ArList As New ArrayList
    Private Structure ArrCust
        Dim sItemNo As String
        Dim sDesc As String
        Dim sItemName As String
        Dim sUOM As String
        Dim sPrice As Double
        Dim sCPCode As String
    End Structure

    Private Sub EffectiveByCategory_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub

    Private Sub EffectiveByCategory_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ConnectDB()
        'ExecuteSQL("Drop table InvoiceList")
        'ExecuteSQL("SELECT Customer.CustNo, Customer.CustName, Customer.PaymentMethod, Invoice.InvNo, Invoice.InvDt, Invoice.AgentID, Invoice.Discount, Invoice.SubTotal, Invoice.GST, Invoice.TotalAmt, Invoice.PayTerms into InvoiceList FROM Customer INNER JOIN Invoice ON Customer.CustNo = Invoice.CustId")
        '        LoadCategory()
        '       LoadAgent()
    End Sub
    Public Sub LoadCategory()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Code, Description from Category")
        cmbCategory.DataSource = Nothing
        aCat.Clear()
        aCat.Add(New ComboValues("ALL", "ALL"))
        While dtr.Read()
            aCat.Add(New ComboValues(dtr("Code").ToString, dtr("Description").ToString))
        End While
        dtr.Close()
        cmbCategory.DataSource = aCat
        cmbCategory.DisplayMember = "Desc"
        cmbCategory.ValueMember = "Code"
        cmbCategory.SelectedIndex = 0

    End Sub
    Public Sub LoadAgent()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Code, Name from SalesAgent")
        cmbAgent.DataSource = Nothing
        aAgent.Clear()
        'aAgent.Add(New ComboValues("ALL", "ALL"))
        While dtr.Read()
            aAgent.Add(New ComboValues(dtr("Code").ToString, dtr("Name").ToString))
        End While
        dtr.Close()
        cmbAgent.DataSource = aAgent
        cmbAgent.DisplayMember = "Desc"
        cmbAgent.ValueMember = "Code"
        cmbAgent.SelectedIndex = 0
    End Sub
    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        '  btnPrint.Enabled = False
        'If cmbAgent.Text = "ALL" Then
        '    strAgent = " "
        'Else
        '    strAgent = "and SalesAgent.Code=" & "'" & cmbAgent.SelectedValue & "'"
        'End If
        'If cmbCategory.Text = "ALL" Then
        '    strCat = ""
        'Else
        '    strCat = " and Category.Code='" & cmbCategory.SelectedValue & "'"
        'End If
        Dim aPr1 As ArrCust
        aPr1.sItemNo = "P30F"
        aPr1.sItemName = "P30F"
        aPr1.sDesc = "P30F Packed Eggs F4 30s"
        aPr1.sPrice = 1.5
        aPr1.sCPCode = ""
        aPr1.sUOM = "TRAY"
        ArList.Add(aPr1)
        'Dim strDel = "Delete from EffectiveByItem"
        '        ExecuteSQL(strDel)
        'ArList.Clear()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select CustNo from Customer where ToPDA=1 and PriceGroup in ('SS')")
        While dtr.Read = True
            ExecuteAnotherSQL("Insert into CustProd(CustID,ItemNo, Description, ItemName, UOM, Price, CustProdCode) Values (" & SafeSQL(dtr("CustNo").ToString) & "," & SafeSQL(aPr1.sItemNo) & "," & SafeSQL(aPr1.sDesc) & "," & SafeSQL(aPr1.sItemName) & "," & SafeSQL(aPr1.sUOM) & "," & aPr1.sPrice & "," & SafeSQL(aPr1.sCPCode) & ")")
        End While
        dtr.Close()
        MsgBox("Comp")
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