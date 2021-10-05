Imports System.Data
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Public Class SalesAnalysis
    Implements ISalesBase
    Private aCat As New ArrayList()
    Private aAgent As New ArrayList()
    Dim strAgent As String = " "
    Dim strCat As String
    Dim ArCustList As New ArrayList
    Dim ArCatList As New ArrayList
    Private Structure ArrCust
        Dim sCustNo As String
        Dim sCustName As String
        Dim sAgent As String
    End Structure
    Private Structure ArrCat
        Dim sCode As String
        Dim sDay As Integer
        ' Dim sAgent As String
    End Structure
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        btnPrint.Enabled = False
        If cmbAgent.Text = "ALL" Then
            strAgent = " "
        Else
            strAgent = "and SalesAgent.Code=" & "'" & cmbAgent.SelectedValue & "'"
        End If
        If cmbCategory.Text = "ALL" Then
            strCat = ""
        Else
            strCat = " and Category.Code='" & cmbCategory.SelectedValue & "'"
        End If
        Dim strDel = "Delete from SalesAnalysis"
        ExecuteSQL(strDel)
        ArCustList.Clear()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("SELECT Customer.CustNo, Customer.CustName, SalesAgent.Name FROM  Customer, SalesAgent Where SalesAgent.Code = Customer.SalesAgent " & strAgent)
        While dtr.Read = True
            Dim aPr1 As ArrCust
            aPr1.sCustNo = dtr("CustNo").ToString
            aPr1.sCustName = dtr("CustName").ToString
            aPr1.sAgent = dtr("Name").ToString
            ArCustList.Add(aPr1)
        End While
        dtr.Close()

        Dim i, j As Integer
      
        Dim icnt As Integer = DateDiff(DateInterval.Day, dtpFromDate.Value, dtpToDate.Value)
        
        For i = 0 To ArCustList.Count - 1
            Dim aPr As ArrCust
            aPr = ArCustList(i)
            For j = 1 To icnt
                ExecuteSQL("Insert into SalesAnalysis(CustNo, CustName, AgentID, EDay, Effective, FromDate, ToDate, Category, Remarks) values(" & SafeSQL(aPr.sCustNo) & "," & SafeSQL(aPr.sCustName) & "," & SafeSQL(aPr.sAgent) & "," & j & ",0," & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(cmbCategory.Text) & ",'')")
            Next j
        Next i
        Dim str As String
        Dim rs1 As SqlDataReader

        For i = 0 To ArCustList.Count - 1
            Dim aPr1 As ArrCust
            aPr1 = ArCustList(i)
            str = "SELECT Distinct Category.Code FROM Invoice INNER JOIN InvItem ON Invoice.InvNo = InvItem.InvNo INNER JOIN Item ON InvItem.ItemNo = Item.ItemNo INNER JOIN Category ON Item.Category = Category.Code where Invoice.CustID=" & SafeSQL(aPr1.sCustNo) & " and InvDt Between " & SafeSQL(Format(dtpFromDate.Value.AddDays(-30), "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value.AddDays(-30), "yyyyMMdd 23:59:59")) '& strCat
            rs1 = ReadRecordAnother(str)
            ArCatList.Clear()
            While rs1.Read = True
                Dim aC1 As ArrCat
                aC1.sCode = rs1("Code").ToString
                '         aC1.sDay = rs1("Day")
                ArCatList.Add(aC1)
            End While
            rs1.Close()

            For j = 0 To ArCatList.Count - 1
                Dim aCat As ArrCat
                aCat = ArCatList(j)
                str = "SELECT Category.Code, DAy(Invoice.InvDt) as Day FROM Invoice INNER JOIN InvItem ON Invoice.InvNo = InvItem.InvNo INNER JOIN Item ON InvItem.ItemNo = Item.ItemNo INNER JOIN Category ON Item.Category = Category.Code where Invoice.CustID=" & SafeSQL(aPr1.sCustNo) & " and InvDt Between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) '& strCat
                rs1 = ReadRecord(str)
                'Dim iCnt As Integer = 0
                While rs1.Read = True
                    If aCat.sCode <> rs1("Code").ToString Then
                        ExecuteAnotherSQL("Update SalesAnalysis Set Effective =1 Where CustNo=" & SafeSQL(aPr1.sCustNo) & " and EDay=" & rs1("Day"))
                    End If
                End While
                rs1.Close()
                '  ExecuteSQL("Insert into EffectiveByItem(CustNo, CustName, AgentID, Qty, Effective, EffDate, Category, Remarks) values(" & SafeSQL(aPr.sCustNo) & "," & SafeSQL(aPr.sCustName) & "," & SafeSQL(aPr.sAgent) & "," & dQty & "," & dEff & "," & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd HH:mm:ss")) & "," & SafeSQL(cmbCategory.Text) & ",'')")
            Next
        Next
        Dim strSql As String = "Select *  from SalesAnalysis"
        ExecuteReport(strSql, "SalesAnalysisRep")
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

    Private Sub SalesAnalysis_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub

    Private Sub SalesAnalysis_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        '        MsgBox(DateDiff(DateInterval.Day, Date.Now.AddDays(-30), Date.Now))
        ConnectDB()
        'ExecuteSQL("Drop table InvoiceList")
        'ExecuteSQL("SELECT Customer.CustNo, Customer.CustName, Customer.PaymentMethod, Invoice.InvNo, Invoice.InvDt, Invoice.AgentID, Invoice.Discount, Invoice.SubTotal, Invoice.GST, Invoice.TotalAmt, Invoice.PayTerms into InvoiceList FROM Customer INNER JOIN Invoice ON Customer.CustNo = Invoice.CustId")
        LoadCategory()
        LoadAgent()
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
End Class