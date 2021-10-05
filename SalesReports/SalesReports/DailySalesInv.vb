Imports System.Data
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Public Class DailySalesInv
    Implements ISalesBase
    Private aCust As New ArrayList()
    Private aAgent As New ArrayList()
    Dim strPay As String = " "
    Dim strTerms As String
    
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        If rbQty.Checked = True Then
            LoadReportByQty()
        Else
            LoadReportByAmt()
        End If
    End Sub

    Private Sub LoadReportByQty()
        Dim dtr As SqlDataReader
        Dim dRet As Double
        Dim dQty As Double
        Dim dFoc As Double
        Dim dEx As Double
        Dim str As String = ""
        Dim sItemNo As String = ""
        Dim strAgent As String = ""
        Dim strAgentRet As String = ""
        Dim strCust As String = ""
        Dim strCustRet As String = ""
        If cmbAgent.Text = "ALL" Then
            strAgent = " "
            strAgentRet = " "
        Else
            strAgent = " AND Invoice.AgentId = " & SafeSQL(cmbAgent.SelectedValue.ToString.Trim)
            strAgentRet = " AND GoodsReturn.SalesPersonCode = " & SafeSQL(cmbAgent.SelectedValue.ToString.Trim)
        End If
        If cmbCust.Text = "ALL" Then
            strCust = " "
            strCustRet = " "
        Else
            strCust = " AND Invoice.CustId = " & SafeSQL(cmbCust.SelectedValue.ToString.Trim)
            strCustRet = " AND GoodsReturn.CustNo = " & SafeSQL(cmbCust.SelectedValue.ToString.Trim)
        End If
        ExecuteSQL("Delete from DailySalesInv")
        str = "SELECT Distinct Customer.CustName, Invoice.CustId, InvItem.ItemNo, Item.ShortDesc as ShortDesc, Item.DisplayNo FROM Customer, Invoice, Item, InvItem Where Customer.CustNo = Invoice.CustId and Invoice.InvNo = InvItem.InvNo and InvItem.ItemNo = Item.ItemNo and Void=0 and InvDT Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strAgent & strCust & " Union " & _
                "SELECT Distinct Customer.CustName, GoodsReturn.CustNo as CustId, GoodsReturnItem.ItemNo, Item.ShortDesc as ShortDesc, Item.DisplayNo FROM Customer, GoodsReturn, Item, GoodsReturnItem Where Customer.CustNo = GoodsReturn.CustNo and GoodsReturn.ReturnNo = GoodsReturnItem.ReturnNo and GoodsReturnItem.ItemNo = Item.ItemNo and ReturnDate Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strAgentRet & strCustRet & " Order by DisplayNo"
        dtr = ReadRecord(str)
        While dtr.Read
            sItemNo = dtr("ItemNo").ToString
            dQty = 0
            dRet = 0
            dEx = 0
            dFoc = 0
            Dim dtrTemp As SqlDataReader
            dtrTemp = ReadRecordAnother("Select isnull(Sum(Qty * UOM.BaseQty),0) as Qty from InvItem, Invoice,UOM Where InvItem.ItemNo = UOM.ItemNo and InvItem.Uom = UOM.Uom and InvItem.InvNo = Invoice.InvNo and (InvItem.Description not like 'EX%' or InvItem.Description is null) and InvItem.ItemNo = " & SafeSQL(sItemNo) & " and CustId = " & SafeSQL(dtr("CustID").ToString) & strAgent & " and (Void=0 or void is null) and InvDT Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")))
            If dtrTemp.Read Then
                If dtrTemp("Qty") Is System.DBNull.Value <> True Then
                    dQty = dtrTemp("Qty")
                End If
            End If
            dtrTemp.Close()
            dtrTemp = ReadRecordAnother("Select isnull(Sum(Qty * UOM.BaseQty),0) as Qty from InvItem, Invoice,UOM Where InvItem.ItemNo = UOM.ItemNo and InvItem.Uom = UOM.Uom and InvItem.InvNo = Invoice.InvNo and InvItem.Description like 'EX%' and InvItem.ItemNo = " & SafeSQL(sItemNo) & " and CustId = " & SafeSQL(dtr("CustID").ToString) & strAgent & " and (Void=0 or void is null) and InvDT Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")))
            If dtrTemp.Read Then
                If dtrTemp("Qty") Is System.DBNull.Value <> True Then
                    dEx = dtrTemp("Qty")
                End If
            End If
            dtrTemp.Close()
            dtrTemp = ReadRecordAnother("Select isnull(Sum(Qty * UOM.BaseQty),0) as Qty from InvItem, Invoice,UOM Where InvItem.ItemNo = UOM.ItemNo and InvItem.Uom = UOM.Uom and InvItem.InvNo = Invoice.InvNo and InvItem.Price = 0 and InvItem.ItemNo = " & SafeSQL(sItemNo) & " and CustId = " & SafeSQL(dtr("CustID").ToString) & strAgent & " and (Void=0 or void is null) and InvDT Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")))
            If dtrTemp.Read Then
                If dtrTemp("Qty") Is System.DBNull.Value <> True Then
                    dFoc = dtrTemp("Qty")
                End If
            End If
            dtrTemp.Close()
            dtrTemp = ReadRecordAnother("Select isnull(Sum(Quantity * UOM.BaseQty),0) as Qty from GoodsReturnItem, GoodsReturn,UOM Where GoodsReturnItem.ItemNo = UOM.ItemNo and GoodsReturnItem.Uom = UOM.Uom and GoodsReturnItem.ReturnNo = GoodsReturn.ReturnNo and GoodsReturnItem.ItemNo = " & SafeSQL(sItemNo) & " and CustNo = " & SafeSQL(dtr("CustID").ToString) & strAgent & " and (Void=0 or void is null) and ReturnDate Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")))
            If dtrTemp.Read Then
                If dtrTemp("Qty") Is System.DBNull.Value <> True Then
                    dRet = dtrTemp("Qty")
                End If
            End If
            dtrTemp.Close()
            Dim sQty As String = ""
            sQty = CStr(dQty) & "/" & CStr(dRet)
            ExecuteAnotherSQL("Insert into DailySalesInv(CustNo, CustName, AgentID, InvDt, Qty, Foc, Exchange, ShortDesc, TransType, ItemNo) values (" & SafeSQL(dtr("CustID").ToString) & "," & SafeSQL(dtr("CustName").ToString) & "," & SafeSQL(cmbAgent.Text) & "," & SafeSQL(Format(dtpDate.Value, "yyyyMMdd")) & "," & dQty & "," & dRet & "," & dEx & "," & SafeSQL(dtr("ShortDesc").ToString) & ",'Sales'," & SafeSQL(dtr("ItemNo").ToString) & ")")
            ExecuteAnotherSQL("Insert into DailySalesInv(CustNo, CustName, AgentID, InvDt, Qty, Foc, Exchange, ShortDesc, TransType, ItemNo) values (" & SafeSQL(dtr("CustID").ToString) & "," & SafeSQL(dtr("CustName").ToString) & "," & SafeSQL(cmbAgent.Text) & "," & SafeSQL(Format(dtpDate.Value, "yyyyMMdd")) & "," & dRet & "," & dRet & "," & dEx & "," & SafeSQL(dtr("ShortDesc").ToString) & ",'EX'," & SafeSQL(dtr("ItemNo").ToString) & ")")
            ExecuteAnotherSQL("Insert into DailySalesInv(CustNo, CustName, AgentID, InvDt, Qty, Foc, Exchange, ShortDesc, TransType, ItemNo) values (" & SafeSQL(dtr("CustID").ToString) & "," & SafeSQL(dtr("CustName").ToString) & "," & SafeSQL(cmbAgent.Text) & "," & SafeSQL(Format(dtpDate.Value, "yyyyMMdd")) & "," & dEx & "," & dRet & "," & dEx & "," & SafeSQL(dtr("ShortDesc").ToString) & ",'GRN'," & SafeSQL(dtr("ItemNo").ToString) & ")")
            ExecuteAnotherSQL("Insert into DailySalesInv(CustNo, CustName, AgentID, InvDt, Qty, Foc, Exchange, ShortDesc, TransType, ItemNo) values (" & SafeSQL(dtr("CustID").ToString) & "," & SafeSQL(dtr("CustName").ToString) & "," & SafeSQL(cmbAgent.Text) & "," & SafeSQL(Format(dtpDate.Value, "yyyyMMdd")) & "," & dFoc & "," & dRet & "," & dEx & "," & SafeSQL(dtr("ShortDesc").ToString) & ",'FOC'," & SafeSQL(dtr("ItemNo").ToString) & ")")
        End While
        dtr.Close()
        ExecuteReport("Select Distinct DailySalesInv.* from DailySalesInv order by ItemNo", "DailySalesInvRep")
    End Sub

    Private Sub LoadReportByAmt()
        Dim dtr As SqlDataReader
        Dim dRet As Double
        Dim dQty As Double
        Dim dFoc As Double
        Dim dEx As Double
        Dim str As String = ""
        Dim sItemNo As String = ""
        Dim strAgent As String = ""
        Dim strAgentRet As String = ""
        Dim strCust As String = ""
        Dim strCustRet As String = ""
        If cmbAgent.Text = "ALL" Then
            strAgent = " "
            strAgentRet = " "
        Else
            strAgent = " AND Invoice.AgentId = " & SafeSQL(cmbAgent.SelectedValue.ToString.Trim)
            strAgentRet = " AND GoodsReturn.SalesPersonCode = " & SafeSQL(cmbAgent.SelectedValue.ToString.Trim)
        End If
        If cmbCust.Text = "ALL" Then
            strCust = " "
            strCustRet = " "
        Else
            strCust = " AND Invoice.CustId = " & SafeSQL(cmbCust.SelectedValue.ToString.Trim)
            strCustRet = " AND GoodsReturn.CustNo = " & SafeSQL(cmbCust.SelectedValue.ToString.Trim)
        End If
        ExecuteSQL("Delete from DailySalesInv")
        str = "SELECT Distinct Customer.CustName, Invoice.CustId,  InvItem.ItemNo, Item.ShortDesc as ShortDesc, Item.DisplayNo FROM Customer, Invoice, Item, InvItem Where Customer.CustNo = Invoice.CustId and Invoice.InvNo = InvItem.InvNo and InvItem.ItemNo = Item.ItemNo and Void=0 and InvDT Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strAgent & strCust & " Union " & _
                "SELECT Distinct Customer.CustName, GoodsReturn.CustNo as CustId,  GoodsReturnItem.ItemNo, Item.ShortDesc as ShortDesc, Item.DisplayNo FROM Customer, GoodsReturn, Item, GoodsReturnItem Where Customer.CustNo = GoodsReturn.CustNo and GoodsReturn.ReturnNo = GoodsReturnItem.ReturnNo and GoodsReturnItem.ItemNo = Item.ItemNo and ReturnDate Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")) & strAgentRet & strCustRet & " Order by DisplayNo"
        dtr = ReadRecord(str)
        While dtr.Read
            sItemNo = dtr("ItemNo").ToString
            dQty = 0
            dRet = 0
            dEx = 0
            dFoc = 0
            Dim dtrTemp As SqlDataReader
            dtrTemp = ReadRecordAnother("Select isnull(Sum(SubAmt),0) as Qty from InvItem, Invoice,UOM Where InvItem.ItemNo = UOM.ItemNo and InvItem.Uom = UOM.Uom and InvItem.InvNo = Invoice.InvNo and InvItem.Description not like 'EX%' and InvItem.ItemNo = " & SafeSQL(sItemNo) & " and CustId = " & SafeSQL(dtr("CustID")) & strAgent & " and (Void=0 or void is null) and InvDT Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")))
            If dtrTemp.Read Then
                If dtrTemp("Qty") Is System.DBNull.Value <> True Then
                    dQty = dtrTemp("Qty")
                End If
            End If
            dtrTemp.Close()
            dtrTemp = ReadRecordAnother("Select isnull(Sum(SubAmt),0) as Qty from InvItem, Invoice,UOM Where InvItem.ItemNo = UOM.ItemNo and InvItem.Uom = UOM.Uom and InvItem.InvNo = Invoice.InvNo and InvItem.Description like 'EX%' and InvItem.ItemNo = " & SafeSQL(sItemNo) & " and CustId = " & SafeSQL(dtr("CustID")) & strAgent & " and (Void=0 or void is null) and InvDT Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")))
            If dtrTemp.Read Then
                If dtrTemp("Qty") Is System.DBNull.Value <> True Then
                    dEx = dtrTemp("Qty")
                End If
            End If
            dtrTemp.Close()
            dtrTemp = ReadRecordAnother("Select isnull(Sum(SubAmt),0) as Qty from InvItem, Invoice,UOM Where InvItem.ItemNo = UOM.ItemNo and InvItem.Uom = UOM.Uom and InvItem.InvNo = Invoice.InvNo and InvItem.Price = 0 and InvItem.ItemNo = " & SafeSQL(sItemNo) & " and CustId = " & SafeSQL(dtr("CustID")) & strAgent & " and (Void=0 or void is null) and InvDT Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")))
            If dtrTemp.Read Then
                If dtrTemp("Qty") Is System.DBNull.Value <> True Then
                    dFoc = dtrTemp("Qty")
                End If
            End If
            dtrTemp.Close()
            dtrTemp = ReadRecordAnother("Select isnull(Sum(Amt),0) as Qty from GoodsReturnItem, GoodsReturn,UOM Where GoodsReturnItem.ItemNo = UOM.ItemNo and GoodsReturnItem.Uom = UOM.Uom and GoodsReturnItem.ReturnNo = GoodsReturn.ReturnNo and GoodsReturnItem.ItemNo = " & SafeSQL(sItemNo) & " and CustNo = " & SafeSQL(dtr("CustID")) & strAgent & " and (Void=0 or void is null) and ReturnDate Between " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpDate.Value, "yyyyMMdd 23:59:59")))
            If dtrTemp.Read Then
                If dtrTemp("Qty") Is System.DBNull.Value <> True Then
                    dRet = dtrTemp("Qty")
                End If
            End If
            dtrTemp.Close()
            Dim sQty As String = ""
            sQty = CStr(dQty) & "/" & CStr(dRet)
            ExecuteAnotherSQL("Insert into DailySalesInv(CustNo, CustName, AgentID, InvDt, Qty, Foc, Exchange, ShortDesc, TransType, ItemNo) values (" & SafeSQL(dtr("CustID").ToString) & "," & SafeSQL(dtr("CustName").ToString) & "," & SafeSQL(cmbAgent.Text) & "," & SafeSQL(Format(dtpDate.Value, "yyyyMMdd")) & "," & dQty & "," & dRet & "," & dEx & "," & SafeSQL(dtr("ShortDesc").ToString) & ",'Sales'," & SafeSQL(dtr("ItemNo").ToString) & ")")
            ExecuteAnotherSQL("Insert into DailySalesInv(CustNo, CustName, AgentID, InvDt, Qty, Foc, Exchange, ShortDesc, TransType, ItemNo) values (" & SafeSQL(dtr("CustID").ToString) & "," & SafeSQL(dtr("CustName").ToString) & "," & SafeSQL(cmbAgent.Text) & "," & SafeSQL(Format(dtpDate.Value, "yyyyMMdd")) & "," & dRet & "," & dRet & "," & dEx & "," & SafeSQL(dtr("ShortDesc").ToString) & ",'EX'," & SafeSQL(dtr("ItemNo").ToString) & ")")
            ExecuteAnotherSQL("Insert into DailySalesInv(CustNo, CustName, AgentID, InvDt, Qty, Foc, Exchange, ShortDesc, TransType, ItemNo) values (" & SafeSQL(dtr("CustID").ToString) & "," & SafeSQL(dtr("CustName").ToString) & "," & SafeSQL(cmbAgent.Text) & "," & SafeSQL(Format(dtpDate.Value, "yyyyMMdd")) & "," & dEx & "," & dRet & "," & dEx & "," & SafeSQL(dtr("ShortDesc").ToString) & ",'GRN'," & SafeSQL(dtr("ItemNo").ToString) & ")")
            ExecuteAnotherSQL("Insert into DailySalesInv(CustNo, CustName, AgentID, InvDt, Qty, Foc, Exchange, ShortDesc, TransType, ItemNo) values (" & SafeSQL(dtr("CustID").ToString) & "," & SafeSQL(dtr("CustName").ToString) & "," & SafeSQL(cmbAgent.Text) & "," & SafeSQL(Format(dtpDate.Value, "yyyyMMdd")) & "," & dFoc & "," & dRet & "," & dEx & "," & SafeSQL(dtr("ShortDesc").ToString) & ",'FOC'," & SafeSQL(dtr("ItemNo").ToString) & ")")
        End While
        dtr.Close()

        ExecuteReport("Select Distinct DailySalesInv.* from DailySalesInv order by ItemNo", "DailySalesInvRepAmt")
    End Sub

    Private Sub DailySales_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        DisconnectDB()
    End Sub
    Private Sub DailySales_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ConnectDB()

        LoadCustomer()
        loadCombo()

    End Sub
    Public Sub LoadCustomer()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select CustNo, CustName from Customer where active = 1 order by CustNo")
        cmbCust.DataSource = Nothing
        aCust.Clear()
        aCust.Add(New ComboValues("ALL", "ALL"))
        While dtr.Read()
            aCust.Add(New ComboValues(dtr("CustNo").ToString, dtr("CustNo").ToString & " - " & dtr("CustName").ToString))
        End While
        dtr.Close()
        cmbCust.DataSource = aCust
        cmbCust.DisplayMember = "Desc"
        cmbCust.ValueMember = "Code"
        cmbCust.SelectedIndex = 0

    End Sub
    Public Sub loadCombo()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Code, Name from SalesAgent order by Name")
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

    Private Sub cmbAgent_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAgent.SelectedIndexChanged

    End Sub

    Private Sub cmbTerms_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged

    End Sub

    Private Sub Label20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

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