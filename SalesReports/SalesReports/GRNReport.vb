Imports System.Data
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales

Public Class GRNReport
    Implements ISalesBase

    Private aItem As New ArrayList()
    Private aAgent As New ArrayList()
    Dim strAgent As String = "", strItem As String = "", sLocation As String = ""
    Dim ArList As New ArrayList

    Private Structure ArrItemList
        Dim ItemCode As String
        Dim Desc As String
        Dim Qty As Double
        'Dim BQty As Double
        Dim Amt As Double
        'Dim BAmt As Double
    End Structure

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        If dtpToDate.Value.Month <> dtpFromDate.Value.Month Then
            MsgBox("Sorry, please select date on the same month !!")
            Exit Sub
        End If
        If dtpToDate.Value.Day < dtpFromDate.Value.Day Then
            MsgBox("Sorry, day value on To Date should be bigger than day value on From Date !!")
            Exit Sub
        End If
        btnPrint.Enabled = False
        'If cmbAgent.Text = "ALL" Then
        '    strAgent = " "
        'Else
        '    strAgent = "and SalesAgent.Code=" & "'" & cmbAgent.SelectedValue & "'"
        'End If
        'If cmbItem.Text = "ALL" Then
        '    strItem = ""
        'Else
        '    strItem = " and InvItem.ItemNo in (Select ItemNo from Item where Category='" & cmbCategory.SelectedValue.ToString.Trim & "')"
        'End If
        If rdoSum.Checked = True Then
            ExecuteSQL("Delete from GRNRptBySum")
        Else
            ExecuteSQL("Delete from GRNRptByDet")
            ExecuteSQL("Delete from GRNRptByDetTemp")
        End If

        Dim strSql As String = ""

        Dim dtr As SqlDataReader
        dtr = ReadRecord("SELECT location from MDT where AgentID=" & SafeSQL(cmbAgent.SelectedValue))
        If dtr.Read = True Then
            sLocation = dtr("Location").ToString
        End If
        dtr.Close()

        'Dim str As String
        'str = "Select Item.ItemNo, Item.Description, IsNull(Sum(CreditNoteDet.Qty),0) as Qty,IsNull(Sum(TotalAmt),0) as Amt from CreditNote, CreditNoteDet, Item where (Void =0 or Void is Null) and CreditNote.CreditNoteNo=CreditNoteDet.CreditNoteNo and CreditNoteDet.ItemNo=Item.ItemNo and CreditDate between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & "  and SalesPersonCode = " & SafeSQL(cmbAgent.SelectedValue) & " group by CreditNoteDet.ItemNo,Item.Description"
        'dtr = ReadRecord(str)
        'While dtr.Read = True
        '    ExecuteSQL("Insert into GRNRptBySum(ItemId,ItemName,Qty,Amt,AgentId,AgentName,FromDate,ToDate) Values(" & SafeSQL(dtr("ItemNo").ToString) & "," & dtr("Description").ToString & "," & Format(dtr("Qty"), "0") & "," & Format(dtr("Amt"), "0.00") & "," & SafeSQL(cmbAgent.SelectedValue.ToString) & "," & SafeSQL(cmbAgent.Text) & "," & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & "," & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 00:00:00")) & ")")
        '    'Dim aPr1 As ArrItemList
        '    'aPr1.ItemCode = dtr("ItemNo").ToString
        '    'aPr1.Desc = dtr("Description").ToString
        '    'aPr1.Qty = dtr("Qty")
        '    'aPr1.Amt = dtr("Amt")
        '    'ArList.Add(aPr1)
        'End While
        'dtr.Close()

        If rdoSum.Checked = True Then
            If cmbAgent.Text.ToUpper = "ALL" Then
                strSql = "Insert GRNRptBySum(ItemId,ItemName,Qty,Amt,AgentId,AgentName,FromDate,ToDate) Select Item.ItemNo, Item.Description, IsNull(Sum(CreditNoteDet.Qty),0) * BaseQty as Qty, IsNull(Sum(Amt),0) as Amt, SalesPersonCode, SalesAgent.Name," & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & "," & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 00:00:00")) & " from CreditNote, CreditNoteDet, Item, SalesAgent, Uom where CreditNoteDet.Uom = Uom.Uom and CreditNoteDet.ItemNo = Uom.ItemNo and CreditNote.SalesPersonCode = SalesAgent.Code and (Void =0 or Void is Null) and CreditNote.CreditNoteNo=CreditNoteDet.CreditNoteNo and CreditNoteDet.ItemNo=Item.ItemNo and CreditDate between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & " group by SalesPersonCode,SalesAgent.Name,Item.ItemNo,Item.Description,Uom.BaseQty"
            Else
                strSql = "Insert GRNRptBySum(ItemId,ItemName,Qty,Amt,AgentId,AgentName,FromDate,ToDate) Select Item.ItemNo, Item.Description, IsNull(Sum(CreditNoteDet.Qty),0) * BaseQty as Qty, IsNull(Sum(Amt),0) as Amt," & SafeSQL(cmbAgent.SelectedValue.ToString) & "," & SafeSQL(cmbAgent.Text) & "," & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & "," & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 00:00:00")) & " from CreditNote, CreditNoteDet, Item, Uom where CreditNoteDet.Uom = Uom.Uom and CreditNoteDet.ItemNo = Uom.ItemNo and (Void =0 or Void is Null) and CreditNote.CreditNoteNo=CreditNoteDet.CreditNoteNo and CreditNoteDet.ItemNo=Item.ItemNo and CreditDate between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & " and SalesPersonCode = " & SafeSQL(cmbAgent.SelectedValue) & " group by Item.ItemNo,Item.Description,Uom.BaseQty"
            End If
            ExecuteSQL(strSql)

            strSql = "SELECT ItemId,ItemName,Qty,Amt,AgentId,AgentName,FromDate,ToDate from GRNRptBySum order by ItemId"
            ExecuteReport(strSql, "GRNRptBySum")
        Else
            If cmbAgent.Text.ToUpper = "ALL" Then
                strSql = "Insert GRNRptByDetTemp(GRNNo, Date, ExtNo, AgentID, AgentName, CustomerID, CustomerName, ItemID, ItemName, BaseUom, TotalQty, GoodQty, BadQty, Price, TotalAmt, GoodAmt, BadAmt, FromDate, ToDate) SELECT CreditNote.CreditNoteNo, CreditNote.CreditDate, isnull(PoNo,''), CreditNote.SalesPersonCode, SalesAgent.Name, CreditNote.CustNo, Customer.CustName, CreditNoteDet.ItemNo, Item.Description, Item.BaseUom, CreditNoteDet.Qty * Uom.BaseQty, CreditNoteDet.Qty * Uom.BaseQty, 0, CreditNoteDet.Price * Round(1 / Uom.BaseQty, 0), CreditNoteDet.Amt, CreditNoteDet.Amt, 0, " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & "," & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 00:00:00")) & " FROM CreditNote INNER JOIN CreditNoteDet ON CreditNote.CreditNoteNo = CreditNoteDet.CreditNoteNo INNER JOIN UOM ON CreditNoteDet.ItemNo = Uom.ItemNo and CreditNoteDet.Uom = Uom.Uom INNER JOIN SalesAgent ON CreditNote.SalesPersonCode = SalesAgent.Code INNER JOIN Customer ON CreditNote.CustNo = Customer.CustNo INNER JOIN Item ON CreditNoteDet.ItemNo = Item.ItemNo WHERE CreditNote.CreditDate between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & " order by CreditDate"
            Else
                strSql = "Insert GRNRptByDetTemp(GRNNo, Date, ExtNo, AgentID, AgentName, CustomerID, CustomerName, ItemID, ItemName, BaseUom, TotalQty, GoodQty, BadQty, Price, TotalAmt, GoodAmt, BadAmt, FromDate, ToDate) SELECT CreditNote.CreditNoteNo, CreditNote.CreditDate, isnull(PoNo,''), CreditNote.SalesPersonCode, SalesAgent.Name, CreditNote.CustNo, Customer.CustName, CreditNoteDet.ItemNo, Item.Description, Item.BaseUom, CreditNoteDet.Qty * Uom.BaseQty, CreditNoteDet.Qty * Uom.BaseQty, 0, CreditNoteDet.Price * Round(1 / Uom.BaseQty, 0), CreditNoteDet.Amt, CreditNoteDet.Amt, 0, " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & "," & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 00:00:00")) & " FROM CreditNote INNER JOIN CreditNoteDet ON CreditNote.CreditNoteNo = CreditNoteDet.CreditNoteNo INNER JOIN UOM ON CreditNoteDet.ItemNo = Uom.ItemNo and CreditNoteDet.Uom = Uom.Uom INNER JOIN SalesAgent ON CreditNote.SalesPersonCode = SalesAgent.Code INNER JOIN Customer ON CreditNote.CustNo = Customer.CustNo INNER JOIN Item ON CreditNoteDet.ItemNo = Item.ItemNo WHERE CreditNote.CreditDate between " & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & " and SalesAgent.Code = " & SafeSQL(cmbAgent.SelectedValue) & " order by CreditDate"
            End If
            ExecuteSQL(strSql)

            strSql = "Insert GRNRptByDet Select AgentID,AgentName,Date,GRNNo,ExtNo,CustomerID,CustomerName,sum(TotalQty),sum(GoodQty),sum(BadQty),sum(TotalAmt),sum(GoodAmt),sum(BadAmt),FromDate,ToDate,ItemId,ItemName,avg(Price),BaseUom from GRNRptByDetTemp group by GRNNo,ItemId,ItemName,AgentID,AgentName,Date,ExtNo,CustomerID,CustomerName,FromDate,ToDate,BaseUom"
            ExecuteSQL(strSql)

            dtr = ReadRecord("Select * from GRNRptByDet")
            While dtr.Read()
                ExecuteAnotherSQL("Update GRNRptByDet set TotalAmt = TotalAmt + (isnull((select sum(Qty) * BaseQty from ItemTrans, Uom where ItemNo = ItemId and Uom.Uom = ItemTrans.Uom and DocNo =" & SafeSQL(dtr("GRNNo")) & " and DocType= 'BOUT' and ItemID = " & SafeSQL(dtr("ItemID")) & " group by BaseQty),0) * " & dtr("Price") & "), BadAmt = (isnull((select sum(Qty) * BaseQty from ItemTrans, Uom where ItemNo = ItemId and Uom.Uom = ItemTrans.Uom and DocNo = " & SafeSQL(dtr("GRNNo")) & " and DocType='BOUT' and ItemID = " & SafeSQL(dtr("ItemID")) & " group by BaseQty),0) * " & dtr("Price") & "), TotalQty = TotalQty + (isnull((select sum(Qty) * BaseQty from ItemTrans, Uom where ItemNo = ItemId and Uom.Uom = ItemTrans.Uom and DocNo =" & SafeSQL(dtr("GRNNo")) & " and DocType='BOUT' and ItemID = " & SafeSQL(dtr("ItemID")) & " group by BaseQty),0)), BadQty = (isnull((select sum(Qty) * BaseQty from ItemTrans, Uom where ItemNo = ItemId and Uom.Uom =ItemTrans.Uom and DocNo =" & SafeSQL(dtr("GRNNo")) & " and DocType='BOUT' and ItemID = " & SafeSQL(dtr("ItemID")) & " group by BaseQty),0)) where GRNRptByDet.GRNNo = " & SafeSQL(dtr("GRNNo")) & " and ItemID = " & SafeSQL(dtr("ItemID")) & " ")
            End While
            dtr.Close()

            strSql = "SELECT * from GRNRptByDet order by ItemId"
            ExecuteReport(strSql, "GRNRptByDet")
        End If

        btnPrint.Enabled = True
    End Sub

    Private Sub GRNReport_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub

    Private Sub GRNReport_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ConnectDB()
        'ExecuteSQL("Drop table InvoiceList")
        'ExecuteSQL("SELECT Customer.CustNo, Customer.CustName, Customer.PaymentMethod, Invoice.InvNo, Invoice.InvDt, Invoice.AgentID, Invoice.Discount, Invoice.SubTotal, Invoice.GST, Invoice.TotalAmt, Invoice.PayTerms into InvoiceList FROM Customer INNER JOIN Invoice ON Customer.CustNo = Invoice.CustId")
        LoadCategory()
        loadCombo()
        dtpFromDate.Value = Date.Now.AddDays(-Date.Now.Day + 1)
    End Sub
    Public Sub loadCombo()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select Code, Name from SalesAgent, MDT where MDT.AgentID=SalesAgent.Code")
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

    Public Sub LoadCategory()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select ItemNo as Code, Description from Item order by Description")
        cmbItem.DataSource = Nothing
        aItem.Clear()
        aItem.Add(New ComboValues("ALL", "ALL"))
        While dtr.Read()
            aItem.Add(New ComboValues(dtr("Code").ToString, dtr("Description").ToString))
        End While
        dtr.Close()
        cmbItem.DataSource = aItem
        cmbItem.DisplayMember = "Desc"
        cmbItem.ValueMember = "Code"
        cmbItem.SelectedIndex = 0
    End Sub

    Private Sub dtpToDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpToDate.ValueChanged

    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub dtpFromDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpFromDate.ValueChanged

    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

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