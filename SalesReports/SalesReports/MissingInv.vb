Imports System.Data
Imports System.Data.SqlClient
Imports DataInterface.IbizDO
Imports SalesInterface.MobileSales
Public Class MissingInv
    Implements ISalesBase

    Private objDO As New DataInterface.IbizDO

    Private Sub MissingInv_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        DisconnectDB()
    End Sub

    Private Sub MissingInv_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ConnectDB()
        loadCombo()
    End Sub
    Public Sub loadCombo()
        Dim dtr As SqlDataReader
        dtr = ReadRecord("Select AgentID from MDT")
        cmbAgent.Items.Add("All")
        Do While dtr.Read = True
            cmbAgent.Items.Add(dtr(0))
        Loop
        dtr.Close()
        If cmbAgent.Items.Count > 0 Then
            cmbAgent.SelectedIndex = 0
        End If
    End Sub
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim sPreInv As String = ""
        Dim lNum, sNum, i As Integer
        Dim sInvNo As String = ""
        Dim strSql As String = ""
        Dim dtr As SqlDataReader
        Dim rs As SqlDataReader
        ExecuteSQL("Delete from MissingInv")
        If cmbAgent.Text = "All" Then
            For i = 1 To cmbAgent.Items.Count - 1
                cmbAgent.SelectedIndex = i
                rs = ReadRecord("Select * from MDT where AgentID=" & SafeSQL(cmbAgent.Text))
                If rs.Read = True Then
                    sPreInv = rs("PreInvNo").ToString()
                    lNum = CInt(rs("LastInvNo"))
                End If
                rs.Close()
                strSql = "SELECT InvNo, InvDt, CustID, AgentID, TotalAmt FROM Invoice where InvDt Between" & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & " and AgentID=" & SafeSQL(cmbAgent.Text) & "and Void=0"
                dtr = ReadRecord(strSql)
                While dtr.Read
                    ExecuteAnotherSQL("Insert into MissingInv(InvNo, InvDt, CustId, AgentId, TotalAmt, FromDate, ToDate, Type) values (" & SafeSQL(dtr("InvNo")) & "," & SafeSQL(Format(dtr("InvDt"), "dd-MMM-yyyy")) & "," & SafeSQL(dtr("CustID")) & "," & SafeSQL(dtr("AgentID")) & "," & CDbl(dtr("TotalAmt")) & "," & SafeSQL(Format(dtpFromDate.Value, "dd-MMM-yyyy")) & "," & SafeSQL(Format(dtpToDate.Value, "dd-MMM-yyyy")) & ",'MASTER INVOICES')")
                End While
                dtr.Close()
                ''''Void Invoices
                strSql = "SELECT InvNo, InvDt, CustID, AgentID, TotalAmt FROM Invoice where InvDt Between" & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & " and AgentID=" & SafeSQL(cmbAgent.Text) & " and Void=1"
                dtr = ReadRecord(strSql)
                While dtr.Read
                    ExecuteAnotherSQL("Insert into MissingInv(InvNo, InvDt, CustId, AgentId, TotalAmt, FromDate, ToDate, Type) values (" & SafeSQL(dtr("InvNo")) & "," & SafeSQL(Format(dtr("InvDt"), "dd-MMM-yyyy")) & "," & SafeSQL(dtr("CustID")) & "," & SafeSQL(dtr("AgentID")) & "," & CDbl(dtr("TotalAmt")) & "," & SafeSQL(Format(dtpFromDate.Value, "dd-MMM-yyyy")) & "," & SafeSQL(Format(dtpToDate.Value, "dd-MMM-yyyy")) & ",'VOID INVOICES')")
                End While
                dtr.Close()
                ''''Unpaid Invoices
                strSql = "SELECT InvNo, InvDt, CustID, AgentID, TotalAmt FROM Invoice where InvDt Between" & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & " and AgentID=" & SafeSQL(cmbAgent.Text) & " and TotalAmt > PaidAmt"
                dtr = ReadRecord(strSql)
                While dtr.Read
                    ExecuteAnotherSQL("Insert into MissingInv(InvNo, InvDt, CustId, AgentId, TotalAmt, FromDate, ToDate, Type) values (" & SafeSQL(dtr("InvNo")) & "," & SafeSQL(Format(dtr("InvDt"), "dd-MMM-yyyy")) & "," & SafeSQL(dtr("CustID")) & "," & SafeSQL(dtr("AgentID")) & "," & CDbl(dtr("TotalAmt")) & "," & SafeSQL(Format(dtpFromDate.Value, "dd-MMM-yyyy")) & "," & SafeSQL(Format(dtpToDate.Value, "dd-MMM-yyyy")) & ",'UNPAID INVOICES')")
                End While
                dtr.Close()
                ''''Missing Invoices
                strSql = "SELECT InvNo, InvDt, CustID, AgentID, TotalAmt FROM Invoice where InvDt Between" & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & " and AgentID=" & SafeSQL(cmbAgent.Text) & " Order By InvNo"
                dtr = ReadRecord(strSql)
                If dtr.Read = True Then
                    sNum = CInt(Mid(dtr("InvNo").ToString, Len(sPreInv) + 1, Len(dtr("InvNo").ToString)))
                    lNum = sNum
                End If
                dtr.Close()
                'For i = sNum To lNum
                '   sInvNo = sPreInv & i
                strSql = "SELECT InvNo, InvDt, CustID, AgentID, TotalAmt FROM Invoice where InvDt Between" & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & " and AgentID=" & SafeSQL(cmbAgent.Text) & " Order By InvNo"
                dtr = ReadRecord(strSql)
                While dtr.Read
                    sNum = CInt(Mid(dtr("InvNo").ToString, Len(sPreInv) + 1, Len(dtr("InvNo").ToString)))
                    If sNum <> lNum Then
                        ExecuteAnotherSQL("Insert into MissingInv(InvNo, InvDt, CustId, AgentId, TotalAmt, FromDate, ToDate, Type) values (" & SafeSQL(sPreInv & lNum) & ",'','','',0," & SafeSQL(Format(dtpFromDate.Value, "dd-MMM-yyyy")) & "," & SafeSQL(Format(dtpToDate.Value, "dd-MMM-yyyy")) & ",'MISSING INVOICES')")
                        lNum = sNum
                    End If
                    lNum = sNum + 1
                End While
                dtr.Close()
            Next
        Else
            rs = ReadRecord("Select * from MDT where AgentID=" & SafeSQL(cmbAgent.Text))
            If rs.Read = True Then
                sPreInv = rs("PreInvNo").ToString()
                lNum = CInt(rs("LastInvNo"))
            End If
            rs.Close()
            'ExecuteSQL("Delete from MissingInv")
            ''''Master Invoices
            strSql = "SELECT InvNo, InvDt, CustID, AgentID, TotalAmt FROM Invoice where InvDt Between" & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & " and AgentID=" & SafeSQL(cmbAgent.Text) & "and Void=0"
            dtr = ReadRecord(strSql)
            While dtr.Read
                ExecuteAnotherSQL("Insert into MissingInv(InvNo, InvDt, CustId, AgentId, TotalAmt, FromDate, ToDate, Type) values (" & SafeSQL(dtr("InvNo")) & "," & SafeSQL(Format(dtr("InvDt"), "dd-MMM-yyyy")) & "," & SafeSQL(dtr("CustID")) & "," & SafeSQL(dtr("AgentID")) & "," & CDbl(dtr("TotalAmt")) & "," & SafeSQL(Format(dtpFromDate.Value, "dd-MMM-yyyy")) & "," & SafeSQL(Format(dtpToDate.Value, "dd-MMM-yyyy")) & ",'MASTER INVOICES')")
            End While
            dtr.Close()
            ''''Void Invoices
            strSql = "SELECT InvNo, InvDt, CustID, AgentID, TotalAmt FROM Invoice where InvDt Between" & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & " and AgentID=" & SafeSQL(cmbAgent.Text) & " and Void=1"
            dtr = ReadRecord(strSql)
            While dtr.Read
                ExecuteAnotherSQL("Insert into MissingInv(InvNo, InvDt, CustId, AgentId, TotalAmt, FromDate, ToDate, Type) values (" & SafeSQL(dtr("InvNo")) & "," & SafeSQL(Format(dtr("InvDt"), "dd-MMM-yyyy")) & "," & SafeSQL(dtr("CustID")) & "," & SafeSQL(dtr("AgentID")) & "," & CDbl(dtr("TotalAmt")) & "," & SafeSQL(Format(dtpFromDate.Value, "dd-MMM-yyyy")) & "," & SafeSQL(Format(dtpToDate.Value, "dd-MMM-yyyy")) & ",'VOID INVOICES')")
            End While
            dtr.Close()
            ''''Unpaid Invoices
            strSql = "SELECT InvNo, InvDt, CustID, AgentID, TotalAmt FROM Invoice where InvDt Between" & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & " and AgentID=" & SafeSQL(cmbAgent.Text) & " and TotalAmt > PaidAmt"
            dtr = ReadRecord(strSql)
            While dtr.Read
                ExecuteAnotherSQL("Insert into MissingInv(InvNo, InvDt, CustId, AgentId, TotalAmt, FromDate, ToDate, Type) values (" & SafeSQL(dtr("InvNo")) & "," & SafeSQL(Format(dtr("InvDt"), "dd-MMM-yyyy")) & "," & SafeSQL(dtr("CustID")) & "," & SafeSQL(dtr("AgentID")) & "," & CDbl(dtr("TotalAmt")) & "," & SafeSQL(Format(dtpFromDate.Value, "dd-MMM-yyyy")) & "," & SafeSQL(Format(dtpToDate.Value, "dd-MMM-yyyy")) & ",'UNPAID INVOICES')")
            End While
            dtr.Close()
            ''''Missing Invoices
            strSql = "SELECT InvNo, InvDt, CustID, AgentID, TotalAmt FROM Invoice where InvDt Between" & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & " and AgentID=" & SafeSQL(cmbAgent.Text) & " Order By InvNo"
            dtr = ReadRecord(strSql)
            If dtr.Read = True Then
                sNum = CInt(Mid(dtr("InvNo").ToString, Len(sPreInv) + 1, Len(dtr("InvNo").ToString)))
                lNum = sNum
            End If
            dtr.Close()
            'For i = sNum To lNum
            '   sInvNo = sPreInv & i
            strSql = "SELECT InvNo, InvDt, CustID, AgentID, TotalAmt FROM Invoice where InvDt Between" & SafeSQL(Format(dtpFromDate.Value, "yyyyMMdd 00:00:00")) & " and " & SafeSQL(Format(dtpToDate.Value, "yyyyMMdd 23:59:59")) & " and AgentID=" & SafeSQL(cmbAgent.Text) & " Order By InvNo"
            dtr = ReadRecord(strSql)
            While dtr.Read
                sNum = CInt(Mid(dtr("InvNo").ToString, Len(sPreInv) + 1, Len(dtr("InvNo").ToString)))
                If sNum <> lNum Then
                    ExecuteAnotherSQL("Insert into MissingInv(InvNo, InvDt, CustId, AgentId, TotalAmt, FromDate, ToDate, Type) values (" & SafeSQL(sPreInv & lNum) & ",'','','',0," & SafeSQL(Format(dtpFromDate.Value, "dd-MMM-yyyy")) & "," & SafeSQL(Format(dtpToDate.Value, "dd-MMM-yyyy")) & ",'MISSING INVOICES')")
                    lNum = sNum
                End If
                lNum = sNum + 1
            End While
            dtr.Close()
        End If
        strSql = "SELECT Distinct CustName,CustNo from Customer where CustNo in (Select CustId from missingInv)"
        dtr = ReadRecord(strSql)
        While dtr.Read
            ExecuteAnotherSQL("Update MissingInv set CustName =" & SafeSQL(dtr("CustName")) & " where CustId=" & SafeSQL(dtr("CustNo")))
        End While
        dtr.Close()
        ExecuteReport("Select * from MissingInv", "MissingInvRep")
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