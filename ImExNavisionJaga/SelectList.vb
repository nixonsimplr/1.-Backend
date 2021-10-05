Imports System
Imports System.Resources
Imports System.Globalization
Imports System.Threading
Imports System.Reflection

Imports SalesInterface.MobileSales

Public Class SelectList
    Implements ISalesBase


#Region "Localization Variants"
    Private rMgr As ResourceManager
    Private cInfo As CultureInfo
    Private sLang As String
#End Region
    Dim i As Integer
    Public frm As Imex

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        '  Me.Hide()
        '   System.Threading.Thread.Sleep(50)
        ' Me.Refresh()
        'Commented for Gold Killi
        'frm.ImportBOM()
        'frm.ImportItemVariant()
        'frm.ImportAssemblyBOM()
        'frm.ImportReservationQty()
        'frm.ImportPOLine()
        Dim cnt As Integer = 0
        For i = 0 To clbImport.Items.Count - 1
            If clbImport.GetItemChecked(i) = True Then
                frm.dgvStatus.Rows.Add(clbImport.Items(i).ToString, "")
                Select Case clbImport.Items(i).ToString
                    Case "Category"
                        frm.ImportCategory()
                    Case "Product"
                        frm.ImportProduct()
                    Case "Inventory"
                        frm.ImportInventory()
                        '           Case "Country"
                        '     frm.ImportCountry()
                    Case "Customer"
                        frm.ImportCustomer()
                    Case "Invoice Discount"
                        frm.ImportCustInvDiscount()
                    Case "Invoice"
                        frm.ImportInvoice()
                    Case "Item Price"
                        frm.ImportItemPrice()
                    Case "Location"
                        frm.ImportLocation()
                    Case "Agent"
                        frm.ImportSalesAgent()
                    Case "Shipment Method"
                        frm.ImportShipMethod()
                    Case "Payment Method"
                        frm.ImportPayMethod()
                    Case "Payment Terms"
                        frm.ImportPayterms()
                    Case "Price Group"
                        frm.ImportPriceGroup()
                    Case "UOM"
                        frm.ImportUOM()
                    Case "Zone"
                        frm.ImportZone()
                    Case "Customer Product"
                        frm.ImportCustProd()
                End Select
                frm.GetImagetick(cnt)
                Me.Refresh()
                '        frm.Refresh()
                cnt = cnt + 1
            End If
        Next
        frm.DisConnect()
        Me.Close()
    End Sub

    Private Sub SelectList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ArrList As New ArrayList
        ArrList.Add("Category")
        ArrList.Add("Product")
        ArrList.Add("Inventory")
        ArrList.Add("Country")
        ArrList.Add("Customer")
        ArrList.Add("Invoice Discount")
        ArrList.Add("Invoice")
        ArrList.Add("Item Price")
        ArrList.Add("Location")
        ArrList.Add("Agent")
        ArrList.Add("Shipment Method")
        ArrList.Add("Payment Method")
        ArrList.Add("Payment Terms")
        ArrList.Add("Price Group")
        ArrList.Add("UOM")
        ArrList.Add("Zone")
        ArrList.Add("Customer Product")
        For i = 0 To ArrList.Count - 1
            clbImport.Items.Add(ArrList.Item(i).ToString, True)
            'dgvStatus.Rows.Add(ArrList.Item(i).ToString, imgTick)
        Next
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

    'Localization
    Private Sub Localization()
        Try
            cInfo = New CultureInfo(sLang)
            Thread.CurrentThread.CurrentCulture = cInfo
            Thread.CurrentThread.CurrentUICulture = cInfo
            rMgr = New ResourceManager("ImExNavision.ImExNavision", [Assembly].GetExecutingAssembly())

            Me.Text = rMgr.GetString("SelectList")
            btnOK.Text = rMgr.GetString("btnOK")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub SetCulture(ByVal CultureName As String) Implements SalesInterface.MobileSales.ISalesBase.SetCulture
        sLang = CultureName
        Localization()
    End Sub
End Class