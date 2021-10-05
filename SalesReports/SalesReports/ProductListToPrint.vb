Imports System.Data
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Imports System.IO

Public Class ProductListToPrint
    Implements ISalesBase

    Private aFromItem As New ArrayList
    Private aToItem As New ArrayList

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


    Private Sub ProductListToPrint_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub

    Private Sub ProductListToPrint_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ConnectDB()
        LoadCategory()
    End Sub

    Private Sub LoadCategory()
        Dim rs As SqlDataReader
        cbCategory.Items.Clear()
        cbCategory.Items.Add("ALL")
        rs = ReadRecord("Select distinct Code from Category order by Code")
        While rs.Read
            cbCategory.Items.Add(rs("Code").ToString)
        End While
        rs.Close()
        cbCategory.SelectedIndex = 0

        cbSupplier.Items.Clear()
        cbSupplier.Items.Add("ALL")
        rs = ReadRecord("Select distinct left(itemno,3) SupplierCode from Item order by left(itemno,3)")
        While rs.Read
            cbSupplier.Items.Add(rs("SupplierCode").ToString)
        End While
        rs.Close()
        cbSupplier.SelectedIndex = 0

    End Sub


    Private Sub LoadItem()
        Dim sItem As String = " "
        If cbCategory.Text <> "ALL" Then
            sItem = " AND Category = " & SafeSQL(cbCategory.Text)
        End If
        If cbSupplier.Text <> "ALL" Then
            sItem = " AND left(ItemNo,3) = " & SafeSQL(cbSupplier.Text)
        End If
        Dim rs As SqlDataReader
        rs = ReadRecord("Select distinct ItemNo,Description+' '+ShortDesc as Description, BaseUOM, UnitPrice, ToPda, Barcode, CompanyName from Item where Active = 1 " & sItem & " order by Item.ItemNo")

        dgvOrdItem.Rows.Clear()
        While rs.Read
            If IsDBNull(rs("ToPDA")) = False Then
                If CBool(rs("ToPDA")) = True Then
                    Dim row As String() = New String() _
                                {False, rs("ItemNo").ToString, rs("Description").ToString, rs("BaseUOM").ToString, rs("UnitPrice").ToString, rs("Barcode").ToString, rs("CompanyName").ToString}
                    dgvOrdItem.Rows.Add(row)
                Else
                    Dim row As String() = New String() _
                               {True, rs("ItemNo").ToString, rs("Description").ToString, rs("BaseUOM").ToString, rs("UnitPrice").ToString, rs("Barcode").ToString, rs("CompanyName").ToString}
                    dgvOrdItem.Rows.Add(row)
                End If
            End If
        End While
        rs.Close()
    End Sub

    Private Sub btnSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelect.Click
        Dim i As Integer
        For i = 0 To dgvOrdItem.Rows.Count - 1
            dgvOrdItem.Item(0, i).Value = True
        Next
    End Sub

    Private Sub btnUnselect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUnselect.Click
        Dim i As Integer
        For i = 0 To dgvOrdItem.Rows.Count - 1
            dgvOrdItem.Item(0, i).Value = False
        Next
    End Sub


    Private Sub btnApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnApply.Click
        LoadItem()
    End Sub

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        If txtStartPosition.Text.Trim.Length = 0 Then
            MsgBox("Please enter Start Position.")
        End If
        If IsNumeric(txtStartPosition.Text) = False Then
            MsgBox("Start position should be numeric value between 1 to 12.")
        End If
        If CInt(txtStartPosition.Text) < 1 Or CInt(txtStartPosition.Text) > 12 Then
            MsgBox("Start position should be numeric value between 1 to 12.")
        End If
        ExecuteSQL("Delete from ItemRep")
        Dim i As Integer = 0
        Dim iStart As Integer = CInt(txtStartPosition.Text)
        For i = 1 To iStart - 1
            Dim cmd As SqlCommand = New SqlCommand("INSERT INTO [Sales].[dbo].[ItemRep]([ItemNo],[Description],[UOM],[BaseQty],[BaseUOM],[UnitPrice],[VendorNo],[Category],[Barcode],[Photo],[UOMList],[PriceList],[DisplayNo]) VALUES (@ItemNo,@Description,@UOM,@BaseQty,@BaseUOM,@UnitPrice,@VendorNo,@Category,@Barcode,@Photo,@UOMList,@PriceList,@DisplayNo)", MyConnection)

            Dim Item_ID As New SqlParameter("@ItemNo", SqlDbType.VarChar, _
                            250, ParameterDirection.Input, False, _
                           0, 0, Nothing, DataRowVersion.Current, "BLANK" & i.ToString)
            Dim Description As New SqlParameter("@Description", SqlDbType.VarChar, _
             250, ParameterDirection.Input, False, _
            0, 0, Nothing, DataRowVersion.Current, "BLANK")
            Dim UOM As New SqlParameter("@UOM", SqlDbType.VarChar, _
             250, ParameterDirection.Input, False, _
            0, 0, Nothing, DataRowVersion.Current, "BLANK")
            Dim BaseQty As New SqlParameter("@BaseQty", SqlDbType.Float, _
             250, ParameterDirection.Input, False, _
            0, 0, Nothing, DataRowVersion.Current, "0")
            Dim BaseUOM As New SqlParameter("@BaseUOM", SqlDbType.VarChar, _
             250, ParameterDirection.Input, False, _
            0, 0, Nothing, DataRowVersion.Current, "BLANK")
            Dim UnitPrice As New SqlParameter("@UnitPrice", SqlDbType.Float, _
             250, ParameterDirection.Input, False, _
            0, 0, Nothing, DataRowVersion.Current, "0")
            Dim VendorNo As New SqlParameter("@VendorNo", SqlDbType.VarChar, _
                         250, ParameterDirection.Input, False, _
                        0, 0, Nothing, DataRowVersion.Current, cbCategory.Text)
            Dim Category As New SqlParameter("@Category", SqlDbType.VarChar, _
             250, ParameterDirection.Input, False, _
            0, 0, Nothing, DataRowVersion.Current, "BLANK")
            Dim Barcode As New SqlParameter("@Barcode", SqlDbType.VarChar, _
             250, ParameterDirection.Input, False, _
            0, 0, Nothing, DataRowVersion.Current, "BLANK")
            Dim UOMList As New SqlParameter("@UOMList", SqlDbType.VarChar, _
            250, ParameterDirection.Input, False, _
           0, 0, Nothing, DataRowVersion.Current, "BLANK")
            Dim PriceList As New SqlParameter("@PriceList", SqlDbType.VarChar, _
            250, ParameterDirection.Input, False, _
           0, 0, Nothing, DataRowVersion.Current, "BLANK")
            cmd.Parameters.Add(Item_ID)
            cmd.Parameters.Add(Description)
            cmd.Parameters.Add(UOM)
            cmd.Parameters.Add(BaseQty)
            cmd.Parameters.Add(BaseUOM)
            cmd.Parameters.Add(UnitPrice)
            cmd.Parameters.Add(VendorNo)
            cmd.Parameters.Add(Category)
            cmd.Parameters.Add(Barcode)


            Dim sImg As String = Windows.Forms.Application.StartupPath & "\blank.jpg"
         
            Dim imageProduct As Image = Image.FromFile(sImg)

            Dim iHeight As Integer = 108
            Dim iWidth As Integer = 144
            If imageProduct.Size.Height / iHeight > imageProduct.Size.Width / iWidth Then
                iHeight = iHeight
                iWidth = (iHeight / imageProduct.Size.Height) * imageProduct.Size.Width
            Else
                iHeight = (iWidth / imageProduct.Size.Width) * imageProduct.Size.Height
                iWidth = iWidth
            End If

            Dim bm As New Bitmap(imageProduct, iWidth, iHeight)
            Dim ms As New MemoryStream

            bm.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp)
            Dim bytImageData(ms.Length() - 1) As Byte
            bytImageData = ms.ToArray

            Dim IMG_DATA As New SqlParameter("@Photo", SqlDbType.Image, _
                bytImageData.Length, ParameterDirection.Input, False, _
               0, 0, Nothing, DataRowVersion.Current, bytImageData)
            cmd.Parameters.Add(IMG_DATA)
            cmd.Parameters.Add(UOMList)
            cmd.Parameters.Add(PriceList)

            Dim DisplayNo As New SqlParameter("@DisplayNo", SqlDbType.Int, _
         250, ParameterDirection.Input, False, _
        0, 0, Nothing, DataRowVersion.Current, i)
            cmd.Parameters.Add(DisplayNo)
            cmd.ExecuteNonQuery()

        Next


        For i = 0 To dgvOrdItem.Rows.Count - 1
            If dgvOrdItem.Item(0, i).Value = True Then
                Dim strSql As String
                strSql = "Select distinct Item.ItemNo, Item.Description  as Description, UOM.UOM, UOM.BaseQty, Item.BaseUOM, Item.UnitPrice, Item.Category as VendorNo, Item.ShortDesc as Category, Item.Barcode,Sales.dbo.FnUOM (Item.ItemNo) as UOMList, Sales.dbo.FnPrice(Item.ItemNo, Item.CompanyName) as PriceList from Item,UOM where Item.ItemNo = UOM.ItemNo and UOM.Uom = Item.BaseUOM and Item.ItemNo = " & SafeSQL(dgvOrdItem.Item(1, i).Value) & " and Item.CompanyName = " & SafeSQL(dgvOrdItem.Item(6, i).Value) & " order by Item.ItemNo"
                Dim rs As SqlDataReader
                rs = ReadRecordAnother(strSql)
                While rs.Read = True

                    Dim cmd As SqlCommand = New SqlCommand("INSERT INTO [Sales].[dbo].[ItemRep]([ItemNo],[Description],[UOM],[BaseQty],[BaseUOM],[UnitPrice],[VendorNo],[Category],[Barcode],[Photo],[UOMList],[PriceList],[DisplayNo]) VALUES (@ItemNo,@Description,@UOM,@BaseQty,@BaseUOM,@UnitPrice,@VendorNo,@Category,@Barcode,@Photo,@UOMList,@PriceList,@DisplayNo)", MyConnection)

                    Dim Item_ID As New SqlParameter("@ItemNo", SqlDbType.VarChar, _
                                    250, ParameterDirection.Input, False, _
                                   0, 0, Nothing, DataRowVersion.Current, rs("ItemNo"))
                    Dim Description As New SqlParameter("@Description", SqlDbType.VarChar, _
                     250, ParameterDirection.Input, False, _
                    0, 0, Nothing, DataRowVersion.Current, rs("Description"))
                    Dim UOM As New SqlParameter("@UOM", SqlDbType.VarChar, _
                     250, ParameterDirection.Input, False, _
                    0, 0, Nothing, DataRowVersion.Current, rs("UOM"))
                    Dim BaseQty As New SqlParameter("@BaseQty", SqlDbType.Float, _
                     250, ParameterDirection.Input, False, _
                    0, 0, Nothing, DataRowVersion.Current, rs("BaseQty"))
                    Dim BaseUOM As New SqlParameter("@BaseUOM", SqlDbType.VarChar, _
                     250, ParameterDirection.Input, False, _
                    0, 0, Nothing, DataRowVersion.Current, rs("BaseUOM"))
                    Dim UnitPrice As New SqlParameter("@UnitPrice", SqlDbType.Float, _
                     250, ParameterDirection.Input, False, _
                    0, 0, Nothing, DataRowVersion.Current, rs("UnitPrice"))
                    Dim VendorNo As New SqlParameter("@VendorNo", SqlDbType.VarChar, _
                                 250, ParameterDirection.Input, False, _
                                0, 0, Nothing, DataRowVersion.Current, cbCategory.Text)
                    Dim Category As New SqlParameter("@Category", SqlDbType.VarChar, _
                     250, ParameterDirection.Input, False, _
                    0, 0, Nothing, DataRowVersion.Current, rs("Category"))
                    Dim Barcode As New SqlParameter("@Barcode", SqlDbType.VarChar, _
                     250, ParameterDirection.Input, False, _
                    0, 0, Nothing, DataRowVersion.Current, rs("Barcode"))
                    Dim UOMList As New SqlParameter("@UOMList", SqlDbType.VarChar, _
                    250, ParameterDirection.Input, False, _
                   0, 0, Nothing, DataRowVersion.Current, rs("UOMList"))
                    Dim PriceList As New SqlParameter("@PriceList", SqlDbType.VarChar, _
                    250, ParameterDirection.Input, False, _
                   0, 0, Nothing, DataRowVersion.Current, rs("PriceList"))
                    cmd.Parameters.Add(Item_ID)
                    cmd.Parameters.Add(Description)
                    cmd.Parameters.Add(UOM)
                    cmd.Parameters.Add(BaseQty)
                    cmd.Parameters.Add(BaseUOM)
                    cmd.Parameters.Add(UnitPrice)
                    cmd.Parameters.Add(VendorNo)
                    cmd.Parameters.Add(Category)
                    cmd.Parameters.Add(Barcode)


                    Dim sImg As String = Windows.Forms.Application.StartupPath & "\Photo\" & rs("ItemNo").ToString & ".jpg"
                    If File.Exists(sImg) Then
                    Else
                        sImg = Windows.Forms.Application.StartupPath & "\blank.jpg"
                    End If
                    Dim imageProduct As Image = Image.FromFile(sImg)

                    Dim iHeight As Integer = 108
                    Dim iWidth As Integer = 144
                    If imageProduct.Size.Height / iHeight > imageProduct.Size.Width / iWidth Then
                        iHeight = iHeight
                        iWidth = (iHeight / imageProduct.Size.Height) * imageProduct.Size.Width
                    Else
                        iHeight = (iWidth / imageProduct.Size.Width) * imageProduct.Size.Height
                        iWidth = iWidth
                    End If

                    Dim bm As New Bitmap(imageProduct, iWidth, iHeight)
                    Dim ms As New MemoryStream

                     bm.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp)
                    Dim bytImageData(ms.Length() - 1) As Byte
                    bytImageData = ms.ToArray


                    Dim IMG_DATA As New SqlParameter("@Photo", SqlDbType.Image, _
                        bytImageData.Length, ParameterDirection.Input, False, _
                       0, 0, Nothing, DataRowVersion.Current, bytImageData)
                    cmd.Parameters.Add(IMG_DATA)
                    cmd.Parameters.Add(UOMList)
                    cmd.Parameters.Add(PriceList)

                    Dim DisplayNo As New SqlParameter("@DisplayNo", SqlDbType.Int, _
                             250, ParameterDirection.Input, False, _
                            0, 0, Nothing, DataRowVersion.Current, iStart)
                    cmd.Parameters.Add(DisplayNo)

                    cmd.ExecuteNonQuery()

                    iStart = iStart + 1
                End While
                rs.Close()
            End If
        Next


        ExecuteReport("Select * from ItemRep order by DisplayNo,unitprice * baseqty", "ListItemCatalogRep")
    End Sub
End Class