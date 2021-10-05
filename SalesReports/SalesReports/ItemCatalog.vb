Imports System.Data
Imports System.Data.SqlClient
Imports SalesInterface.MobileSales
Imports System.IO

Public Class ItemCatalog
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


    Private Sub ItemCatalog_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DisconnectDB()
    End Sub

    Private Sub ItemCatalog_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
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

        LoadItem()

        cbCompany.Items.Clear()
        rs = ReadRecord("Select distinct CompanyName from Item order by CompanyName")
        While rs.Read
            cbCompany.Items.Add(rs("CompanyName").ToString)
        End While
        rs.Close()
        cbCompany.SelectedIndex = 0
    End Sub

    Private Sub LoadItem()
        Dim sItem As String = " "
        If cbCategory.Text <> "ALL" Then
            sItem = " AND Category = " & SafeSQL(cbCategory.Text)
        End If
        If cbSupplier.Text <> "ALL" Then
            sItem = " AND left(ItemNo,3) = " & SafeSQL(cbSupplier.Text)
        End If
        aFromItem = New ArrayList
        aToItem = New ArrayList
        Dim rs As SqlDataReader
        cbFromItemNo.DataSource = Nothing
        cbToItemNo.DataSource = Nothing
        aFromItem.Add(New ComboValues("ALL", "ALL"))
        aToItem.Add(New ComboValues("ALL", "ALL"))
        rs = ReadRecord("Select distinct ItemNo,Description+' '+ShortDesc as Description from Item where Active = 1 and ToPda = 1 " & sItem & " order by Item.ItemNo")
        While rs.Read
            aFromItem.Add(New ComboValues(rs("ItemNo").ToString, rs("ItemNo").ToString & " (" & rs("Description").ToString & ")"))
            aToItem.Add(New ComboValues(rs("ItemNo").ToString, rs("ItemNo").ToString & " (" & rs("Description").ToString & ")"))
        End While
        rs.Close()

        cbFromItemNo.DataSource = aFromItem
        cbFromItemNo.DisplayMember = "Desc"
        cbFromItemNo.ValueMember = "Code"
        cbFromItemNo.SelectedIndex = 0
        cbToItemNo.DataSource = aToItem
        cbToItemNo.DisplayMember = "Desc"
        cbToItemNo.ValueMember = "Code"
        cbToItemNo.SelectedIndex = 0
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim sCategory As String = " "
        Dim sItemNo As String = " "
        Dim sSupplier As String = " "
        ExecuteSQL("delete from itemrep")
        If cbCategory.Text <> "ALL" Then
            sCategory = " and Item.Category = " & SafeSQL(cbCategory.Text)
        End If
        If cbFromItemNo.Text <> "ALL" Or cbToItemNo.Text <> "ALL" Then
            sItemNo = " and Item.ItemNo between " & SafeSQL(cbFromItemNo.SelectedValue) & " and " & SafeSQL(cbToItemNo.SelectedValue)
        End If
        If cbSupplier.Text <> "ALL" Then
            sSupplier = " AND left(Item.ItemNo,3) = " & SafeSQL(cbSupplier.Text)
        End If
        Dim strSql As String = ""
        strSql = strSql & "Select distinct Item.ItemNo, Item.Description  as Description, UOM.UOM, UOM.BaseQty, Item.BaseUOM, Item.UnitPrice, Item.Category as VendorNo, Item.ShortDesc as Category, Item.Barcode,Sales.dbo.FnUOM (Item.ItemNo) as UOMList, Sales.dbo.FnPrice( Item.ItemNo, Item.CompanyName) as PriceList from Item,UOM where Item.ItemNo = UOM.ItemNo and Item.Active = 1 and Item.ToPDA = 1 and UOM.Uom = Item.BaseUOM " & sCategory & sItemNo & sSupplier & " and CompanyName = " & SafeSQL(cbCompany.Text) & " order by Item.ItemNo"
        Dim rs As SqlDataReader
        rs = ReadRecordAnother(strSql)
        While rs.Read = True

            Dim cmd As SqlCommand = New SqlCommand("INSERT INTO [Sales].[dbo].[ItemRep]([ItemNo],[Description],[UOM],[BaseQty],[BaseUOM],[UnitPrice],[VendorNo],[Category],[Barcode],[Photo],[UOMList],[PriceList]) VALUES (@ItemNo,@Description,@UOM,@BaseQty,@BaseUOM,@UnitPrice,@VendorNo,@Category,@Barcode,@Photo,@UOMList,@PriceList)", MyConnection)

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
            'bm.Save("C:\Photo.bmp")
            'PictureBox1.Image = bm
            'MsgBox("height: " & bm.Height & vbCrLf & "width: " & bm.Width)
            Dim ms As New MemoryStream
           
            '            MsgBox("height: " & PictureBox1.Image.Height & vbCrLf & "width: " & PictureBox1.Image.Width)
            bm.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp)
            Dim bytImageData(ms.Length() - 1) As Byte
            bytImageData = ms.ToArray
           
            ''Dim fsImageFile As New System.IO.FileStream("C:\Photo.jpg", _
            ''                                     FileMode.Open, FileAccess.Read)
            ''Dim bytImageData(fsImageFile.Length() - 1) As Byte
            ''fsImageFile.Read(bytImageData, 0, bytImageData.Length)
            ''fsImageFile.Close()


            ''PictureBox1.Image.Save("C:\Photo.jpg", System.Drawing.Imaging.ImageFormat.Jpeg)

            ''Dim fsImageFile As New System.IO.FileStream("C:\Photo.jpg", _
            ''                                     FileMode.Open, FileAccess.Read)
            ''Dim bytImageData(fsImageFile.Length() - 1) As Byte
            ''fsImageFile.Read(bytImageData, 0, bytImageData.Length)
            ''fsImageFile.Close()

            'Dim sImg As String = Windows.Forms.Application.StartupPath & "\Photo\" & rs("ItemNo").ToString & ".jpg"
            'If File.Exists(sImg) Then
            'Else
            '    sImg = Windows.Forms.Application.StartupPath & "\blank.jpg"
            'End If
            'Dim fsImageFile As New System.IO.FileStream(sImg, _
            '                                     FileMode.Open, FileAccess.Read)
            'Dim bytImageData(fsImageFile.Length() - 1) As Byte
            'fsImageFile.Read(bytImageData, 0, bytImageData.Length)
            'fsImageFile.Close()

            Dim IMG_DATA As New SqlParameter("@Photo", SqlDbType.Image, _
                bytImageData.Length, ParameterDirection.Input, False, _
               0, 0, Nothing, DataRowVersion.Current, bytImageData)
            cmd.Parameters.Add(IMG_DATA)
            cmd.Parameters.Add(UOMList)
            cmd.Parameters.Add(PriceList)

            cmd.ExecuteNonQuery()

        End While
        rs.Close()
        strSql = "Select * from ItemRep order by ItemRep.ItemNo,unitprice * baseqty"
        'strSql = "Select ROW_NUMBER() OVER(ORDER BY ItemRep.ItemNo DESC) AS 'RNO',* from ItemRep"
        ExecuteReport(strSql, "ItemCatalogRep")
    End Sub

    Private Sub cbCategory_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbCategory.SelectedIndexChanged
        LoadItem()
    End Sub

    Private Sub cbSupplier_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbSupplier.SelectedIndexChanged
        LoadItem()
    End Sub
End Class