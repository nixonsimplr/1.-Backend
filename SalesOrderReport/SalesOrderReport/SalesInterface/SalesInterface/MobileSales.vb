Public Class MobileSales
    Public Interface ISalesBase
        Sub MoveFirstClick()
        Sub MovePreviousClick()
        Sub MoveNextClick()
        Sub MoveLastClick()
        Sub ListViewClick()
        Sub MovePositionClick(ByVal Position As Long)
        Sub Print()
        Sub Print(ByVal PageNo As Integer)
        Sub SetCulture(ByVal CultureName As String)
        Sub ReturnedData(ByVal Value As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer)
        Sub ReturnSearch(ByVal SQL As String)
        Function GetListViewForm() As String
        Event LoadForm(ByVal Location As String, ByVal LoadType As String)
        Event LoadDataForm(ByVal Location As String, ByVal LoadType As String, ByVal ParentLoadType As String, ByVal ResultTo As String, ByVal XPos As Integer, ByVal YPos As Integer)
        Event ResultData(ByVal ChildLoadType As String, ByVal Value As String)
        Event SearchField(ByVal FormName As String, ByVal FieldName As String, ByVal FieldType As String, ByVal CurValue As String)
    End Interface
End Class
