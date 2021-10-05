Public Class ComboValues
    Private sDesc As String
    Private sCode As String

    Public Sub New(ByVal Code As String, ByVal Desc As String)
        Me.sCode = Code
        Me.sDesc = Desc
    End Sub 'New

    Public ReadOnly Property Desc() As String
        Get
            Return sDesc
        End Get
    End Property

    Public ReadOnly Property Code() As String
        Get
            Return sCode
        End Get
    End Property
End Class
