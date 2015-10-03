Public Class Nodo

    Private name As String
    Private inner As String

    Public Property sName() As String
        Get
            Return Name
        End Get
        Set(ByVal Value As String)
            name = Value
        End Set
    End Property

    Public Property sInner() As String
        Get
            Return inner
        End Get
        Set(ByVal Value As String)
            inner = Value
        End Set
    End Property


End Class
