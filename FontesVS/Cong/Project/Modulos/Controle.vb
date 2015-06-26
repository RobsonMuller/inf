Module Controle
    Private strEmp As String
    Private lngCodUsuario As Long
    Private intNivelAcesso As Integer

    Public Property CodUsuario() As Long
        Get
            Return lngCodUsuario
        End Get
        Set(value As Long)
            lngCodUsuario = value
        End Set
    End Property

    Public Property NivelAcesso() As Integer
        Get
            Return intNivelAcesso
        End Get
        Set(value As Integer)
            intNivelAcesso = value
        End Set
    End Property

    Public Property Empresa() As String
        Get
            Return strEmp
        End Get
        Set(value As String)
            strEmp = value
        End Set
    End Property
End Module
