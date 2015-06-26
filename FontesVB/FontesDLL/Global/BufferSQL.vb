Public Class BufferSQL
    Private strSQLAux As String

    Public Function Texto() As String
        Texto = strSQLAux
    End Function

    Public Function Tamanho() As Double
        Tamanho = Len(strSQLAux)
    End Function

    Public Sub Mais(Texto As String)
        strSQLAux = strSQLAux & Texto
    End Sub

    Public Sub Limpar()
        strSQLAux = ""
    End Sub
End Class
