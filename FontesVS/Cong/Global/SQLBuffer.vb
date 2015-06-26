Public Class SQLBuffer
    Private strSQL As String

    Public Sub Mais(Texto As String)
        strSQL = strSQL & Texto
    End Sub

    Public Sub Limpar()
        strSQL = ""
    End Sub

    Public Function Texto() As String
        Texto = strSQL
    End Function
End Class
