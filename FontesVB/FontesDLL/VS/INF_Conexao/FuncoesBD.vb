Public Class FuncoesBD
    Public Function SeNulo(Campo As String, Valor As String, Optional ComVirgula As Boolean = False) As String
        SeNulo = "ISNULL(" & Campo & ", " & Valor & ")"
        If ComVirgula Then SeNulo = SeNulo & ", "
    End Function

    Public Function DataServer(Optional ComVirgula As Boolean = False) As String
        DataServer = "GETDATE()"
        If ComVirgula Then DataServer = DataServer & ", "
    End Function
End Class
