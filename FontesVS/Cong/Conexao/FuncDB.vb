Public Class FuncDB
    Public Function DataServer(Optional ComVirgula As Boolean = False) As String
        DataServer = "GETDATE()" & IIf(ComVirgula, ", ", "")
    End Function
End Class
