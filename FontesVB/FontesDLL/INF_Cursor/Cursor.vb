Public Class Cursor
    Public SQL As New BufferSQL
    Private ADOCon As ADODB.Connection
    Private ADORst As New ADODB.Recordset

    Public Sub Inicializar(Conexao As Object)
        ADOCon = Conexao
    End Sub

    Public Function Abrir(Texto As String) As Boolean
        On Error GoTo Abrir_E

        ADORst.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        ADORst.CursorType = ADODB.CursorTypeEnum.adOpenStatic
        ADORst.LockType = ADODB.LockTypeEnum.adLockBatchOptimistic

        ADORst.ActiveConnection = ADOCon
        ADORst.DataSource = Texto
        ADORst.Open()

        Abrir = True

Abrir_E:
        MsgBox(Err.Description, MsgBoxStyle.Critical, "Infinity - Erro")
    End Function
End Class
