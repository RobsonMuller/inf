Public Class Cursor
    Public SQL As New SQLBuffer
    Private clsErro As New Erro
    Private blnCursorOpen As Boolean
    Private ADOConex As ADODB.Connection
    Private ADORst As ADODB.Recordset

    Public ReadOnly Property TransferirErro() As Object
        Get
            Return clsErro.Transferir
        End Get
    End Property

    Public Sub Inicializar(Conexao As Object)
        ADOConex = Conexao
        ADORst = New ADODB.Recordset
    End Sub

    Public Sub Fechar()
        If blnCursorOpen Then
            ADORst.Close()
            blnCursorOpen = False
        End If
    End Sub

    Public Function Abrir(Texto As String) As Boolean
        On Error GoTo Abrir_E

        ADORst.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        ADORst.CursorType = ADODB.CursorTypeEnum.adOpenStatic
        ADORst.LockType = ADODB.LockTypeEnum.adLockBatchOptimistic
        ADORst.ActiveConnection = ADOConex
        ADORst.Open(Texto)

        blnCursorOpen = True
        Abrir = True

        Exit Function

Abrir_E:
        clsErro.Salvar(Err)
        clsErro.ModRotina = "Abrir"
    End Function

    Public Function Txt(Texto As String, Optional ComVirgula As Boolean = False) As String
        Txt = "'" & Texto & "'"
        If ComVirgula Then Txt = Txt & ", "
    End Function

    Public Function Vlr(Valor As VariantType, Optional ComVirgula As Boolean = False) As VariantType
        Vlr = Replace(Valor, ",", ".")
        If ComVirgula Then Vlr = Vlr & ", "
    End Function

    Public Function EOF() As Boolean
        EOF = ADORst.EOF
    End Function

    Public Sub ProximoRegistro()
        ADORst.MoveNext()
    End Sub

    Public Function Valor(Campo As String) As Object
        Valor = ADORst.Fields(Campo).Value
    End Function
End Class
