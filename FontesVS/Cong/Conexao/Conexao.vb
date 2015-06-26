Public Class Conexao
    Public SQL As New SQLBuffer
    Public FB As New FuncDB
    Private clsErro As New Erro

    Private strAux As String
    Private strConnectionString As String
    Private ConexOpen As Boolean
    Private TransOpen As Boolean
    Private ADOCon As ADODB.Connection
    Private ADOCommand As ADODB.Command

    Public ReadOnly Property ConnectionString() As String
        Get
            Return strConnectionString
        End Get
    End Property

    Public ReadOnly Property Connect() As Object
        Get
            Return ADOCon
        End Get
    End Property

    Public Function Abrir(Connect As String) As Boolean
        On Error GoTo Abrir_E

        Abrir = False
        ADOCon = New ADODB.Connection
        strConnectionString = Connect
        ADOCon.ConnectionString = strConnectionString
        ADOCon.Open()

        ConexOpen = True
        Abrir = True

        Exit Function

Abrir_E:
        clsErro.Salvar(Err)
        clsErro.ModRotina = "Abrir"
    End Function

    Public Sub Begin()
        If Not TransOpen Then ADOCon.BeginTrans()
        TransOpen = True
    End Sub

    Public Sub Commit()
        ADOCon.CommitTrans()
        TransOpen = False
    End Sub

    Public Sub RollBack()
        On Error Resume Next
        ADOCon.RollbackTrans()
        TransOpen = False
    End Sub

    Public Function Executar(CmdSQL As String) As Boolean
        On Error GoTo Executar_E

        Executar = False

        ADOCommand = CreateObject("ADODB.Command")
        ADOCommand.ActiveConnection = ADOCon
        ADOCommand.Prepared = True
        ADOCommand.CommandText = CmdSQL
        ADOCommand.Execute()

        Executar = True

        Exit Function

Executar_E:
        clsErro.Salvar(Err)
        clsErro.ModRotina = "Executar"
    End Function

    Public ReadOnly Property TransferirErro() As Object
        Get
            Return clsErro.Transferir
        End Get
    End Property

    Public Function Vlr(Valor As Object, Optional ComVirgula As Boolean = False) As String

        Vlr = Replace(CStr(Valor), ",", ".") & IIf(ComVirgula, ",", "")
    End Function

    Public Function Txt(Texto As String, Optional ComVirgula As Boolean = False) As String
        Txt = "'" & Texto & "'"
        If ComVirgula Then Txt = Txt & ", "
    End Function

    Public ReadOnly Property Dt(Data As Object, Optional ComVirgula As Boolean = False, Optional ComHora As Boolean = False) As String
        Get
            If Not IsDate(Data) Then
                strAux = "CONVERT(datetime, NULL)" & IIf(ComVirgula, ", ", "")
            ElseIf IsDate(Data) Then
                If Year(Data) < 1900 Then
                    strAux = "CONVERT(datetime, NULL)" & IIf(ComVirgula, ", ", "")
                Else
                    If ComHora Then
                        strAux = "{ts '" & Format$(Data, "yyyy-mm-dd hh:mm:ss") & "'}" & IIf(ComVirgula, ", ", "")
                    Else
                        strAux = "{d '" & Format(Data, "yyyy-mm-dd") & "'}" & IIf(ComVirgula, ", ", "")
                    End If
                End If
            End If
            Return strAux
        End Get
    End Property

    Public Sub Fechar()
        If ConexOpen Then ADOCon.Close()
    End Sub
End Class
