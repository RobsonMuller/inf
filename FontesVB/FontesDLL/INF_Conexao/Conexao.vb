Public Class Conexao
    Public SQL As New BufferSQL
    Public FB As New FuncoesBD
    Private blnConnecOpen As Boolean
    Private blnTransOpen As Boolean

    Private ADOCon As ADODB.Connection

    Public Function Abrir(strConnect As String) As Boolean
        On Error GoTo Abrir_E

        Abrir = False

        ADOCon.ConnectionString = strConnect
        ADOCon.Open()

        blnConnecOpen = True
        Abrir = True

        Exit Function

Abrir_E:
        MsgBox(Err.Description, MsgBoxStyle.Critical, "Infinity - Erro ")
    End Function

    Public Sub Begin()
        If Not blnTransOpen Then
            ADOCon.BeginTrans()
            blnTransOpen = True
        End If
    End Sub

    Public Sub Commit()
        ADOCon.CommitTrans()
    End Sub

    Public Sub Close()
        If blnConnecOpen Then ADOCon.Close()
    End Sub

    Public Function Vlr(Valor As VariantType, Optional ComVirgula As Boolean = False) As String
        If ((Not IsNumeric(Valor)) Or (Valor = 0)) Then
            Vlr = "NULL"
        Else
            Vlr = Replace(Valor, ",", ".")
        End If
        If ComVirgula Then Vlr = Vlr & ", "
    End Function

    Public Function Txt(Texto As String, Optional ComVirgula As Boolean = False) As String
        If Len(Trim(Texto)) = 0 Then
            Txt = "NULL"
        Else
            Txt = "'" & Texto & "'"
        End If
        If ComVirgula Then Txt = Txt & ", "
    End Function

    Public Function Dt(Data As Date, Optional ComVirgula As Boolean = False) As String
        If (Not IsDate(Data)) Then
            Dt = "NULL"
        Else
            Dt = "{d '" & Format(Data, "yyyy-MM-dd") & "}"
        End If
        If ComVirgula Then Dt = Dt & ", "
    End Function
End Class
