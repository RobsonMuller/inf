Module Project
    Public clsConexao As New Conexao

    Public Sub Main()
        Dim MnPrinc As New MenuPrincipal

        MnPrinc.Show()
    End Sub

    Public Sub Exibir(objErr As Object, Rotina As String)
        Dim strMsg As String

        If objErr.Transferir.NroPrj > 0 Then
            MsgBox(objErr.Transferir.DscPrj, vbInformation, "Infinity - Informativo")
        Else
            strMsg = "Error Number: " & objErr.Transferir.NroVB & vbNewLine
            strMsg = strMsg & "Description: " & objErr.Transferir.DscVB & vbNewLine
            strMsg = strMsg & "Mod. Project: " & Rotina & IIf(Len(Trim(objErr.Transferir.ModRotPrj)) > 0, "\", "") & objErr.Transferir.ModRotPrj

            MsgBox(strMsg, vbCritical, "Infinity - Erro")
        End If
    End Sub
End Module
