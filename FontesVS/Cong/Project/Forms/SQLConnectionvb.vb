Public Class frmSQLConnection

    Private Sub btnAvancar_Click(sender As Object, e As EventArgs) Handles btnAvancar.Click
        On Error GoTo btnAvancar_Click_E

        'Verificação
        If Len(Trim(Me.txtInstancia.Text)) = 0 Then
            MsgBox("O campo Instância deve ser informado! Verifique.", MsgBoxStyle.Information, "Innfinity - Informa.")
            Me.txtInstancia.Focus()
            Exit Sub
        End If

        If Len(Trim(Me.txtBanco.Text)) = 0 Then
            MsgBox("O campo Banco deve ser informado! Verifique.", MsgBoxStyle.Information, "Innfinity - Informa.")
            Me.txtBanco.Focus()
            Exit Sub
        End If

        If Len(Trim(Me.txtUsuario.Text)) = 0 Then
            MsgBox("O campo Usuário deve ser informado! Verifique.", MsgBoxStyle.Information, "Innfinity - Informa.")
            Me.txtUsuario.Focus()
            Exit Sub
        End If

        If Len(Trim(Me.txtSenha.Text)) = 0 Then
            MsgBox("O campo Senha deve ser informado! Verifique.", MsgBoxStyle.Information, "Innfinity - Informa.")
            Me.txtSenha.Focus()
            Exit Sub
        End If
        '-=-=-=-=-=-'

        SaveSetting(Application.ProductName, "Connect", "Server", Me.txtInstancia.Text)
        SaveSetting(Application.ProductName, "Connect", "Base", Me.txtBanco.Text)
        SaveSetting(Application.ProductName, "Connect", "User", Me.txtUsuario.Text)
        SaveSetting(Application.ProductName, "Connect", "Password", Me.txtSenha.Text)


        frmLogin.Show()
        
        Exit Sub

btnAvancar_Click_E:
        MsgBox(Err.Description, MsgBoxStyle.Critical, "Infinity - Erro.")
    End Sub

End Class