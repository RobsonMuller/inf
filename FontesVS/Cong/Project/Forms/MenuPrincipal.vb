Public Class MenuPrincipal
    Private Sub MenuPrincipal_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        frmLogin.ShowDialog()
    End Sub

    Private Sub UsuariosCadastro_Click(sender As Object, e As EventArgs) Handles UsuariosCadastro.Click
        frmUsuarios.Show()
    End Sub
End Class