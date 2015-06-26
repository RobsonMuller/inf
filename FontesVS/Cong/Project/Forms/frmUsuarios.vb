Public Class frmUsuarios
    Private clsErro As New Erro

    Private Sub cmdConsultar_Click(sender As Object, e As EventArgs) Handles cmdConsultar.Click
        On Error GoTo cmdConsultar_Clik_E

        Dim clsCursor As Cursor

        If Len(Trim(Me.vlrCod.Text)) Then
            MsgBox("O código do usuário deve ser informado! Verifique.", MsgBoxStyle.Information, "Infinity - Informa")
            Me.vlrCod.Focus()
            GoTo DestruirObjetos
        End If

        clsCursor = New Cursor
        With clsCursor
            .Inicializar(clsConexao)

            .SQL.Limpar()
            .SQL.Mais(" SELECT Empresa, Codigo, Nome, Usuario, Senha, NivelAcesso, Situacao ")
            .SQL.Mais(" FROM Usuarios ")
            .SQL.Mais(" WHERE Empresa = " & .Txt(Empresa))
            .SQL.Mais(" AND Codigo = " & .Vlr(CodUsuario))

            If Not .Abrir(.SQL.Texto) Then
                clsErro.Transferir = .TransferirErro
                Exibir(clsErro, "cmdConsultar_Click")
                GoTo DestruirObjetos
            End If

            If Not .EOF Then
                Me.txtNome = .Valor("Nome")
                Me.txtUsuario = .Valor("Usuario")
            End If
        End With

        Exit Sub

cmdConsultar_Clik_E:
        clsErro.Salvar(Err)
        Exibir(clsErro, "cmdConsultar_Click")

DestruirObjetos:
        If Not (clsCursor Is Nothing) Then clsCursor.Fechar()
        clsCursor = Nothing
    End Sub

    Private Sub frmUsuarios_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

    End Sub
End Class