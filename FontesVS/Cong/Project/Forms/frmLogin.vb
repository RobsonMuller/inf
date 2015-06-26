Public Class frmLogin
    Private clsErro As New Erro

    Private Sub cmdEntrar_Click(sender As Object, e As EventArgs) Handles cmdEntrar.Click
        On Error GoTo cmdEntrar_Click_E

        Dim clsCursor As Cursor

        'Verificação de preenchimento
        If Len(Trim(Me.txtEmpresa.Text)) = 0 Then
            MsgBox("O código da empresa deve ser informado! Verifique.", MsgBoxStyle.Information, "Infinity - Informa")
            Me.txtEmpresa.Focus()
            GoTo DestruirObjetos
        End If

        If Len(Trim(Me.txtUsuario.Text)) = 0 Then
            MsgBox("O nome do usuário deve ser informado! Verifique.", MsgBoxStyle.Information, "Infinity - Informa")
            Me.txtUsuario.Focus()
            GoTo DestruirObjetos
        End If

        If Len(Trim(Me.txtSenha.Text)) = 0 Then
            MsgBox("A senha do usuário deve ser informado! Verifique.", MsgBoxStyle.Information, "Infinity - Informa")
            Me.txtSenha.Focus()
            GoTo DestruirObjetos
        End If

        'Verifica Empresa
        clsCursor = New Cursor
        With clsCursor
            .Inicializar(clsConexao.Connect)

            .SQL.Limpar()
            .SQL.Mais(" SELECT Codigo, RazaoSocial, CPFCNPJ, DataFundacao, Situacao ")
            .SQL.Mais(" FROM Empresas ")
            .SQL.Mais(" WHERE Codigo = " & .Txt(Me.txtEmpresa.Text))

            If Not .Abrir(.SQL.Texto) Then
                clsErro.Transferir = .TransferirErro
                Exibir(clsErro, "cmdEntrar_Click")
                GoTo DestruirObjetos
            End If

            If .EOF Then
                MsgBox("Empresa não localizado! Verifique.", MsgBoxStyle.Information, "Infinity - Informa")
                Me.txtUsuario.Focus()
                GoTo DestruirObjetos
            Else
                If UCase(.Valor("Situacao")) = "N" Then
                    MsgBox("Empresa desativado. " & vbNewLine & "Entre em contato com o administrador do sistema.", MsgBoxStyle.Information, "Infinity - Informa")
                    Me.txtUsuario.Focus()
                    GoTo DestruirObjetos
                End If
            End If
            .Fechar()
        End With
        Empresa = Me.txtEmpresa.Text

        'Verifica Usuario
        With clsCursor
            .SQL.Limpar()
            .SQL.Mais(" SELECT Codigo, Nome, Usuario, Senha, NivelAcesso, Situacao ")
            .SQL.Mais(" FROM Usuarios ")
            .SQL.Mais(" WHERE Empresa = " & .Txt(Empresa))
            .SQL.Mais(" AND Usuario = " & .Txt(Me.txtUsuario.Text))

            If Not .Abrir(.SQL.Texto) Then
                clsErro.Transferir = .TransferirErro
                Exibir(clsErro, "cmdEntrar_Click")
                GoTo DestruirObjetos
            End If

            If .EOF Then
                MsgBox("Usuário não localizado! Verifique.", MsgBoxStyle.Information, "Infinity - Informa")
                Me.txtUsuario.Focus()
                GoTo DestruirObjetos
            Else
                CodUsuario = .Valor("Codigo")
                NivelAcesso = .Valor("NivelAcesso")

                If Not Me.txtUsuario.Text = .Valor("Senha") Then
                    MsgBox("Senha inválida! Verifique.", MsgBoxStyle.Information, "Infinity - Informa")
                    Me.txtSenha.Focus()
                    GoTo DestruirObjetos
                End If
            End If
            .Fechar()
        End With

        'Verifica Parametros
        With clsCursor
            .SQL.Limpar()
            .SQL.Mais(" SELECT Empresa ")
            .SQL.Mais(" FROM Parametros ")
            .SQL.Mais(" WHERE Empresa = " & .Txt(Empresa))

            If Not .Abrir(.SQL.Texto) Then
                clsErro.Transferir = .TransferirErro
                Exibir(clsErro, "cmdEntrar_Click")
                GoTo DestruirObjetos
            End If
            .Fechar()
        End With

        'Log de Acesso
        clsConexao.Begin()
        With clsConexao
            .SQL.Limpar()
            .SQL.Mais(" INSERT INTO LogAcesso (Empresa, CodUsuario, DtHr, TpAcesso ")
            .SQL.Mais(" ) VALUES ( ")
            .SQL.Mais(.Txt(Empresa, True))
            .SQL.Mais(.Vlr(CodUsuario, True))
            .SQL.Mais(.FB.DataServer(True))
            .SQL.Mais(.Txt("E"))
            .SQL.Mais(" )")

            If Not .Executar(.SQL.Texto) Then
                clsErro.Transferir = .TransferirErro
                clsConexao.RollBack()
                Exibir(clsErro, "cmdEntrar_Click")
                GoTo DestruirObjetos
            End If
        End With
        clsConexao.Commit()

        Me.Close()
        Me.Dispose()

        GoTo DestruirObjetos

cmdEntrar_Click_E:
        clsErro.Salvar(Err)
        Exibir(clsErro, "cmdEntrar_Click")

DestruirObjetos:
        If Not (clsCursor Is Nothing) Then clsCursor.Fechar()
        clsCursor = Nothing
    End Sub

    Private Sub frmLogin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim strconexao As String

        strConexao = "Provider=SQLOLEDB.1;"
        strConexao = strConexao & "Persist Security Info=True;"
        strConexao = strConexao & "Initial Catalog=HOMOLOGA;"
        strconexao = strconexao & "Data Source=infinity\infinity;"
        strconexao = strconexao & "User Id=sa;"
        strconexao = strconexao & "Password=infinity;"

        If Not clsConexao.Abrir(strconexao) Then
            clsErro.Transferir = clsConexao.TransferirErro
            Exibir(clsErro, "frmLogin_Load")
            End
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        End
    End Sub
End Class
