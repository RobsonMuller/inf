VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "InfinityControl.ocx"
Begin VB.Form frmUsuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Usuários"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7305
   Icon            =   "Usuarios.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   7305
   Tag             =   "010101"
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Novo"
      Height          =   345
      Left            =   5925
      TabIndex        =   10
      Top             =   90
      Width           =   1320
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "&Consultar"
      Height          =   345
      Left            =   5925
      TabIndex        =   11
      Top             =   465
      Width           =   1320
   End
   Begin VB.CommandButton cmdSalvar 
      Cancel          =   -1  'True
      Caption         =   "&Salvar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5925
      TabIndex        =   12
      Top             =   840
      Width           =   1320
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5925
      TabIndex        =   13
      Top             =   1215
      Width           =   1320
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "&Limpar"
      Height          =   345
      Left            =   5925
      TabIndex        =   14
      Top             =   1590
      Width           =   1320
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   345
      Left            =   5925
      TabIndex        =   15
      Top             =   1965
      Width           =   1320
   End
   Begin VB.Frame fraParametros 
      Caption         =   "Parâmetros"
      Enabled         =   0   'False
      Height          =   2535
      Left            =   60
      TabIndex        =   3
      Top             =   660
      Width           =   5775
      Begin rdActiveText.ActiveText txtLogin 
         Height          =   315
         Left            =   1050
         TabIndex        =   5
         Top             =   615
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   20
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtNome 
         Height          =   315
         Left            =   1050
         TabIndex        =   4
         Top             =   240
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   60
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin VB.ComboBox cmbPerfil 
         Height          =   315
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1365
         Width           =   2055
      End
      Begin VB.ComboBox cmbSituacao 
         Height          =   315
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1740
         Width           =   1575
      End
      Begin VB.ComboBox cmbAlterarSenha 
         Height          =   315
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Alterar senha no próximo acesso."
         Top             =   2115
         Width           =   1335
      End
      Begin rdActiveText.ActiveText txtSenha 
         Height          =   315
         Left            =   1050
         TabIndex        =   6
         Top             =   990
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   556
         BackColor       =   -2147483644
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   20
         PasswordChar    =   "*"
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
         Locked          =   -1  'True
      End
      Begin VB.Label lblNome 
         Caption         =   "Nome:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label lblID 
         Caption         =   "ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   645
         Width           =   975
      End
      Begin VB.Label lblSenha 
         Caption         =   "Senha:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label lblPerfil 
         Caption         =   "Perfil:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1395
         Width           =   735
      End
      Begin VB.Label lblSituacao 
         Caption         =   "Situação:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1770
         Width           =   1215
      End
      Begin VB.Label lblProxLogin 
         Caption         =   "Prox. Login:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2190
         Width           =   2295
      End
   End
   Begin VB.Frame fraIdentificacao 
      Caption         =   "Identificação"
      Height          =   660
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin rdActiveText.ActiveText vlrCod 
         Height          =   315
         Left            =   900
         TabIndex        =   1
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   10
         Text            =   "0"
         TextMask        =   3
         RawText         =   3
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin VB.CommandButton cmdPesq 
         Caption         =   "..."
         Height          =   255
         Left            =   2055
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   315
         Width           =   255
      End
      Begin VB.Label lblCodigo 
         Caption         =   "Código:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   270
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsErro As INF_Erro.Funcoes

Private Sub cmdConsultar_Click()
   On Error GoTo cmdConsultar_Click_E
   
   Dim clsCursor As INF_Cursor.Cursor
   
   If Not mCmpObrigatorio(clsErro, Me.vlrCod, "Código") Then Exit Sub
   
   Ampulheta True
   
   Set clsCursor = CreateObject("INF_Cursor.Cursor")
   With clsCursor
      .Inicializar clsConexao
      
      .SQL.Limpar
      .SQL.Mais " SELECT Empresa, Codigo, Usuario, Nome, Senha, IdSituacao, IdPerfil, IdAlterarSenha "
      .SQL.Mais " FROM Usuarios "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND Codigo = " & .Vlr(Me.vlrCod)
      
      If Not .Abrir(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         Exibir clsErro, "cmdConsultar_Click"
         GoTo DestruirObjetos
      End If
      
      If Not .EOF Then
         Me.txtLogin = .Valor("Usuario")
         Me.txtNome = .Valor("Nome")
         Me.cmbSituacao.ListIndex = f1.CmbValor(Me.cmbSituacao, .Valor("IdSituacao"))
         Me.cmbPerfil.ListIndex = f1.CmbValor(Me.cmbPerfil, .Valor("IdPerfil"))
         Me.cmbAlterarSenha.ListIndex = f1.CmbValor(Me.cmbAlterarSenha, .Valor("IdAlterarSenha"), enDescricao, 1)
         
         Me.fraIdentificacao.Enabled = False
         Me.fraParametros.Enabled = True
         
         Me.cmdNovo.Enabled = False
         Me.cmdConsultar.Enabled = False
         Me.cmdSalvar.Caption = "&Alterar"
         
         If Not HabilitarBotao(clsErro, Me, Me.cmdSalvar) Then Exibir clsErro, "cmdConsultar_Click"
         If Not HabilitarBotao(clsErro, Me, Me.cmdExcluir) Then Exibir clsErro, "cmdConsultar_Click"
         
         mFocus Me.txtNome
      Else
         mMsgInfo "Usuário não cadastrado! Verifique."
         mFocus Me.vlrCod
      End If
      
      .Fechar
    End With
    
    GoTo DestruirObjetos

cmdConsultar_Click_E:
   clsErro.Salvar Err
   Exibir clsErro, "cmdConsultar_Click"

DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
   Ampulheta False
End Sub

Private Sub cmdExcluir_Click()
   On Error GoTo cmdExcluir_Click_E
   
   If Not mMsgPerg("Você deseja realmente excluir o registro atual?") Then Exit Sub
      
   'Salva o usuário na tabela de Excluidos
   clsConexao.Begin
   With clsConexao
      .SQL.Limpar
      .SQL.Mais " INSERT INTO UsuariosExcluidos "
      .SQL.Mais "    ( Empresa, Codigo, Nome, Usuario, IdSituacao, IdPerfil, DtHrExclusao, CodUsuarioExclusao )"
      .SQL.Mais " SELECT Empresa, Codigo, Nome, Usuario, IdSituacao, IdPerfil, "
      .SQL.Mais "    GETDATE(), " & .Vlr(Prj.Sistema.IdUsuario)
      .SQL.Mais " FROM Usuarios "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND Codigo = " & .Vlr(Me.vlrCod)
      
      If Not .Executar(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         clsConexao.RollBack
         Exibir clsErro, "cmdExcluir_Click"
         Exit Sub
      End If
   End With
   
   'Exclui o usuário da tabela de usuário
   With clsConexao
      .SQL.Limpar
      .SQL.Mais " DELETE FROM Usuarios "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND Codigo = " & .Vlr(Me.vlrCod)

      If Not .Executar(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         clsConexao.RollBack
         Exibir clsErro, "cmdExcluir_Click"
         Exit Sub
      End If
   End With
   clsConexao.Commit
   
   mMsgInfo "Usuário" & Me.txtNome & " excluido com sucesso!"
   cmdLimpar_Click
   
   Exit Sub

cmdExcluir_Click_E:
   clsErro.Salvar Err
   clsConexao.RollBack
   Exibir clsErro, "cmdExcluir_Click"
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub cmdLimpar_Click()
   f1.Limpar Me
   
   Me.fraIdentificacao.Enabled = True
   Me.fraParametros.Enabled = False
   
   If Not HabilitarBotao(clsErro, Me, Me.cmdNovo) Then Exibir clsErro, "cmdLimpar_Click"
   If Not HabilitarBotao(clsErro, Me, Me.cmdConsultar) Then Exibir clsErro, "cmdLimpar_Click"

   Me.cmbSituacao.ListIndex = f1.CmbValor(Me.cmbSituacao, 1)
   Me.cmbPerfil.ListIndex = f1.CmbValor(Me.cmbPerfil, 2)
   Me.cmbAlterarSenha.ListIndex = f1.CmbValor(Me.cmbAlterarSenha, 1)
   
   Me.cmdExcluir.Enabled = False
   Me.cmdSalvar.Enabled = False
   Me.cmdSalvar.Caption = "&Salvar"
   
   Me.txtSenha.PasswordChar = ""
   
   mFocus Me.vlrCod
End Sub

Private Sub cmdNovo_Click()
   On Error GoTo cmdNovo_Click_E
   
   'Controle dos Frames
   Me.fraIdentificacao.Enabled = False
   Me.fraParametros.Enabled = True
   
   'Controle dos Botões
   Me.cmdConsultar.Enabled = False
   Me.cmdNovo.Enabled = False
   Me.cmdSalvar.Caption = "&Inserir"
   If Not HabilitarBotao(clsErro, Me, Me.cmdSalvar) Then Exibir clsErro, "cmdNovo_Click"
   Me.cmdExcluir.Enabled = False
   
   'Define Senha padrão de Acesso
   Me.txtSenha.PasswordChar = "*"
   Me.txtSenha = "12345"
   
   Me.vlrCod = 0
   
   mFocus Me.txtNome
   
   Exit Sub

cmdNovo_Click_E:
   clsErro.Salvar Err
   Exibir clsErro, "cmdNovo_Click"
End Sub

Private Sub cmdSalvar_Click()
   On Error GoTo cmdSalvar_Click_E
   
   Dim strMsg As String
   Dim lngNewSeq As Long
   Dim clsMD5 As sMD5
   Dim clsCursor As INF_Cursor.Cursor
   
   If Not mCmpObrigatorio(clsErro, Me.txtNome, "Nome") Then GoTo Erro_Msg
   If Not mCmpObrigatorio(clsErro, Me.txtLogin, "ID") Then GoTo Erro_Msg
     
   clsConexao.Begin
   Select Case Me.cmdSalvar.Caption
   Case "&Inserir"
      Set clsCursor = CreateObject("INF_Cursor.Cursor")
      With clsCursor
         .Inicializar clsConexao
         
         .SQL.Limpar
         .SQL.Mais " SELECT ISNULL(Max(Codigo), 0) + 1 AS MaxSeq "
         .SQL.Mais " FROM Usuarios "
         .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
         
         If Not .Abrir(.SQL.Texto) Then
            clsErro.Transferir = .TransferirErro
            clsConexao.RollBack
            Exibir clsErro, "cmdSalvar_Click"
            GoTo DestruirObjetos
         End If
      
         lngNewSeq = .Valor("MaxSeq")
         
         .Fechar
      End With
         
      Set clsMD5 = New sMD5
      With clsConexao
         .SQL.Limpar
         .SQL.Mais " INSERT INTO Usuarios "
         .SQL.Mais "    ( Empresa, Codigo, Usuario, Nome, Senha, IdSituacao, IdPerfil, IdAlterarSenha "
         .SQL.Mais " ) VALUES ( "
         .SQL.Mais .Txt(Prj.Sistema.IdEmpresa, True)
         .SQL.Mais .Vlr(lngNewSeq, True)
         .SQL.Mais .Txt(Me.txtLogin, True)
         .SQL.Mais .Txt(Me.txtNome, True)
         .SQL.Mais .Txt(clsMD5.DigestStrToHexStr("12345"), True)
         .SQL.Mais .Txt(f1.CmbParametro(Me.cmbSituacao), True)
         .SQL.Mais .Txt(f1.CmbParametro(Me.cmbPerfil), True)
         .SQL.Mais .Txt(Left(Me.cmbAlterarSenha.Text, 1))
         .SQL.Mais ")"
         
         strMsg = "Usuário registrado com sucesso!"
      End With
   Case "&Alterar"
      
      lngNewSeq = Me.vlrCod
      
      With clsConexao
         .SQL.Limpar
         .SQL.Mais " UPDATE Usuarios SET "
         .SQL.Mais "    Usuario = " & .Txt(Me.txtLogin, True)
         .SQL.Mais "    Nome = " & .Txt(Me.txtNome, True)
         .SQL.Mais "    IdSituacao = " & .Txt(f1.CmbParametro(Me.cmbSituacao), True)
         .SQL.Mais "    IdPerfil = " & .Txt(f1.CmbParametro(Me.cmbPerfil), True)
         .SQL.Mais "    IdAlterarSenha = " & .Txt(Left(Me.cmbAlterarSenha, 1))
         .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
         .SQL.Mais " AND Codigo = " & .Vlr(lngNewSeq)
         
         strMsg = "Registro alterado com sucesso!"
      End With
   End Select
   
   With clsConexao
      If Not .Executar(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         clsConexao.RollBack
         Exibir clsErro, "cmdSalvar_Click"
         GoTo DestruirObjetos
      End If
   End With
   
   clsConexao.Commit
   
   mMsgInfo strMsg
   cmdLimpar_Click
      
   GoTo DestruirObjetos
   
cmdSalvar_Click_E:
   clsErro.Salvar Err
   clsConexao.RollBack
   Exibir clsErro, "cmdSalvar_Click"
   GoTo DestruirObjetos
   
Erro_Msg:
   Exibir clsErro, "cmdSalvar_Click"
   
DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Sub

Private Sub Form_Load()
   On Error GoTo Form_Load_E
   
   Set clsErro = CreateObject("INF_Erro.Funcoes")
   
   'Centralizar form
   f1.FormCentralizar Me
   
   'Carrega Combo Situacao
   f1.CmbAdd Me.cmbSituacao, "Ativo", 1
   f1.CmbAdd Me.cmbSituacao, "Inativo", 2
   f1.CmbAdd Me.cmbSituacao, "Bloqueado", 3
   Me.cmbSituacao.ListIndex = f1.CmbValor(Me.cmbSituacao, 1)
      
   'Carrega Combo Perfil
   f1.CmbAdd Me.cmbPerfil, "Administrador", 1
   f1.CmbAdd Me.cmbPerfil, "Usuário", 2
   Me.cmbPerfil.ListIndex = f1.CmbValor(Me.cmbPerfil, 2)
   
   'Carrega Combo Alterar Senha
   mCmbSimNao Me.cmbAlterarSenha, False
   
   mFocus Me.vlrCod
   
   Exit Sub

Form_Load_E:
   clsErro.Salvar Err
   Exibir clsErro, "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set clsErro = Nothing
End Sub

