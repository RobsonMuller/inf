VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "InfinityControl.ocx"
Begin VB.Form frmAlterarSenha 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alteração de Senha"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   6030
   Tag             =   "10103"
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4650
      TabIndex        =   13
      Tag             =   "20102"
      Top             =   1200
      Width           =   1320
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "&Limpar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4650
      TabIndex        =   12
      Tag             =   "20102"
      Top             =   825
      Width           =   1320
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "&Salvar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4650
      TabIndex        =   11
      Tag             =   "20102"
      Top             =   450
      Width           =   1320
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "&Consultar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4650
      TabIndex        =   10
      Top             =   75
      Width           =   1320
   End
   Begin VB.Frame fraParametros 
      Caption         =   "Parâmetros"
      Height          =   1095
      Left            =   60
      TabIndex        =   3
      Top             =   1035
      Width           =   4515
      Begin rdActiveText.ActiveText txtNovaSenha 
         Height          =   315
         Left            =   1380
         TabIndex        =   7
         Top             =   240
         Width           =   2985
         _ExtentX        =   5265
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
         PasswordChar    =   "*"
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtConfirmaSenha 
         Height          =   315
         Left            =   1380
         TabIndex        =   9
         Top             =   630
         Width           =   2985
         _ExtentX        =   5265
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
         PasswordChar    =   "*"
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin VB.Label Label2 
         Caption         =   "Confirma Senha:"
         Height          =   240
         Left            =   120
         TabIndex        =   8
         Top             =   675
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Nova Senha:"
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   270
         Width           =   1185
      End
   End
   Begin VB.Frame fraIdentificacao 
      Caption         =   "Identificação"
      Height          =   1050
      Left            =   60
      TabIndex        =   0
      Top             =   -15
      Width           =   4530
      Begin rdActiveText.ActiveText txtNome 
         Height          =   315
         Left            =   780
         TabIndex        =   2
         Top             =   240
         Width           =   3600
         _ExtentX        =   6350
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
         Locked          =   -1  'True
      End
      Begin rdActiveText.ActiveText txtSenha 
         Height          =   315
         Left            =   780
         TabIndex        =   6
         Top             =   630
         Width           =   3600
         _ExtentX        =   6350
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
         PasswordChar    =   "*"
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
         Locked          =   -1  'True
      End
      Begin VB.Label lblSenha 
         Caption         =   "Senha:"
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   660
         Width           =   1320
      End
      Begin VB.Label lblNome 
         Caption         =   "Nome:"
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   270
         Width           =   1320
      End
   End
End
Attribute VB_Name = "frmAlterarSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsErro As INF_Erro.Funcoes

Private Sub cmdConsultar_Click()
   On Error GoTo cmdConsultar_Click_E
   
   Dim clsMD5 As sMD5
   Dim clsCursor As INF_Cursor.Cursor
   
   If Len(Trim(Me.txtSenha)) = 0 Then
      mMsgInfo "O campo senha deve ser informado! Verifique."
      mFocus Me.txtSenha
      Exit Sub
   End If
   
   Set clsCursor = New INF_Cursor.Cursor
   With clsCursor
      .Inicializar clsConexao
      
      .SQL.Limpar
      .SQL.Mais " SELECT Senha "
      .SQL.Mais " FROM Usuarios "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND Codigo = " & .Vlr(Prj.Sistema.IdUsuario)
      
      If Not .Abrir(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         Exibir clsErro, "cmdConsultar_Click"
         GoTo DestruirObjetos
      End If
      
      If Not .EOF Then
         Set clsMD5 = New sMD5
         If LCase(clsMD5.DigestStrToHexStr(Me.txtSenha)) = LCase(.Valor("Senha")) Then
            If Not HabilitarBotao(clsErro, Me, Me.cmdSalvar) Then Exibir clsErro, "cmdConsultar_Click"
            Me.cmdConsultar.Enabled = False
            Me.fraIdentificacao.Enabled = False
            Me.fraParametros.Enabled = True
            mFocus Me.txtNovaSenha
         Else
            mMsgInfo "Senha inválida! Verifique."
            mFocus Me.txtSenha
         End If
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
   Set clsMD5 = Nothing
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub cmdLimpar_Click()
   f1.Limpar Me
   
   Me.fraIdentificacao.Enabled = True
   Me.fraParametros.Enabled = False
   
   If Not HabilitarBotao(clsErro, Me, Me.cmdConsultar) Then Exibir clsErro, "cmdLimpar_Click"
   Me.cmdSalvar.Enabled = False
   
   mFocus Me.txtSenha
End Sub

Private Sub cmdSalvar_Click()
   On Error GoTo cmdSalvar_Click_E
   
   Dim clsMD5 As sMD5
   
   If Not mCmpObrigatorio(clsErro, Me.txtNovaSenha, "Nova Senha") Then GoTo Erro_Msg
   If Not mCmpObrigatorio(clsErro, Me.txtConfirmaSenha, "Confirma Senha") Then GoTo Erro_Msg

   If Me.txtNovaSenha <> Me.txtConfirmaSenha Then
      mMsgInfo "A confirmação senha não confere! Verifique."
      mFocus Me.txtConfirmaSenha
      Exit Sub
   End If
   
   Set clsMD5 = New sMD5
   clsConexao.Begin
   With clsConexao
      .SQL.Limpar
      .SQL.Mais " UPDATE Usuarios SET "
      .SQL.Mais "    Senha = " & .Txt(clsMD5.DigestStrToHexStr(Me.txtNovaSenha))
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND Codigo = " & .Vlr(Prj.Sistema.IdUsuario)
      
      If Not .Executar(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         clsConexao.RollBack
         Exibir clsErro, "cmdSalvar_Click"
         GoTo DestruirObjetos
      End If
   End With
   clsConexao.Commit
   
   mMsgInfo "Senha atualizada com sucesso!"
   
   cmdLimpar_Click
   
   GoTo DestruirObjetos
   
cmdSalvar_Click_E:
   clsErro.Salvar Err
   Exibir clsErro, "cmdSalvar_Click"
   GoTo DestruirObjetos
   
Erro_Msg:
   Exibir clsErro, "cmdSalvar_Click"
   
DestruirObjetos:
   Set clsMD5 = Nothing
End Sub

Private Sub Form_Load()
   On Error GoTo Form_Load
   
   Dim clsCursor As INF_Cursor.Cursor
   
   Set clsErro = New INF_Erro.Funcoes
   Set clsCursor = New INF_Cursor.Cursor
   With clsCursor
      .Inicializar clsConexao
      
      .SQL.Limpar
      .SQL.Mais " SELECT Nome "
      .SQL.Mais " FROM Usuarios "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND Codigo = " & .Vlr(Prj.Sistema.IdUsuario)
      
      If Not .Abrir(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         Exibir clsErro, "Form_Load"
         GoTo DestruirObjetos
      End If
      
      Me.txtNome = .Valor("Nome")
      .Fechar
   End With
   
   GoTo DestruirObjetos
   
Form_Load:
   clsErro.Salvar Err
   Exibir clsErro, "Form_Load"

DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set clsErro = Nothing
End Sub

