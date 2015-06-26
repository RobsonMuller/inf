VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "InfinityControl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConfigEmail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuração de Email"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   6525
   Tag             =   "40100"
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   40
      Top             =   6795
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   11466
            MinWidth        =   11466
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5145
      TabIndex        =   21
      Top             =   840
      Width           =   1320
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "&Consultar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5145
      TabIndex        =   19
      Top             =   90
      Width           =   1320
   End
   Begin VB.Frame fraParametros 
      Caption         =   "Parâmetros"
      Enabled         =   0   'False
      Height          =   1410
      Left            =   60
      TabIndex        =   3
      Top             =   660
      Width           =   4995
      Begin rdActiveText.ActiveText txtUserName 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   3690
         _ExtentX        =   6509
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
         MaxLength       =   80
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtSenha 
         Height          =   315
         Left            =   1200
         TabIndex        =   5
         Top             =   615
         Width           =   3690
         _ExtentX        =   6509
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
         MaxLength       =   40
         PasswordChar    =   "*"
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtEmail 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   990
         Width           =   3690
         _ExtentX        =   6509
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
         MaxLength       =   80
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin VB.Label lblEmail 
         Caption         =   "Email:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1020
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   270
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "Senha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   645
         Width           =   1035
      End
   End
   Begin VB.Frame fraIdentificacao 
      Caption         =   "Identificação"
      Height          =   660
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   4995
      Begin VB.CommandButton cmdPesq 
         Caption         =   "..."
         Height          =   255
         Left            =   2070
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   315
         Width           =   255
      End
      Begin rdActiveText.ActiveText vlrCod 
         Height          =   315
         Left            =   915
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
      Begin VB.Label lblCodigo 
         Caption         =   "Código:"
         Height          =   255
         Left            =   135
         TabIndex        =   35
         Top             =   270
         Width           =   795
      End
   End
   Begin VB.Frame fraRead 
      Caption         =   "Read"
      Enabled         =   0   'False
      Height          =   2550
      Left            =   60
      TabIndex        =   28
      Top             =   4230
      Width           =   4995
      Begin VB.ComboBox cmbProtocolRec 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2130
         Width           =   2700
      End
      Begin VB.ComboBox cmbTpAutenticacao 
         Height          =   315
         Left            =   1545
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1740
         Width           =   3330
      End
      Begin VB.ComboBox cmbCopyServer 
         Height          =   315
         Left            =   4050
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1365
         Width           =   825
      End
      Begin VB.ComboBox cmbReadSSL 
         Height          =   315
         Left            =   4050
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   990
         Width           =   825
      End
      Begin rdActiveText.ActiveText vlrPortaRead 
         Height          =   315
         Left            =   1185
         TabIndex        =   13
         Top             =   240
         Width           =   945
         _ExtentX        =   1667
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
         MaxLength       =   5
         Text            =   "0"
         TextMask        =   3
         RawText         =   3
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtPOP 
         Height          =   315
         Left            =   1185
         TabIndex        =   14
         Top             =   615
         Width           =   3690
         _ExtentX        =   6509
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
         MaxLength       =   80
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin VB.Label Label7 
         Caption         =   "Protocolo de Recebimento:"
         Height          =   255
         Left            =   90
         TabIndex        =   34
         Top             =   2145
         Width           =   2430
      End
      Begin VB.Label Label6 
         Caption         =   "Tp. Autenticação:"
         Height          =   255
         Left            =   90
         TabIndex        =   33
         Top             =   1770
         Width           =   1785
      End
      Begin VB.Label Label5 
         Caption         =   "Deixar uma cópia da mensagem no Servidor:"
         Height          =   255
         Left            =   90
         TabIndex        =   32
         Top             =   1395
         Width           =   3945
      End
      Begin VB.Label Label1 
         Caption         =   "Requer autenticação segura (SSL):"
         Height          =   240
         Left            =   90
         TabIndex        =   31
         Top             =   1020
         Width           =   3960
      End
      Begin VB.Label lblPop 
         Caption         =   "POP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   645
         Width           =   1035
      End
      Begin VB.Label lblPortaRead 
         Caption         =   "Porta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   270
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   345
      Left            =   5145
      TabIndex        =   23
      Top             =   1590
      Width           =   1320
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "&Limpar"
      Height          =   345
      Left            =   5145
      TabIndex        =   22
      Top             =   1215
      Width           =   1320
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "&Salvar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5145
      TabIndex        =   20
      Top             =   465
      Width           =   1320
   End
   Begin VB.Frame fraSend 
      Caption         =   "Send"
      Enabled         =   0   'False
      Height          =   2160
      Left            =   60
      TabIndex        =   7
      Top             =   2070
      Width           =   4995
      Begin VB.ComboBox cmbErro 
         Height          =   315
         Left            =   4065
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1740
         Width           =   825
      End
      Begin VB.ComboBox cmbServerAut 
         Height          =   315
         Left            =   4065
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1365
         Width           =   825
      End
      Begin VB.ComboBox cmbSSL 
         Height          =   315
         Left            =   4065
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   990
         Width           =   825
      End
      Begin rdActiveText.ActiveText txtSMTP 
         Height          =   315
         Left            =   1200
         TabIndex        =   9
         Top             =   615
         Width           =   3690
         _ExtentX        =   6509
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
         MaxLength       =   80
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText vlrPorta 
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         Top             =   240
         Width           =   945
         _ExtentX        =   1667
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
         MaxLength       =   5
         Text            =   "0"
         TextMask        =   3
         RawText         =   3
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin VB.Label Label8 
         Caption         =   "Enviar E-mail na Geração de Erro do Sistema:"
         Height          =   255
         Left            =   105
         TabIndex        =   39
         Top             =   1770
         Width           =   3420
      End
      Begin VB.Label lblServerAut 
         Caption         =   "Meu Servidor Requer Autenticação:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1395
         Width           =   3015
      End
      Begin VB.Label lblSSL 
         Caption         =   "Requer autenticação segura (SSL):"
         Height          =   240
         Left            =   120
         TabIndex        =   26
         Top             =   1020
         Width           =   2835
      End
      Begin VB.Label Label2 
         Caption         =   "SMTP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   645
         Width           =   1035
      End
      Begin VB.Label lblSMTPPorta 
         Caption         =   "Porta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   270
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmConfigEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Received As Boolean
Dim Message$
Dim sckError

Private intCod As Integer
Private clsErro As INF_Erro.Funcoes

Private Sub cmdConsultar_Click()
   On Error GoTo cmdConsultar_Click_E
   
   Dim clsCursor As INF_Cursor.Cursor
   
   If Not mCmpObrigatorio(clsErro, Me.vlrCod, "Código") Then GoTo Erro_Obg

   Set clsCursor = CreateObject("INF_Cursor.Cursor")
   With clsCursor
      .Inicializar clsConexao
      
      .SQL.Limpar
      .SQL.Mais " SELECT "
      
      'Identificação
      .SQL.Mais "    Empresa , CodUser, "
      
      'Envio
      .SQL.Mais "    SMTPServerPorta, SMTPServer, SMTPAuthenticateSSL, "
      .SQL.Mais "    SMTPAuthenticateServer, "
      
      'Recebimento
      .SQL.Mais "    POPServerPorta, POPServer, POPAuthenticateSSL, "
      .SQL.Mais "    POPServerCopy, POPTypeAuthenticate, POPProtocolRead, "
      
      'Parametros gerais
      .SQL.Mais "    UserName, Password, Email, SendErro "
      
      .SQL.Mais " FROM ConfigEmail "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND CodUser = " & .Vlr(Me.vlrCod)
      
      If Not .Abrir(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         Exibir clsErro, "cmdConsultar_Click"
         GoTo DestruirObjetos
      End If
      
      If .EOF Then
         'Se não encontrar o registro
         Me.cmdSalvar.Caption = "&Inserir"
      Else
         Me.cmdSalvar.Caption = "&Alterar"
         If Not HabilitarBotao(clsErro, Me, Me.cmdExcluir) Then Exibir clsErro, "cmdConsultar_Click"
         
         Me.vlrPorta = .Valor("SMTPServerPorta")
         Me.txtSMTP = .Valor("SMTPServer")
         Me.cmbSSL.ListIndex = f1.CmbValor(Me.cmbSSL, .Valor("SMTPAuthenticateSSL"), enDescricao, 1)
         Me.cmbServerAut.ListIndex = f1.CmbValor(Me.cmbServerAut, .Valor("SMTPAuthenticateServer"), enDescricao, 1)
         Me.vlrPortaRead = .Valor("POPServerPorta")
         Me.txtPOP = .Valor("POPServer")
         Me.cmbReadSSL.ListIndex = f1.CmbValor(Me.cmbReadSSL, .Valor("POPAuthenticateSSL"), enDescricao, 1)
         Me.cmbCopyServer.ListIndex = f1.CmbValor(Me.cmbCopyServer, .Valor("POPServerCopy"), enDescricao, 1)
         Me.cmbTpAutenticacao.ListIndex = f1.CmbValor(Me.cmbTpAutenticacao, .Valor("POPTypeAuthenticate"))
         Me.cmbProtocolRec.ListIndex = f1.CmbValor(Me.cmbProtocolRec, .Valor("POPProtocolRead"))
         Me.txtUserName = .Valor("UserName")
         Me.txtSenha = .Valor("Password")
         Me.txtEmail = .Valor("Email")
         Me.cmbErro.ListIndex = f1.CmbValor(Me.cmbErro, .Valor("SendErro"), enDescricao, 1)
      End If
      
      If Not HabilitarBotao(clsErro, Me, Me.cmdSalvar) Then Exibir clsErro, "cmdConsultar_Click"
      Me.cmdConsultar.Enabled = False
      Me.fraIdentificacao.Enabled = False
      Me.fraParametros.Enabled = True
      Me.fraRead.Enabled = True
      Me.fraSend.Enabled = True
      
      .Fechar
   End With
   
   GoTo DestruirObjetos
   
cmdConsultar_Click_E:
   clsErro.Salvar Err
   Exibir clsErro, "cmdConsultar_Click"
   GoTo DestruirObjetos
   
Erro_Obg:
   Exibir clsErro, "cmdConsultar_Click"
   
DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub cmdLimpar_Click()
   f1.Limpar Me
   
   Me.cmbServerAut.ListIndex = f1.CmbValor(Me.cmbServerAut, "S", enDescricao, 1)
   Me.cmbSSL.ListIndex = f1.CmbValor(Me.cmbSSL, "S", enDescricao, 1)
   Me.cmbCopyServer.ListIndex = f1.CmbValor(Me.cmbCopyServer, "S", enDescricao, 1)
   Me.cmbErro.ListIndex = f1.CmbValor(Me.cmbErro, "S", enDescricao, 1)
   Me.cmbProtocolRec.ListIndex = f1.CmbValor(Me.cmbProtocolRec, "U", enDescricao, 1)
   Me.cmbTpAutenticacao.ListIndex = f1.CmbValor(Me.cmbTpAutenticacao, "P", enDescricao, 1)
   Me.cmbReadSSL.ListIndex = f1.CmbValor(Me.cmbReadSSL, "S", enDescricao, 1)
   
   If Not HabilitarBotao(clsErro, Me, Me.cmdConsultar) Then Exibir clsErro, "cmdConsultar_Click"
   Me.fraIdentificacao.Enabled = True
   Me.fraParametros.Enabled = False
   Me.fraRead.Enabled = False
   Me.fraSend.Enabled = False
   
   Me.cmdSalvar.Enabled = False
   Me.cmdExcluir.Enabled = False
   
   Me.cmdSalvar.Caption = "&Salvar"
   
   mFocus Me.vlrCod
End Sub

Private Sub cmdSalvar_Click()
   On Error GoTo cmdSalvar_Click_E
   
   If Not mCmpObrigatorio(clsErro, Me.vlrPorta, "Porta SMTP") Then GoTo Erro_Obg
   If Not mCmpObrigatorio(clsErro, Me.txtSMTP, "SMTP") Then GoTo Erro_Obg
   If Not mCmpObrigatorio(clsErro, Me.txtUserName, "Nome do Usuário") Then GoTo Erro_Obg
   If Not mCmpObrigatorio(clsErro, Me.txtSenha, "Senha do Usuário") Then GoTo Erro_Obg
   
   Ampulheta True
   
   With clsConexao
      .Begin

      Select Case Me.cmdSalvar.Caption
      Case "&Alterar"
         .SQL.Limpar
         .SQL.Mais " UPDATE ConfigEmail SET "
         .SQL.Mais "    SMTPServerPorta = " & .Vlr(Me.vlrPorta, True)
         .SQL.Mais "    SMTPServer = " & .Txt(Me.txtSMTP, True)
         .SQL.Mais "    SMTPAuthenticateServer = " & .Txt(Left(Me.cmbServerAut, 1), True)
         .SQL.Mais "    SMTPAuthenticateSSL = " & .Txt(Left(Me.cmbSSL, 1), True)
         .SQL.Mais "    POPServerPorta = " & .Txt(Me.vlrPortaRead, True)
         .SQL.Mais "    POPServer = " & .Txt(Me.txtPOP, True)
         .SQL.Mais "    POPAuthenticateSSL = " & .Txt(Left(Me.cmbReadSSL, 1), True)
         .SQL.Mais "    POPServerCopy = " & .Txt(Left(Me.cmbCopyServer, 1), True)
         .SQL.Mais "    POPTypeAuthenticate = " & .Txt(f1.CmbParametro(Me.cmbTpAutenticacao), True)
         .SQL.Mais "    POPProtocolRead = " & .Txt(f1.CmbParametro(Me.cmbProtocolRec), True)
         .SQL.Mais "    SendErro = " & .Txt(Left(Me.cmbErro, 1), True)
         .SQL.Mais "    UserName = " & .Txt(Me.txtUserName, True)
         .SQL.Mais "    PassWord = " & .Txt(Me.txtSenha, True)
         .SQL.Mais "    Email = " & .Txt(Me.txtEmail)
         .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
         .SQL.Mais " AND CodUser = " & .Vlr(Me.vlrCod)
      Case "&Inserir"
         .SQL.Limpar
         .SQL.Mais " INSERT INTO ConfigEmail ( "
         .SQL.Mais "    Empresa, CodUser, SMTPServerPorta, SMTPServer, SMTPAuthenticateSSL, "
         .SQL.Mais "    SMTPAuthenticateServer, POPServerPorta, POPServer, POPAuthenticateSSL,"
         .SQL.Mais "    POPServerCopy, POPTypeAuthenticate, POPProtocolRead, UserName, PassWord, "
         .SQL.Mais "    Email, SendErro "
         .SQL.Mais " ) VALUES ( "
         .SQL.Mais .Txt(Prj.Sistema.IdEmpresa, True)
         .SQL.Mais .Vlr(Me.vlrCod, True)
         .SQL.Mais .Vlr(Me.vlrPorta, True)
         .SQL.Mais .Txt(Me.txtSMTP, True)
         .SQL.Mais .Txt(Left(Me.cmbSSL, 1), True)
         .SQL.Mais .Txt(Left(Me.cmbServerAut, 1), True)
         .SQL.Mais .Vlr(Me.vlrPortaRead, True)
         .SQL.Mais .Txt(Me.txtPOP, True)
         .SQL.Mais .Txt(Left(Me.cmbReadSSL, 1), True)
         .SQL.Mais .Txt(Left(Me.cmbCopyServer, 1), True)
         .SQL.Mais .Txt(f1.CmbParametro(Me.cmbTpAutenticacao), True)
         .SQL.Mais .Txt(f1.CmbParametro(Me.cmbProtocolRec), True)
         .SQL.Mais .Txt(Me.txtUserName, True)
         .SQL.Mais .Txt(Me.txtSenha, True)
         .SQL.Mais .Txt(Me.txtEmail, True)
         .SQL.Mais .Txt(Left(Me.cmbErro, 1))
         .SQL.Mais " )"
      End Select
      
      If Not .Executar(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         Exibir clsErro, "cmdSalvar_Click"
         GoTo DestruirObjetos
      End If
      
      .Commit
   End With
   
   mMsgInfo "Registro salvo com sucesso! Verifique."
   
   cmdLimpar_Click
   
   GoTo DestruirObjetos

cmdSalvar_Click_E:
   clsErro.Salvar Err
   Exibir clsErro, "cmdSalvar_Click"
   GoTo DestruirObjetos

Erro_Obg:
   Exibir clsErro, "cmdSalvar_Click"

DestruirObjetos:
   Ampulheta False
End Sub

Private Sub Form_Load()
   Set clsErro = CreateObject("INF_Erro.Funcoes")
   
   mCmbSimNao Me.cmbSSL
   mCmbSimNao Me.cmbServerAut
   mCmbSimNao Me.cmbErro
   mCmbSimNao Me.cmbCopyServer
   mCmbSimNao Me.cmbReadSSL
   
   If Not HabilitarBotao(clsErro, Me, Me.cmdConsultar) Then Exibir clsErro, "Form_Load"
   
   f1.FormCentralizar Me
   
   f1.CmbAdd Me.cmbTpAutenticacao, "POP3", 0
   f1.CmbAdd Me.cmbTpAutenticacao, "IMAP4", 1
   f1.CmbAdd Me.cmbTpAutenticacao, "Exchange WebDAV - 2000/2003", 2
   f1.CmbAdd Me.cmbTpAutenticacao, "Exchange Web Service - 2007/2010", 3
   Me.cmbTpAutenticacao.ListIndex = f1.CmbValor(Me.cmbTpAutenticacao, 0)
   
   f1.CmbAdd Me.cmbProtocolRec, "USER/LOGIN", 0
   f1.CmbAdd Me.cmbProtocolRec, "APOP(CRAM-MD5)", 1
   f1.CmbAdd Me.cmbProtocolRec, "NTLM", 2
   Me.cmbProtocolRec.ListIndex = f1.CmbValor(Me.cmbProtocolRec, 0)
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set clsErro = Nothing
End Sub



