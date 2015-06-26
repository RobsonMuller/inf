VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "InfinityControl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRecebeEmail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recebimento de E-mail"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6600
   ScaleWidth      =   9105
   Tag             =   "20502"
   Begin VB.CommandButton cmdResponder 
      Caption         =   "&Responder"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5940
      TabIndex        =   14
      Top             =   615
      Width           =   1455
   End
   Begin VB.Frame fraComandos 
      Caption         =   "Comandos"
      Height          =   2055
      Left            =   5820
      TabIndex        =   11
      Top             =   0
      Width           =   3225
      Begin rdActiveText.ActiveText txtEmails 
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   1620
         Width           =   2985
         _ExtentX        =   5265
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
      Begin VB.CommandButton cmdDeletar 
         Caption         =   "&Deletar"
         Enabled         =   0   'False
         Height          =   345
         Left            =   1650
         TabIndex        =   17
         Top             =   990
         Width           =   1455
      End
      Begin VB.CommandButton cmdAtualizar 
         Caption         =   "&Atualizar"
         Height          =   345
         Left            =   120
         TabIndex        =   16
         Top             =   990
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Responder Todos"
         Enabled         =   0   'False
         Height          =   345
         Left            =   1650
         TabIndex        =   15
         Top             =   615
         Width           =   1455
      End
      Begin VB.CommandButton cmdEncaminhar 
         Caption         =   "&Encaminhar"
         Enabled         =   0   'False
         Height          =   345
         Left            =   1650
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdEnviar 
         Caption         =   "&Enviar"
         Enabled         =   0   'False
         Height          =   345
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.Line Line1 
         X1              =   135
         X2              =   3075
         Y1              =   1500
         Y2              =   1500
      End
   End
   Begin VB.Frame fraEntrada 
      Caption         =   "Caixa de Entrada"
      Height          =   4485
      Left            =   60
      TabIndex        =   9
      Top             =   2055
      Width           =   8985
      Begin MSComctlLib.ListView lstMail 
         Height          =   4110
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   8730
         _ExtentX        =   15399
         _ExtentY        =   7250
         SortKey         =   3
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "From"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Subject"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "SortValue"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.Frame fraIdentificacao 
      Caption         =   "Identificação"
      Height          =   2055
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   5700
      Begin rdActiveText.ActiveText txtUsuario 
         Height          =   315
         Left            =   2625
         TabIndex        =   2
         Top             =   615
         Width           =   2940
         _ExtentX        =   5186
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
      Begin rdActiveText.ActiveText datData 
         Height          =   315
         Left            =   2625
         TabIndex        =   3
         Top             =   240
         Width           =   1155
         _ExtentX        =   2037
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
         MaxLength       =   10
         TextMask        =   1
         RawText         =   1
         Mask            =   "##/##/####"
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
         Locked          =   -1  'True
      End
      Begin rdActiveText.ActiveText txtEmail 
         Height          =   315
         Left            =   2625
         TabIndex        =   6
         Top             =   990
         Width           =   2940
         _ExtentX        =   5186
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
      Begin MSComctlLib.ProgressBar pgBar 
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Top             =   1635
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Verificando ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1905
         TabIndex        =   8
         Top             =   1410
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "E-mail:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1905
         TabIndex        =   5
         Top             =   1020
         Width           =   915
      End
      Begin VB.Label Label7 
         Caption         =   "Data:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1905
         TabIndex        =   4
         Top             =   270
         Width           =   915
      End
      Begin VB.Label Label6 
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1905
         TabIndex        =   1
         Top             =   645
         Width           =   915
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1680
         Left            =   120
         Picture         =   "RecebeEmail.frx":0000
         Top             =   240
         Width           =   1680
      End
   End
End
Attribute VB_Name = "frmRecebeEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsErro As INF_Erro.Funcoes

'Componente de terceiro
Private GetMail As OSPOP3_Plus.Session

Private SMTPServerPorta As String
Private SMTPServer As String
Private SMTPAuthenticateSSL As String
Private SMTPAuthenticateServer As String

Private POPServerPorta As String
Private POPServer As String
Private POPAuthenticateSSL As String
Private POPServerCopy As String
Private POPTypeAuthenticate As String
Private POPProtocolRead As String

Private UserName As String
Private Password As String

Private blnConectado As Boolean

Private Sub Form_Load()
   On Error GoTo Form_Load_E
   
   Dim clsCursor As INF_Cursor.Cursor
   Dim frmModal As frmConnectMail
 
   f1.FormCentralizar Me
   
   Set GetMail = CreateObject("OSPOP3_Plus.Session")
   Set clsCursor = CreateObject("INF_Cursor.Cursor")
   Set clsErro = CreateObject("INF_Erro.Funcoes")
   
   pgBar.Min = 0
   pgBar.Max = 100
   pgBar.value = 0
   
   Me.datData = Date
   blnConectado = False
   
   Me.lblStatus.Caption = "Verificando dados do usuário ..."
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
      
      Me.txtUsuario = .Valor("Nome")
      .Fechar
   End With
   
   'Busca as configurações do e-mail do usuário
   Me.lblStatus.Caption = "Verificando configurações de e-mail ..."
   With clsCursor
      .SQL.Limpar
      .SQL.Mais " SELECT Empresa, CodUser, "
      
      .SQL.Mais "    SMTPServerPorta, SMTPServer, SMTPAuthenticateSSL, "
      .SQL.Mais "    SMTPAuthenticateServer, "
      
      .SQL.Mais "    POPServerPorta, POPServer, POPAuthenticateSSL, "
      .SQL.Mais "    POPServerCopy, POPTypeAuthenticate, POPProtocolRead, "
      
      .SQL.Mais "    UserName, Password, Email "
      
      .SQL.Mais " FROM ConfigEmail "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND CodUser = " & .Vlr(Prj.Sistema.IdUsuario)
      
      If Not .Abrir(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         Exibir clsErro, "Form_Load"
         GoTo DestruirObjetos
      End If
      
      If .EOF Then
         Me.lblStatus.Caption = "Informações de e-mail não localizadas ..."
         mMsgInfo "Nenhuma configuração de e-mail foi realizada até o momento! Verifique."
      Else
         'Envio
         SMTPServerPorta = .Valor("SMTPServerPorta")
         SMTPServer = .Valor("SMTPServer")
         SMTPAuthenticateSSL = .Valor("SMTPAuthenticateSSL")
         SMTPAuthenticateServer = .Valor("SMTPAuthenticateServer")
         
         'Recebimento
         POPServerPorta = .Valor("POPServerPorta")
         POPServer = .Valor("POPServer")
         POPAuthenticateSSL = .Valor("POPAuthenticateSSL")
         POPServerCopy = .Valor("POPServerCopy")
         POPTypeAuthenticate = .Valor("POPTypeAuthenticate")
         POPProtocolRead = .Valor("POPProtocolRead")
         UserName = .Valor("UserName")
         Password = .Valor("Password")
         Me.txtEmail = .Valor("Email")
      End If
      .Fechar
   End With
        
   DoEvents
   Set frmModal = New frmConnectMail
   With frmModal
      
      .POP3 = POPServer
      .Porta = POPServerPorta
      .Usuario = UserName
      .Senha = Password
      .SSL = IIf(POPAuthenticateSSL = "S", True, False)
      
      Me.lblStatus.Caption = "Estabelecendo conexão com o provedor ..."
      If Not .Conectar(clsErro, GetMail) Then Exibir clsErro, "Form_Activate"
      
   End With
      
   GoTo DestruirObjetos

Form_Load_E:
   clsErro.Salvar Err
   Exibir clsErro, "Form_Load"
   
DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Sub
