VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "InfinityControl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEnvioEmail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Envio de Email"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   8205
   Tag             =   "20501"
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "&Enviar"
      Height          =   345
      Left            =   6825
      TabIndex        =   18
      Tag             =   "6"
      Top             =   90
      Width           =   1320
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "&Limpar"
      Height          =   345
      Left            =   6825
      TabIndex        =   17
      Top             =   465
      Width           =   1320
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   345
      Left            =   6825
      TabIndex        =   16
      Top             =   840
      Width           =   1320
   End
   Begin VB.Frame fraAnexos 
      Caption         =   "Anexos:"
      Height          =   1575
      Left            =   60
      TabIndex        =   10
      Top             =   1440
      Width           =   6675
      Begin VB.CommandButton cmdAdicionar 
         Caption         =   "<<"
         Height          =   315
         Left            =   6195
         TabIndex        =   13
         ToolTipText     =   "Adicionar Imagem."
         Top             =   240
         Width           =   345
      End
      Begin VB.CommandButton cmdRemover 
         Caption         =   ">>"
         Height          =   315
         Left            =   6195
         TabIndex        =   12
         ToolTipText     =   "Remover Imagem."
         Top             =   570
         Width           =   345
      End
      Begin MSComctlLib.ListView lstAnexos 
         Height          =   1155
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Caminho"
            Object.Width           =   7937
         EndProperty
      End
   End
   Begin VB.Frame fraMensagem 
      Caption         =   "Mensagem"
      Height          =   3495
      Left            =   60
      TabIndex        =   7
      Top             =   3015
      Width           =   6675
      Begin VB.TextBox txtMensagem 
         Height          =   2745
         Left            =   1215
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   615
         Width           =   5310
      End
      Begin rdActiveText.ActiveText txtAssunto 
         Height          =   315
         Left            =   1215
         TabIndex        =   9
         Top             =   240
         Width           =   5310
         _ExtentX        =   9366
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin VB.Label lblMensagem 
         Caption         =   "Mensagem:"
         Height          =   270
         Left            =   120
         TabIndex        =   14
         Top             =   645
         Width           =   1065
      End
      Begin VB.Label lblAssunto 
         Caption         =   "Assunto:"
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   270
         Width           =   1065
      End
   End
   Begin VB.Frame fraParametros 
      Caption         =   "Parâmetros"
      Height          =   1440
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   6675
      Begin rdActiveText.ActiveText txtRemetente 
         Height          =   315
         Left            =   1215
         TabIndex        =   2
         Top             =   240
         Width           =   5310
         _ExtentX        =   9366
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtDestinatario 
         Height          =   315
         Left            =   1215
         TabIndex        =   4
         Top             =   615
         Width           =   5310
         _ExtentX        =   9366
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtComCopia 
         Height          =   315
         Left            =   1215
         TabIndex        =   6
         Top             =   990
         Width           =   5310
         _ExtentX        =   9366
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin VB.Label lblComCopia 
         Caption         =   "Com Cópia:"
         Height          =   270
         Left            =   120
         TabIndex        =   5
         Top             =   1020
         Width           =   1155
      End
      Begin VB.Label lblDestinatario 
         Caption         =   "Destinatário:"
         Height          =   270
         Left            =   120
         TabIndex        =   3
         Top             =   645
         Width           =   1155
      End
      Begin VB.Label lblRemetente 
         Caption         =   "Remetente:"
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   270
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmEnvioEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intCod As Integer
Private clsErro As INF_Erro.Funcoes

Private Sub cmdEnviar_Click()
   On Error GoTo cmdEnviar_Click_E
   
   Dim objEmail As CDO.Message
   Dim clsCursor As INF_Cursor.Cursor
   
   If Not mCmpObrigatorio(clsErro, Me.txtRemetente, "Remetente") Then GoTo Erro_Obg
   If Not mCmpObrigatorio(clsErro, Me.txtDestinatario, "Destinatário") Then GoTo Erro_Obg
   If Len(Trim(Me.txtAssunto)) = 0 Then
      If Not mMsgPerg("O assunto do e-mail não foi informado, deseja enviar o emaul assim mesmo ?") Then Exit Sub
   End If
   If Not mCmpObrigatorio(clsErro, Me.txtMensagem, "Mensagem") Then GoTo Erro_Obg
   
   'Consulta os dados de configuração de e-mail
   Set objEmail = CreateObject("CDO.Message")
   Set clsCursor = CreateObject("INF_Cursor.Cursor")
   With clsCursor
      .Inicializar clsConexao

      .SQL.Limpar
      .SQL.Mais " SELECT Email, SMTPServerPorta, SMTPServer, SMTPAuthenticateServer, SMTPAuthenticateSSL, UserName, PassWord "
      .SQL.Mais " FROM ConfigEmail "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND CodUser = " & .Vlr(Prj.Sistema.IdUsuario)
   
      If Not .Abrir(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         Exibir clsErro, " cmdEnviar_Click"
         GoTo DestruirObjetos
      End If
      
      If Not .EOF Then
         objEmail.Configuration.Fields(LINK_EMAIL & "smtpserver") = .Valor("SMTPServer")
         objEmail.Configuration.Fields(LINK_EMAIL & "smtpserverport") = .Valor("SMTPServerPorta")
         
         'SendUsing
         'cdoSendUsingPickup = 1; Send the message using the local SMTP service pickup directory.
         'cdoSendUsingPort = 2; Envie a mensagem usando a rede (SMTP através da rede).
         'cdoSendUsingExchange = 3; Envie a mensagem usando o Microsoft Exchange envio de mensagens Uniform Resource Identifier (URI). Este URI é encontrada no usuário urn: schemas: httpmail: pasta caixa de correio sendmsg
         
         objEmail.Configuration.Fields(LINK_EMAIL & "sendusing") = 2
         
         objEmail.Configuration.Fields(LINK_EMAIL & "smtpauthenticate") = IIf(.Valor("SMTPAuthenticateServer") = "S", 1, 0)
         objEmail.Configuration.Fields(LINK_EMAIL & "smtpusessl") = IIf(.Valor("SMTPAuthenticateSSL") = "S", True, False)
         objEmail.Configuration.Fields(LINK_EMAIL & "sendusername") = .Valor("UserName")
         objEmail.Configuration.Fields(LINK_EMAIL & "sendpassword") = .Valor("PassWord")
         objEmail.Configuration.Fields.Update
         objEmail.To = .Valor("Email")
      Else
         mMsgInfo "Nenhuma configuração de e-mail foi encontrada! Verifique."
         mFocus Me.txtRemetente
         GoTo DestruirObjetos
      End If
      .Fechar
   End With
   
   objEmail.From = Me.txtDestinatario
   objEmail.CC = Me.txtComCopia
   
   objEmail.Subject = Me.txtAssunto
   objEmail.TextBody = Me.txtMensagem
   
   objEmail.send
   
   GoTo DestruirObjetos

cmdEnviar_Click_E:
   clsErro.Salvar Err
   Exibir clsErro, "cmdEnviar_Click"

Erro_Obg:
   Exibir clsErro, "cmdEnviar_Click"

DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub cmdLimpar_Click()
   f1.Limpar Me
   
   intCod = 0
   Me.lstAnexos.ListItems.Clear
   
   mFocus Me.txtRemetente
End Sub

Private Sub Form_Load()
   Set clsErro = CreateObject("INF_Erro.Funcoes")
   
   f1.FormCentralizar Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set clsErro = Nothing
End Sub
