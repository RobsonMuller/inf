VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "InfinityControl.ocx"
Begin VB.Form frmConnection 
   BorderStyle     =   0  'None
   Caption         =   "Connection SQL Server"
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   10305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbSalvarSenha 
      Height          =   315
      Left            =   2175
      TabIndex        =   9
      Top             =   1905
      Width           =   855
   End
   Begin rdActiveText.ActiveText txtServidor 
      Height          =   315
      Left            =   900
      TabIndex        =   4
      Top             =   345
      Width           =   2130
      _ExtentX        =   3757
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
      TextCase        =   1
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtBanco 
      Height          =   315
      Left            =   900
      TabIndex        =   5
      Top             =   735
      Width           =   2130
      _ExtentX        =   3757
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
   Begin rdActiveText.ActiveText txtUsuario 
      Height          =   315
      Left            =   900
      TabIndex        =   6
      Top             =   1125
      Width           =   2130
      _ExtentX        =   3757
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
   Begin rdActiveText.ActiveText txtSenha 
      Height          =   315
      Left            =   900
      TabIndex        =   7
      Top             =   1515
      Width           =   2130
      _ExtentX        =   3757
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
      PasswordChar    =   "*"
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin VB.Image cmdEntrar 
      Height          =   240
      Left            =   9915
      Picture         =   "Connection.frx":0000
      Top             =   3375
      Width           =   240
   End
   Begin VB.Image cmdFechar 
      Height          =   240
      Index           =   1
      Left            =   9945
      Picture         =   "Connection.frx":014A
      Top             =   120
      Width           =   240
   End
   Begin VB.Label lblSalvarSenha 
      BackStyle       =   0  'Transparent
      Caption         =   "Salvar Senha Acesso:"
      Height          =   225
      Left            =   120
      TabIndex        =   8
      Top             =   1935
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C000&
      X1              =   3240
      X2              =   3240
      Y1              =   270
      Y2              =   3465
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Senha:"
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   1545
      Width           =   780
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário:"
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   1155
      Width           =   780
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Servidor:"
      Height          =   225
      Left            =   120
      TabIndex        =   1
      Top             =   375
      Width           =   810
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Banco:"
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   765
      Width           =   780
   End
   Begin VB.Image img 
      Height          =   3750
      Index           =   0
      Left            =   0
      Picture         =   "Connection.frx":0294
      Top             =   0
      Width           =   10350
   End
End
Attribute VB_Name = "frmConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsErro As INF_Erro.Funcoes

Private Sub cmdEntrar_Click()

   If Not mCmpObrigatorio(clsErro, Me.txtServidor, "Servidor") Then
      Exibir clsErro, "cmdEntrar_Click"
      Exit Sub
   End If
   
   If Not mCmpObrigatorio(clsErro, Me.txtBanco, "Banco") Then
      Exibir clsErro, "cmdEntrar_Click"
      Exit Sub
   End If
   
   If Not mCmpObrigatorio(clsErro, Me.txtUsuario, "Usuário") Then
      Exibir clsErro, "cmdEntrar_Click"
      Exit Sub
   End If
   
   If Not mCmpObrigatorio(clsErro, Me.txtSenha, "Senha") Then
      Exibir clsErro, "cmdEntrar_Click"
      Exit Sub
   End If
   
   'Salva as informações no perfil do usuário
   SaveSetting NM_APP, "Connect", "Server", Me.txtServidor
   SaveSetting NM_APP, "Connect", "Base", Me.txtBanco
   SaveSetting NM_APP, "Connect", "User", Me.txtUsuario
   SaveSetting NM_APP, "Connect", "Password", Me.txtSenha
   SaveSetting NM_APP, "Connect", "SavePassword", Left(Me.cmbSalvarSenha, 1)
   
   frmLogin.Show
   Unload Me
   
End Sub

Private Sub cmdFechar_Click(Index As Integer)
   End
End Sub

Private Sub Form_Load()
   On Error GoTo Form_Load_E
   
   Set clsErro = New INF_Erro.Funcoes
   
   mCmbSimNao Me.cmbSalvarSenha
   
   If Len(Trim(GetSetting(NM_APP, "Connect", "Server"))) = 0 Then Exit Sub
   
   Me.txtServidor = GetSetting(NM_APP, "Connect", "Server")
   Me.txtBanco = GetSetting(NM_APP, "Connect", "Base")
   Me.txtUsuario = GetSetting(NM_APP, "Connect", "User")
   Me.txtSenha = GetSetting(NM_APP, "Connect", "Password")
   
   Me.cmbSalvarSenha.ListIndex = f1.CmbValor(Me.cmbSalvarSenha, GetSetting(NM_APP, "Connect", "SavePassword"), 1, 1)
   
   Exit Sub

Form_Load_E:
   clsErro.Salvar Err
   Exibir clsErro, "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set clsErro = Nothing
End Sub

Private Sub txtBanco_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then cmdEntrar_Click
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then cmdEntrar_Click
End Sub

Private Sub txtServidor_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then cmdEntrar_Click
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then cmdEntrar_Click
End Sub
