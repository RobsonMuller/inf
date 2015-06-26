VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "InfinityControl.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   4485
   ClientLeft      =   120
   ClientTop       =   150
   ClientWidth     =   5190
   ControlBox      =   0   'False
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin rdActiveText.ActiveText txtEmpresa 
      Height          =   315
      Left            =   975
      TabIndex        =   3
      Top             =   3285
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      Alignment       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   2
      Text            =   "01"
      RawText         =   0
      FontName        =   "Times New Roman"
      FontSize        =   9
   End
   Begin VB.CommandButton cmdEntrar 
      Caption         =   "&Entrar"
      Height          =   345
      Left            =   2325
      TabIndex        =   2
      Top             =   3240
      Visible         =   0   'False
      Width           =   870
   End
   Begin rdActiveText.ActiveText txtNome 
      Height          =   315
      Left            =   975
      TabIndex        =   0
      Top             =   3660
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   556
      ForeColor       =   -2147483642
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   20
      RawText         =   0
      FontName        =   "Times New Roman"
      FontSize        =   9
   End
   Begin rdActiveText.ActiveText txtSenha 
      Height          =   315
      Left            =   975
      TabIndex        =   1
      Top             =   4035
      Width           =   2880
      _ExtentX        =   5080
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
   Begin VB.Image Image1 
      Height          =   2745
      Left            =   465
      Picture         =   "Login.frx":000C
      Top             =   120
      Width           =   4215
   End
   Begin VB.Image ImgFechar 
      Height          =   240
      Left            =   4800
      Picture         =   "Login.frx":28C8
      Top             =   105
      Width           =   240
   End
   Begin VB.Image Img_EntrarOn 
      Height          =   315
      Left            =   3945
      Picture         =   "Login.frx":2A12
      Stretch         =   -1  'True
      Top             =   4035
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Image img_EntrarOff 
      Height          =   315
      Left            =   3945
      Picture         =   "Login.frx":3AA3
      Stretch         =   -1  'True
      Top             =   4035
      Width           =   1065
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      BorderStyle     =   6  'Inside Solid
      X1              =   180
      X2              =   5010
      Y1              =   3150
      Y2              =   3150
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Senha:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   180
      TabIndex        =   6
      Top             =   4050
      Width           =   825
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Usuário:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   180
      TabIndex        =   5
      Top             =   3675
      Width           =   825
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Empresa:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   180
      TabIndex        =   4
      Top             =   3300
      Width           =   825
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsErro As INF_Erro.Funcoes

Private Sub cmdEntrar_Click()
   On Error GoTo cmdEntrar_Click_E
   
   Dim clsMD5 As sMD5
   Dim strSenha As String
   Dim strConnect As String
   Dim clsCursor As INF_Cursor.Cursor
   
   Prj.Sistema.IdEmpresa = Me.txtEmpresa
   
   strConnect = "Provider=sqloledb;"
   strConnect = strConnect & "Persist Security Info=False;"
   strConnect = strConnect & "User ID=" & GetSetting(NM_APP, "Connect", "User") & ";"
   strConnect = strConnect & "Password=" & GetSetting(NM_APP, "Connect", "Password") & ";"
   strConnect = strConnect & "Initial Catalog=" & GetSetting(NM_APP, "Connect", "Base") & ";"
   strConnect = strConnect & "Data Source=" & f1.NmComp & "\" & GetSetting(NM_APP, "Connect", "Server") & ";"
   
   Prj.Servidor.ConnectionString = strConnect
   
   'mMsgInfo strConnect
   
   If Not Prj.Sistema.ConexActive Then
      If Not clsConexao.Abrir(strConnect) Then
         clsErro.Transferir = clsConexao.TransferirErro
         Exibir clsErro, "cmdEntrar_Click"
         End
      End If
   End If
   
   Prj.Sistema.ConexActive = True
   Set clsCursor = CreateObject("INF_Cursor.Cursor")
   With clsCursor
      .Inicializar clsConexao
      
      .SQL.Limpar
      .SQL.Mais " SELECT Codigo, RazaoSocial, NomeFantasia, CPFCNPJ, DtMovimento, Ativo "
      .SQL.Mais " FROM Empresas "
      .SQL.Mais " WHERE Codigo = " & .Txt(Prj.Sistema.IdEmpresa)
      
      If Not .Abrir(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         Exibir clsErro, "cmdEntrar_Click"
         GoTo DestruirObjetos
      End If
      
      If Not .EOF Then
         If .Valor("Ativo") = "N" Then
            mMsgAten "Empresa desativada! " & vbNewLine & "Entre em contato com o administrador do sistema."
            mFocus Me.txtEmpresa
            GoTo DestruirObjetos
         End If
         
         Prj.Sistema.DtMovimento = .Valor("DtMovimento")
         Prj.Sistema.CPFCNPJ = .Valor("CPFCNPJ")
         Prj.Sistema.NomeFantasia = .Valor("NomeFantasia")
         Prj.Sistema.RazaoSocial = .Valor("RazaoSocial")
      
      Else
         mMsgInfo "Empresa não localizada! Verifique."
         mFocus Me.txtEmpresa
         GoTo DestruirObjetos
      End If
      .Fechar
   End With
   
   With clsCursor
      .SQL.Limpar
      .SQL.Mais " SELECT Empresa, Codigo, Usuario, Senha, IdSituacao,  "
      .SQL.Mais "    IdAlterarSenha "
      .SQL.Mais " FROM Usuarios "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND Usuario = " & .Txt(Me.txtNome)
      
      If Not .Abrir(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         Exibir clsErro, "cmdEntrar_Click"
         GoTo DestruirObjetos
      End If
      
      If Not .EOF Then
         Set clsMD5 = New sMD5
         
         If LCase(clsMD5.DigestStrToHexStr(Me.txtSenha)) = LCase(.Valor("Senha")) Then
         
            Select Case .Valor("IdSituacao")
            Case 1   'Ativo
               Prj.Sistema.IdUsuario = .Valor("Codigo")
               Unload Me
               MDIInfinity.Show
            Case 2   'Inativo
               mMsgInfo "Usuário desativado! Verifique."
               mFocus Me.txtNome
               GoTo DestruirObjetos
            Case 3   'Bloqueado
               mMsgInfo "Usuário bloqueado! " & vbNewLine & _
                  "Entre em contato com o Administrador do sistema."
               mFocus Me.txtNome
               GoTo DestruirObjetos
            End Select
            
         Else
            mMsgInfo "Senha inválida! Verifique."
            mFocus Me.txtSenha
            GoTo DestruirObjetos
         End If
      Else
         mMsgInfo "Usuário inválido! Verifique."
         mFocus Me.txtNome
         GoTo DestruirObjetos
      End If
      
      .Fechar
   End With
   
   GoTo DestruirObjetos
   
cmdEntrar_Click_E:
   clsErro.Salvar Err
   Exibir clsErro, "cmdEntrar_Click"

DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Sub

Private Sub Form_Load()
   Set clsErro = CreateObject("INF_Erro.Funcoes")
   If ModoDesenvolvimento Then
      Me.txtNome = "master"
      Me.txtSenha = "Master"
      mFocus Me.cmdEntrar
   End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   img_EntrarOff.Visible = True
   Img_EntrarOn.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set clsErro = Nothing
End Sub

Private Sub Image2_Click()
   frmAlterarSenha.Show
End Sub

Private Sub ImgFechar_Click()
   If GetSetting(NM_APP, "Connect", "SavePassword") = "N" Then DeleteSetting NM_APP, "Connect", "Password"
   Unload Me
End Sub

Private Sub img_EntrarOff_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   img_EntrarOff.Visible = False
   Img_EntrarOn.Visible = True
End Sub

Private Sub Img_EntrarOn_Click()
    cmdEntrar_Click
End Sub

Private Sub txtEmpresa_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then cmdEntrar_Click
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then cmdEntrar_Click
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then cmdEntrar_Click
End Sub
