VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "InfinityControl.ocx"
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5325
   ClientLeft      =   15
   ClientTop       =   45
   ClientWidth     =   5685
   ControlBox      =   0   'False
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFechar 
      Height          =   315
      Left            =   5220
      Picture         =   "Login.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4140
      Width           =   405
   End
   Begin VB.CommandButton cmdAcesso 
      Default         =   -1  'True
      Height          =   315
      Left            =   5220
      Picture         =   "Login.frx":0156
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4605
      Width           =   405
   End
   Begin rdActiveText.ActiveText txtEmpresa 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   4140
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   556
      Appearance      =   0
      ForeColor       =   -2147483638
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Código da empresa..."
      RawText         =   0
      FontName        =   "Times New Roman"
      FontSize        =   9
   End
   Begin VB.CommandButton cmdEntrar 
      Caption         =   "&Entrar"
      Height          =   345
      Left            =   75
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   870
   End
   Begin rdActiveText.ActiveText txtSenha 
      Height          =   315
      Left            =   1785
      TabIndex        =   3
      Top             =   4605
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      Appearance      =   0
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
      Text            =   "Senha de Acesso"
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtNome 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   4140
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      Appearance      =   0
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
      Text            =   "Nome do Usuário"
      RawText         =   0
      FontName        =   "Times New Roman"
      FontSize        =   9
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "v1.001.0000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   90
      TabIndex        =   6
      Top             =   4635
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   5370
      Left            =   -15
      Picture         =   "Login.frx":02A0
      Stretch         =   -1  'True
      Top             =   -30
      Width           =   5700
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsErro As INF_Erro.Funcoes

Private Sub cmdAcesso_Click()
   cmdEntrar_Click
End Sub

Private Sub cmdEntrar_Click()
   On Error GoTo cmdEntrar_Click_E
   
   Dim clsMD5 As sMD5
   Dim strSenha As String
   Dim strConnect As String
   Dim clsCursor As INF_Cursor.Cursor
   
   If ModoDesenvolvimento Then
      Me.txtEmpresa = "01"
      Me.txtNome = "master"
      Me.txtSenha = "master"
   End If
   
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
      .SQL.Mais " SELECT Codigo, RazaoSocial, NomeFantasia, CPFCNPJ, DtMovimento, Situacao "
      .SQL.Mais " FROM Empresas "
      .SQL.Mais " WHERE Codigo = " & .Txt(Prj.Sistema.IdEmpresa)
      
      If Not .Abrir(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         Exibir clsErro, "cmdEntrar_Click"
         GoTo DestruirObjetos
      End If
      
      If Not .EOF Then
         If .Valor("Situacao") = "D" Then
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
      .SQL.Mais " SELECT Empresa, Codigo, Usuario, Senha, Situacao, NivelAcesso "
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
         
            Select Case .Valor("Situacao")
            Case "A"   'Ativo
               Prj.Sistema.IdUsuario = .Valor("Codigo")
            Case "I"   'Inativo
               mMsgInfo "Usuário desativado! Verifique."
               mFocus Me.txtNome
               GoTo DestruirObjetos
            Case "B"   'Bloqueado
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
   
   'Registra a entrada
   With clsConexao
      .Begin
      
      .SQL.Limpar
      .SQL.Mais " INSERT INTO LogAcesso (Empresa, CodUsuario, DtHr, Processo "
      .SQL.Mais " ) VALUES ( "
      .SQL.Mais .Txt(Prj.Sistema.IdEmpresa, True)
      .SQL.Mais .Vlr(Prj.Sistema.IdUsuario, True)
      .SQL.Mais .FB.DataServer, True
      .SQL.Mais .Txt("E")
      .SQL.Mais " )"
      
      If Not .Executar(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         clsConexao.RollBack
         Exibir clsErro, "cmdEntrar_Click"
         GoTo DestruirObjetos
      End If
      
      .Commit
   End With
   
   Unload Me
   MDIInfinity.Show
   
   GoTo DestruirObjetos
   
cmdEntrar_Click_E:
   clsErro.Salvar Err
   Exibir clsErro, "cmdEntrar_Click"

DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Sub

Private Sub cmdFechar_Click()
   If GetSetting(NM_APP, "Connect", "SavePassword") = "N" Then DeleteSetting NM_APP, "Connect", "Password"
   Unload Me
End Sub

Private Sub Form_Load()
   Set clsErro = CreateObject("INF_Erro.Funcoes")
   
   Me.txtEmpresa.MaxLength = 0
   Me.txtEmpresa.ForeColor = &H8000000A
   Me.txtEmpresa = "Código da Empresa ..."
   
   Me.txtNome.MaxLength = 0
   Me.txtNome.ForeColor = &H8000000A
   Me.txtNome = "Nome do Usuário ..."
   
   Me.txtSenha.MaxLength = 0
   Me.txtSenha.ForeColor = &H8000000A
   Me.txtSenha.PasswordChar = ""
   Me.txtSenha = "Senha de Acesso ..."
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set clsErro = Nothing
End Sub

Private Sub Image2_Click()
   frmAlterarSenha.Show
End Sub

Private Sub Img_EntrarOn_Click()
    cmdEntrar_Click
End Sub

Private Sub img_fecharOff_Click()

End Sub

Private Sub img_fecharOn_Click()
End Sub

Private Sub txtEmpresa_GotFocus()
   If Me.txtEmpresa.MaxLength = 0 Then
      Me.txtEmpresa = ""
      Me.txtEmpresa.MaxLength = 2
      Me.txtEmpresa.ForeColor = &H80000008
   End If
End Sub

Private Sub txtEmpresa_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then cmdEntrar_Click
End Sub

Private Sub txtEmpresa_LostFocus()
   If Len(Trim(Me.txtEmpresa)) = 0 Then
      Me.txtEmpresa.MaxLength = 0
      Me.txtEmpresa.ForeColor = &H8000000A
      Me.txtEmpresa = "Código da Empresa ..."
   End If
End Sub

Private Sub txtNome_GotFocus()
   If Me.txtNome.MaxLength = 0 Then
      Me.txtNome = ""
      Me.txtNome.MaxLength = 40
      Me.txtNome.ForeColor = &H80000008
   End If
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then cmdEntrar_Click
End Sub

Private Sub txtNome_LostFocus()
   If Len(Trim(Me.txtNome)) = 0 Then
      Me.txtNome.MaxLength = 0
      Me.txtNome.ForeColor = &H8000000A
      Me.txtNome = "Nome do Usuário ..."
   End If
End Sub

Private Sub txtSenha_GotFocus()
   If Me.txtSenha.MaxLength = 0 Then
      Me.txtSenha = ""
      Me.txtSenha.MaxLength = 40
      Me.txtSenha.ForeColor = &H80000008
      Me.txtSenha.PasswordChar = "*"
   End If
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then cmdEntrar_Click
End Sub

Private Sub txtSenha_LostFocus()
   If Len(Trim(Me.txtSenha)) = 0 Then
      Me.txtSenha.MaxLength = 0
      Me.txtSenha.ForeColor = &H8000000A
      Me.txtSenha.PasswordChar = ""
      Me.txtSenha = "Senha de Acesso ..."
   End If
End Sub
