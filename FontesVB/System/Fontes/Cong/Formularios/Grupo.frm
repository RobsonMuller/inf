VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "InfinityControl.ocx"
Begin VB.Form frmGrupo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Grupo"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6180
   Icon            =   "Grupo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2355
   ScaleWidth      =   6180
   Tag             =   "20101"
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Novo"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4785
      TabIndex        =   6
      Tag             =   "20104"
      Top             =   90
      Width           =   1320
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "&Consultar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4785
      TabIndex        =   7
      Top             =   465
      Width           =   1320
   End
   Begin VB.CommandButton cmdSalvar 
      Cancel          =   -1  'True
      Caption         =   "&Salvar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4785
      TabIndex        =   8
      Top             =   840
      Width           =   1320
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4785
      TabIndex        =   9
      Tag             =   "6"
      Top             =   1215
      Width           =   1320
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "&Limpar"
      Height          =   345
      Left            =   4785
      TabIndex        =   10
      Top             =   1590
      Width           =   1320
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   345
      Left            =   4785
      TabIndex        =   11
      Top             =   1965
      Width           =   1320
   End
   Begin VB.Frame fraParametros 
      Caption         =   "Parâmetros"
      Enabled         =   0   'False
      Height          =   1665
      Left            =   60
      TabIndex        =   13
      Top             =   645
      Width           =   4650
      Begin VB.ComboBox cmbAtivo 
         Height          =   315
         Left            =   1245
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   990
         Width           =   1035
      End
      Begin rdActiveText.ActiveText txtDescricao 
         Height          =   315
         Left            =   1245
         TabIndex        =   3
         Top             =   240
         Width           =   3285
         _ExtentX        =   5794
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
         MaxLength       =   30
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtAbreviatura 
         Height          =   315
         Left            =   1245
         TabIndex        =   4
         Top             =   615
         Width           =   2490
         _ExtentX        =   4392
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
      Begin VB.Label lblNome 
         Caption         =   "Descrição:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   270
         Width           =   1125
      End
      Begin VB.Label lblAbreviatura 
         Caption         =   "Abreviatura:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   645
         Width           =   1125
      End
      Begin VB.Label lblAtivo 
         Caption         =   "Ativo:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1020
         Width           =   1230
      End
   End
   Begin VB.Frame fraIdentificacao 
      Caption         =   "Identificação"
      Height          =   645
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   4650
      Begin VB.CommandButton cmdPesq 
         Caption         =   "..."
         Height          =   255
         Left            =   1950
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   300
         Width           =   270
      End
      Begin rdActiveText.ActiveText vlrCod 
         Height          =   315
         Left            =   780
         TabIndex        =   1
         Top             =   240
         Width           =   1110
         _ExtentX        =   1958
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
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   270
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsErro As INF_Erro.Funcoes

Private Sub cmdConsultar_Click()
   On Error GoTo cmdConsultar_Click_E
   
   Dim clsCursor As INF_Cursor.Cursor
   
   If Not mCmpObrigatorio(clsErro, Me, Me.vlrCod) Then Exit Sub
   
   Ampulheta True
   
   Set clsCursor = CreateObject("INF_Cursor.Cursor")
   With clsCursor
      .Inicializar clsConexao
      
      .SQL.Limpar
      .SQL.Mais " SELECT Empresa, Codigo, Descricao, Abreviatura, Ativo "
      .SQL.Mais " FROM Grupo "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND Codigo = " & .Vlr(Me.vlrCod)
      
      If Not .Abrir(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         Exibir clsErro, "cmdConsultar_Click"
         GoTo DestruirObjetos
      End If
      
      If Not .EOF Then
         Me.txtDescricao = .Valor("Descricao")
         Me.txtAbreviatura = .Valor("Abreviatura")
         Me.cmbAtivo.ListIndex = f1.CmbValor(Me.cmbAtivo, .Valor("Ativo"), 1, 1)
         
         Me.cmdConsultar.Enabled = False
         Me.cmdNovo.Enabled = False
         
         Me.cmdSalvar.Caption = "&Alterar"
         If Not HabilitarBotao(clsErro, Me, Me.cmdSalvar) Then Exibir clsErro, "cmdConsultar_Click"
         If Not HabilitarBotao(clsErro, Me, Me.cmdExcluir) Then Exibir clsErro, "cmdConsultar_Click"
         
         Me.fraIdentificacao.Enabled = False
         Me.fraParametros.Enabled = True
         mFocus Me.txtDescricao
      Else
         mMsgInfo "Registro não localizado! Verifique."
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
   
   With clsConexao
      .Begin
      
      .SQL.Limpar
      .SQL.Mais " DELETE FROM Grupo "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND Codigo = " & .Vlr(Me.vlrCod)
      
      If Not .Executar(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         clsConexao.RollBack
         Exibir clsErro, "cmdExcluir_Click"
         Exit Sub
      End If
      
      .Commit
   End With
   
   mMsgInfo "Registro excluido com sucesso!"
   
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
   
   Me.cmdSalvar.Caption = "&Salvar"
   Me.cmdSalvar.Enabled = False
   Me.cmdExcluir.Enabled = False
   If Not HabilitarBotao(clsErro, Me, Me.cmdNovo) Then Exibir clsErro, "cmdLimpar_Click"
   If Not HabilitarBotao(clsErro, Me, Me.cmdConsultar) Then Exibir clsErro, "cmdLimpar_Click"
   Me.cmbAtivo.ListIndex = f1.CmbValor(Me.cmbAtivo, 1)
   mFocus Me.vlrCod
End Sub

Private Sub cmdNovo_Click()
   Me.fraIdentificacao.Enabled = False
   Me.fraParametros.Enabled = True
   
   Me.vlrCod = 0
   Me.cmdNovo.Enabled = False
   Me.cmdConsultar.Enabled = False
   Me.cmdSalvar.Caption = "&Inserir"
   If Not HabilitarBotao(clsErro, Me, Me.cmdSalvar) Then Exibir clsErro, "cmdNovo_Click"
   mFocus Me.txtDescricao
End Sub

Private Sub cmdPesq_Click()
   Dim frmModal As frmCG
   
   Set frmModal = New frmCG
   With frmModal
      .Codigo = Me.vlrCod
      .TpDefinicao = enGrupo
      .Show vbModal
      
      If Not .Cancelado Then Me.vlrCod = .Codigo
   End With
   
DestruirObjetos:
   If Not (frmModal Is Nothing) Then Unload frmModal
   Set frmModal = Nothing
End Sub

Private Sub cmdSalvar_Click()
   On Error GoTo cmdSalvar_Click_E
   
   Dim strMsg As String
   Dim lngCod As Long
   Dim clsCursor As INF_Cursor.Cursor
   
   If Not mCmpObrigatorio(clsErro, Me.txtDescricao, "Descrição") Then GoTo Erro_Obg
   If Not mCmpObrigatorio(clsErro, Me.txtAbreviatura, "Abreviatura") Then GoTo Erro_Obg
   
   clsConexao.Begin
   Select Case Me.cmdSalvar.Caption
   Case "&Inserir"
      Set clsCursor = CreateObject("INF_Cursor.Cursor")
      With clsCursor
         .Inicializar clsConexao
         
         .SQL.Limpar
         .SQL.Mais " SELECT IsNull(Max(Codigo), 0) + 1 AS MaxCod "
         .SQL.Mais " FROM Grupos "
         .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
         
         If Not .Abrir(.SQL.Texto) Then
            clsErro.Transferir = .TransferirErro
            clsConexao.RollBack
            Exibir clsErro, "cmdSalvar_Click"
            GoTo DestruirObjetos
         End If
         
         lngCod = .Valor("MaxCod")
         
         .Fechar
      End With
      
      With clsConexao
         .SQL.Limpar
         
         .SQL.Mais " INSERT INTO Grupos ( "
         .SQL.Mais "    Empresa, Codigo, Descricao, Abreviatura, Ativo "
         .SQL.Mais " ) VALUES ( "
         .SQL.Mais .Txt(Prj.Sistema.IdEmpresa, True)
         .SQL.Mais .Vlr(lngCod, True)
         .SQL.Mais .Txt(Me.txtDescricao, True)
         .SQL.Mais .Txt(Me.txtAbreviatura, True)
         .SQL.Mais .Txt(Left(Me.cmbAtivo.Text, 1))
         .SQL.Mais ")"
         
         If Not .Executar(.SQL.Texto) Then
            clsErro.Transferir = .TransferirErro
            clsConexao.RollBack
            Exibir clsErro, "cmdSalvar_Click"
            GoTo DestruirObjetos
         End If
      End With
      
      strMsg = "Registro inserido com sucesso! Código: " & lngCod
   
   Case "&Alterar"
      lngCod = Me.vlrCod
      
      With clsConexao
         .SQL.Limpar
         
         .SQL.Mais " UPDATE Grupos SET "
         .SQL.Mais "    Descricao = " & .Txt(Me.txtDescricao, True)
         .SQL.Mais "    Abreviatura = " & .Txt(Me.txtAbreviatura, True)
         .SQL.Mais "    Ativo = " & .Txt(Left(Me.cmbAtivo.Text, 1))
         .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
         .SQL.Mais " AND Codigo = " & .Vlr(lngCod)
         
         If Not .Executar(.SQL.Texto) Then
            clsErro.Transferir = .TransferirErro
            clsConexao.RollBack
            Exibir clsErro, "cmdSalvar_Click"
            GoTo DestruirObjetos
         End If
      End With
      
      strMsg = "Registro alterado com sucesso!"
   End Select
   clsConexao.Commit
   
   mMsgInfo strMsg
   
   cmdLimpar_Click
   
   GoTo DestruirObjetos

cmdSalvar_Click_E:
   clsErro.Salvar Err
   clsConexao.RollBack
   Exibir clsErro, "cmdSalvar_Click"
   GoTo DestruirObjetos

Erro_Obg:
   Exibir clsErro, "cmdSalvar_Click"

DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Sub

Private Sub Form_Load()
   Set clsErro = CreateObject("INF_Erro.Funcoes")
   
   f1.FormCentralizar Me
   
   If Not HabilitarBotao(clsErro, Me, Me.cmdNovo) Then Exibir clsErro, "Form_Load"
   If Not HabilitarBotao(clsErro, Me, Me.cmdConsultar) Then Exibir clsErro, "Form_Load"
   
   mCmbSimNao Me.cmbAtivo
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set clsErro = Nothing
End Sub
