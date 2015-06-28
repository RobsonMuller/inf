VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "InfinityControl.ocx"
Begin VB.Form frmBasicCad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   6960
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   345
      Left            =   5550
      TabIndex        =   15
      Top             =   1965
      Width           =   1320
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "&Limpar"
      Height          =   345
      Left            =   5550
      TabIndex        =   14
      Top             =   1590
      Width           =   1320
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5550
      TabIndex        =   13
      Top             =   1215
      Width           =   1320
   End
   Begin VB.CommandButton cmdSalvar 
      Cancel          =   -1  'True
      Caption         =   "&Salvar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5550
      TabIndex        =   12
      Top             =   840
      Width           =   1320
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "&Consultar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5550
      TabIndex        =   11
      Top             =   465
      Width           =   1320
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Novo"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5550
      TabIndex        =   10
      Top             =   90
      Width           =   1320
   End
   Begin VB.Frame fraParametros 
      Caption         =   "Parâmetros"
      Enabled         =   0   'False
      Height          =   1680
      Left            =   60
      TabIndex        =   3
      Top             =   645
      Width           =   5430
      Begin VB.ComboBox cmbSituacao 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   990
         Width           =   1305
      End
      Begin rdActiveText.ActiveText txtDescricao 
         Height          =   315
         Left            =   1140
         TabIndex        =   5
         Top             =   240
         Width           =   4155
         _ExtentX        =   7329
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtAbreviatura 
         Height          =   315
         Left            =   1140
         TabIndex        =   7
         Top             =   615
         Width           =   2085
         _ExtentX        =   3678
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
      Begin VB.Label lblSituacao 
         Caption         =   "Situacao:"
         Height          =   165
         Left            =   120
         TabIndex        =   9
         Top             =   1020
         Width           =   1050
      End
      Begin VB.Label lblAbreviatura 
         Caption         =   "Abreviatura:"
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   645
         Width           =   1485
      End
      Begin VB.Label lblDescricao 
         Caption         =   "Descrição:"
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   270
         Width           =   1485
      End
   End
   Begin VB.Frame fraIdentificacao 
      Caption         =   "Identificação"
      Height          =   645
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   5430
      Begin rdActiveText.ActiveText vlrCod 
         Height          =   315
         Left            =   780
         TabIndex        =   2
         Top             =   240
         Width           =   1155
         _ExtentX        =   2037
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
      Begin VB.Label lblCod 
         Caption         =   "Código:"
         Height          =   210
         Left            =   105
         TabIndex        =   1
         Top             =   270
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmBasicCad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum EnForm
   enMarca = 0
   enModelo = 1
End Enum

Private strTable As String
Private clsErro As INF_Erro.Funcoes

Private Sub cmdConsultar_Click()
   On Error GoTo cmdConsultar_Click_E

   Dim clsCursor As INF_Cursor.Cursor
   
   If Me.vlrCod = 0 Then
      mMsgInfo "O campo código é de preenchimento obrigatório! Verifique."
      mFocus Me.vlrCod
      GoTo DestruirObjetos
   End If
   
   Set clsCursor = CreateObject("INF_Cursor.Cursor")
   With clsCursor
      .Inicializar clsConexao
      
      .SQL.Limpar
      .SQL.Mais " SELECT Empresa, Codigo, Descricao, Abreviatura, Situacao "
      .SQL.Mais " FROM " & strTable
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND Codigo = " & .Vlr(Me.vlrCod)
      
      If Not .Abrir(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         clsErro.ModRotina = "cmdConsultar_Click"
         GoTo DestruirObjetos
      End If
      
      If Not .EOF Then
         Me.txtDescricao = .Valor("Descricao")
         Me.txtAbreviatura = .Valor("Abreviatura")
         Me.cmbSituacao.ListIndex = f1.CmbValor(Me.cmbSituacao, .Valor("Situacao"), enCodigo)
         
         Me.fraIdentificacao.Enabled = False
         Me.fraParametros.Enabled = True
         
         Me.cmdNovo.Enabled = False
         Me.cmdConsultar.Enabled = False
         
         Me.cmdSalvar.Caption = "&Alterar"
         If Not HabilitarBotao(clsErro, Me, Me.cmdSalvar) Then Exibir clsErro, "cmdConsultar_Click"
         If Not HabilitarBotao(clsErro, Me, Me.cmdExcluir) Then Exibir clsErro, "cmdConsultar_Click"
         
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
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub cmdLimpar_Click()
   f1.Limpar Me
   
   Me.cmdSalvar.Enabled = False
   Me.cmdSalvar.Caption = "&Salvar"
   Me.cmdExcluir.Enabled = False
   
   Me.fraIdentificacao.Enabled = True
   Me.fraParametros.Enabled = False
   
   Me.cmbSituacao.ListIndex = f1.CmbValor(Me.cmbSituacao, 1)
   
   If Not HabilitarBotao(clsErro, Me, Me.cmdNovo) Then Exibir clsErro, "cmdNovo_Click"
   If Not HabilitarBotao(clsErro, Me, Me.cmdConsultar) Then Exibir clsErro, "Form_Load"
   
End Sub

Private Sub cmdNovo_Click()
   Me.fraIdentificacao.Enabled = False
   Me.fraParametros.Enabled = True
   
   Me.cmdNovo.Enabled = False
   Me.cmdConsultar.Enabled = False
   Me.cmdSalvar.Caption = "&Inserir"
   
   If Not HabilitarBotao(clsErro, Me, Me.cmdSalvar) Then Exibir clsErro, "cmdNovo_Click"
End Sub

Private Sub cmdSalvar_Click()
   On Error GoTo cmdSalvar_Click_E
   
   Dim lngSeq As Long
   Dim strMsg As String
   Dim clsCursor As INF_Cursor.Cursor
   
   clsConexao.Begin
   Select Case Me.cmdSalvar.Caption
   Case "&Inserir"
      Set clsCursor = CreateObject("INF_Cursor.Cursor")
      With clsCursor
         .Inicializar clsConexao
         
         .SQL.Limpar
         .SQL.Mais " SELECT (ISNULL(MAX(Codigo), 0) + 1) AS MaxQtd "
         .SQL.Mais " FROM " & strTable
         .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
         
         If Not .Abrir(.SQL.Texto) Then
            clsErro.Transferir = .TransferirErro
            clsConexao.RollBack
            Exibir clsErro, "cmdSalvar_Click"
            GoTo DestruirObjetos
         End If
         
         lngSeq = .Valor("MaxQtd")
         
         .Fechar
      End With
      
      With clsConexao
         .SQL.Limpar
         .SQL.Mais " INSERT INTO " & strTable & " ("
         .SQL.Mais "    Empresa, Codigo, Descricao, Abreviatura, Situacao "
         .SQL.Mais " ) VALUES ( "
         .SQL.Mais .Txt(Prj.Sistema.IdEmpresa, True)
         .SQL.Mais .Vlr(lngSeq, True)
         .SQL.Mais .Txt(Me.txtDescricao, True)
         .SQL.Mais .Txt(Me.txtAbreviatura, True)
         .SQL.Mais .Txt(CStr(f1.CmbParametro(Me.cmbSituacao)))
         .SQL.Mais " )"
         
         If Not .Executar(.SQL.Texto) Then
            clsErro.Transferir = .TransferirErro
            clsConexao.RollBack
            Exibir clsErro, "cmdSalvar_Click"
            GoTo DestruirObjetos
         End If
      End With
      
      strMsg = "Registro inserido com sucesso! Código: " & lngSeq
   Case "&Alterar"
      With clsConexao
         lngSeq = Me.vlrCod
      
         .SQL.Limpar
         .SQL.Mais " UPDATE " & strTable & " SET "
         .SQL.Mais "    Descricao = " & .Txt(Me.txtDescricao, True)
         .SQL.Mais "    Abreviatura = " & .Txt(Me.txtAbreviatura, True)
         .SQL.Mais "    Situacao = " & .Txt(CStr(f1.CmbParametro(Me.cmbSituacao)))
         .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
         .SQL.Mais " AND Codigo = " & .Vlr(lngSeq)
         
         If Not .Executar(.SQL.Texto) Then
            clsErro.Transferir = .TransferirErro
            clsConexao.RollBack
            Exibir clsErro, "cmdSalvar_Click"
            GoTo DestruirObjetos
         End If
      End With
      
      strMsg = "Registro atualizado com sucesso!"
   End Select
   clsConexao.Commit
   
   mMsgInfo strMsg
   cmdLimpar_Click
   
   GoTo DestruirObjetos

cmdSalvar_Click_E:
   clsErro.Salvar Err
   clsConexao.RollBack
   Exibir clsErro, "cmdSalvar_Click"
   
DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Sub

Private Sub Form_Load()
   Set clsErro = New INF_Erro.Funcoes
   
   f1.FormCentralizar Me
   
   f1.CmbAdd Me.cmbSituacao, "Ativado", 1
   f1.CmbAdd Me.cmbSituacao, "Desativado", 2
   Me.cmbSituacao.ListIndex = f1.CmbValor(Me.cmbSituacao, 1, enCodigo)
   
   'Ler comentário:
   'A função habilitabotão deve ser chamada pelo metodo BaseForm
   'E o metodo BaseForm deve ser setado no MDI antes do Show para que o formulário funcione normalmente
   
   mFocus Me.vlrCod
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set clsErro = Nothing
End Sub

Public Sub BaseForm(CodForm As EnForm)

   Select Case CodForm
   Case EnForm.enMarca
      strTable = "Marcas"
      Me.Tag = "020101"
      
   End Select
   
   Me.Caption = "Cadastro de " & strTable
   If Not HabilitarBotao(clsErro, Me, Me.cmdNovo) Then Exibir clsErro, "Form_Load"
   If Not HabilitarBotao(clsErro, Me, Me.cmdConsultar) Then Exibir clsErro, "Form_Load"
End Sub
