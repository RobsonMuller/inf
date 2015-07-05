VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "InfinityControl.ocx"
Begin VB.Form frmProdutosFornMod 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fornecedores"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   Icon            =   "ProdutosFornMod.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   5115
      TabIndex        =   7
      Top             =   1815
      Width           =   1320
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   345
      Left            =   3735
      TabIndex        =   6
      Top             =   1815
      Width           =   1320
   End
   Begin VB.Frame fraParametros 
      Caption         =   "Parâmetros"
      Height          =   1770
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   6390
      Begin rdActiveText.ActiveText txtFone 
         Height          =   300
         Left            =   1260
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   990
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
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
         MaxLength       =   14
         TextMask        =   5
         RawText         =   5
         Mask            =   "(###)####-####"
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
         Locked          =   -1  'True
      End
      Begin VB.CommandButton cmdPesqForn 
         Caption         =   "..."
         Height          =   255
         Left            =   2445
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   315
         Width           =   270
      End
      Begin rdActiveText.ActiveText vlrCod 
         Height          =   315
         Left            =   1260
         TabIndex        =   1
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
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
         Text            =   "0"
         TextMask        =   3
         RawText         =   3
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtDescricao 
         Height          =   315
         Left            =   1260
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   615
         Width           =   5010
         _ExtentX        =   8837
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
         MaxLength       =   80
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
         Locked          =   -1  'True
      End
      Begin rdActiveText.ActiveText vlrCompra 
         Height          =   315
         Left            =   1260
         TabIndex        =   5
         Top             =   1350
         Width           =   1710
         _ExtentX        =   3016
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
         MaxLength       =   13
         Text            =   "0,00"
         TextMask        =   4
         RawText         =   4
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin VB.Label lblCompra 
         Caption         =   "Valor Compra:"
         Height          =   300
         Left            =   120
         TabIndex        =   11
         Top             =   1380
         Width           =   1410
      End
      Begin VB.Label lblTelefone 
         Caption         =   "Telefone:"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1020
         Width           =   1125
      End
      Begin VB.Label lblDescricao 
         Caption         =   "Descrição:"
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   645
         Width           =   1185
      End
      Begin VB.Label lblCodFornecedor 
         Caption         =   "Código:"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   270
         Width           =   1110
      End
   End
End
Attribute VB_Name = "frmProdutosFornMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private blnCancelado As Boolean
Private clsErro As INF_Erro.Funcoes

Private Sub cmdCancelar_Click()
   Unload Me
End Sub

Private Sub cmdOk_Click()
   Me.Hide
End Sub

Private Sub cmdPesqForn_Click()
   Dim frmModal As frmCG
   
   Set frmModal = New frmCG
   With frmModal
      .Codigo = Me.vlrCod
      .TpDefinicao = enFornecedores
      .Ativo = True
      .Show vbModal
      
      If Not .Cancelado Then Me.vlrCod = .Codigo
   End With
   
DestruirObjetos:
   If Not (frmModal Is Nothing) Then Unload frmModal
   Set frmModal = Nothing
End Sub

Private Sub Form_Load()
   Set clsErro = CreateObject("INF_Erro.Funcoes")
End Sub

Public Property Get Cancelado() As Boolean
   Cancelado = blnCancelado
End Property

Public Property Get Codigo() As Long
   Codigo = Me.vlrCod
End Property

Public Property Get Descricao() As String
   Descricao = Me.txtDescricao
End Property

Public Property Get Telefone() As String
   Telefone = Me.txtFone
End Property

Public Property Get ValorCompra() As Currency
   ValorCompra = Me.vlrCompra
End Property

Private Sub Form_Unload(Cancel As Integer)
   blnCancelado = True
   Me.Hide
End Sub

Private Sub vlrCod_LostFocus()
   On Error GoTo vlrCod_LostFocus_E
   
   Dim clsCursor As INF_Cursor.Cursor
   
   If Me.vlrCod = 0 Then GoTo DestruirObjetos
   
   Set clsCursor = CreateObject("INF_Cursor.Cursor")
   With clsCursor
      .Inicializar clsConexao
      
      .SQL.Limpar
      .SQL.Mais " SELECT RazaoSocial, Telefone "
      .SQL.Mais " FROM Fornecedores "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND Codigo = " & .Vlr(Me.vlrCod)
      .SQL.Mais " AND Situacao  = " & .Txt("1")
      
      If Not .Abrir(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         clsErro.ModRotina = "VerificaFornecedor"
         GoTo DestruirObjetos
      End If
      
      If Not .EOF Then
         Me.txtDescricao = .Valor("RazaoSocial")
         Me.txtFone = .Valor("Telefone")
         mFocus Me.vlrCompra
      Else
         mMsgInfo "Fornecedor não localizado! Verifique."
         mFocus Me.vlrCod
      End If
      .Fechar
   End With
   
   GoTo DestruirObjetos
   
vlrCod_LostFocus_E:
   clsErro.Salvar Err
   Exibir clsErro, "vlrCod_LostFocus"

DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Sub
