VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "InfinityControl.ocx"
Begin VB.Form frmRelMarcas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Marcas"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   Icon            =   "RelMarcas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1260
   ScaleWidth      =   5280
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   345
      Left            =   3885
      TabIndex        =   8
      Top             =   840
      Width           =   1320
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "&Limpar"
      Height          =   345
      Left            =   3885
      TabIndex        =   7
      Top             =   465
      Width           =   1320
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   345
      Left            =   3885
      TabIndex        =   6
      Tag             =   "6"
      Top             =   90
      Width           =   1320
   End
   Begin VB.Frame fraParametros 
      Caption         =   "Parâmetros"
      Height          =   1185
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   3750
      Begin VB.ComboBox cmbAtivo 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   615
         Width           =   1125
      End
      Begin VB.CommandButton cmdPesq 
         Caption         =   "..."
         Height          =   255
         Left            =   2025
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   315
         Width           =   270
      End
      Begin rdActiveText.ActiveText vlrCod 
         Height          =   315
         Left            =   840
         TabIndex        =   2
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
      Begin VB.Label lblAtivo 
         Caption         =   "Ativo:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   645
         Width           =   1230
      End
      Begin VB.Label lblCodigo 
         Caption         =   "Código:"
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   270
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmRelMarcas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsErro As INF_Erro.Funcoes

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub cmdImprimir_Click()
   On Error GoTo cmdImprimir_Click_E
   
   Dim clsPrint As INF_Print.Print
   Dim objFrmCry As New frmCrystal
   
   Set clsPrint = CreateObject("INF_Print.Print")
   With clsPrint
      If Not .Inicializar(clsConexao, Prj.Sistema.IdEmpresa, "Marcas.rpt", objFrmCry) Then
         clsErro.Transferir = .TransferirErro
         Exibir clsErro, "cmdImprimir_Click"
         GoTo DestruirObjetos
      End If
      
      .SQL.Limpar
      .SQL.Mais " SELECT Empresa, Codigo, Descricao, Abreviatura, "
      .SQL.Mais "    (CASE WHEN Situacao = 'S' THEN 'Ativado' ELSE 'Desativado' END) AS Situacao "
      .SQL.Mais " FROM Marcas "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      
      If Me.vlrCod > 0 Then .SQL.Mais " AND Codigo = " & .Vlr(Me.vlrCod)
      If f1.CmbParametro(Me.cmbAtivo) <> 99 Then .SQL.Mais " AND Situacao = " & .Txt(Left(Me.cmbAtivo.Text, 1))
      
      If Not .Imprimir() Then
         clsErro.Transferir = .TransferirErro
         Exibir clsErro, "cmdImprimir_Click"
         GoTo DestruirObjetos
      End If
   End With

   GoTo DestruirObjetos

cmdImprimir_Click_E:
   clsErro.Salvar Err
   Exibir clsErro, "cmdImprimir_Click"

DestruirObjetos:
   Set clsPrint = Nothing
End Sub

Private Sub cmdLimpar_Click()
   Me.vlrCod = 0
   Me.cmbAtivo.ListIndex = f1.CmbValor(Me.cmbAtivo, 99)
   mFocus Me.vlrCod
End Sub

Private Sub Form_Load()
   f1.FormCentralizar Me
   
   Set clsErro = CreateObject("INF_Erro.Funcoes")
   
   mCmbSimNao Me.cmbAtivo
   f1.CmbAdd Me.cmbAtivo, "TODOS", 99 'VALOR ADICIONAL
   Me.cmbAtivo.ListIndex = 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set clsErro = Nothing
End Sub
