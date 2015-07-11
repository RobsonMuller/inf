VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "InfinityControl.ocx"
Begin VB.Form frmFornecedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Fornecedores"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   Icon            =   "Fornecedores.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6405
   ScaleWidth      =   9480
   Tag             =   "20401"
   Begin MSComctlLib.ImageList ImageList 
      Left            =   8850
      Top             =   2340
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Fornecedores.frx":000C
            Key             =   "Removido"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Fornecedores.frx":0166
            Key             =   "Gravado"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Fornecedores.frx":02C0
            Key             =   "Inserido"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Fornecedores.frx":041A
            Key             =   "Alterado"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraContatos 
      Caption         =   "Contatos"
      Height          =   1980
      Left            =   60
      TabIndex        =   24
      Top             =   4365
      Width           =   7965
      Begin VB.CommandButton cmdModificar 
         Caption         =   "="
         Height          =   315
         Left            =   7500
         TabIndex        =   27
         ToolTipText     =   "Adicionar Fornecedor."
         Top             =   585
         Width           =   345
      End
      Begin VB.CommandButton cmdRemover 
         Caption         =   ">>"
         Height          =   315
         Left            =   7500
         TabIndex        =   28
         ToolTipText     =   "Remover Fornecedor."
         Top             =   930
         Width           =   345
      End
      Begin VB.CommandButton cmdAdicionar 
         Caption         =   "<<"
         Height          =   315
         Left            =   7500
         TabIndex        =   26
         ToolTipText     =   "Adicionar Fornecedor."
         Top             =   240
         Width           =   345
      End
      Begin MSComctlLib.ListView lstContatos 
         Height          =   1620
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   2858
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList"
         SmallIcons      =   "ImageList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Contato"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Telefone"
            Object.Width           =   2647
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "E-mail"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   345
      Left            =   8100
      TabIndex        =   34
      Top             =   1965
      Width           =   1320
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "&Limpar"
      Height          =   345
      Left            =   8100
      TabIndex        =   33
      Top             =   1590
      Width           =   1320
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Enabled         =   0   'False
      Height          =   345
      Left            =   8100
      TabIndex        =   32
      Tag             =   "6"
      Top             =   1215
      Width           =   1320
   End
   Begin VB.CommandButton cmdSalvar 
      Cancel          =   -1  'True
      Caption         =   "&Salvar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   8100
      TabIndex        =   31
      Tag             =   "20102"
      Top             =   840
      Width           =   1320
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "&Consultar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   8100
      TabIndex        =   30
      Top             =   465
      Width           =   1320
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Novo"
      Enabled         =   0   'False
      Height          =   345
      Left            =   8100
      TabIndex        =   29
      Top             =   90
      Width           =   1320
   End
   Begin VB.Frame fraParametros 
      Caption         =   "Parâmetros:"
      Enabled         =   0   'False
      Height          =   3705
      Left            =   60
      TabIndex        =   3
      Top             =   660
      Width           =   7965
      Begin rdActiveText.ActiveText vlrCEP 
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         Top             =   990
         Width           =   1125
         _ExtentX        =   1984
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
         MaxLength       =   9
         TextMask        =   6
         RawText         =   6
         Mask            =   "#####-###"
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtEmail 
         Height          =   315
         Left            =   1200
         TabIndex        =   22
         Top             =   2895
         Width           =   6600
         _ExtentX        =   11642
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
      Begin rdActiveText.ActiveText txtTelefone 
         Height          =   315
         Left            =   1200
         TabIndex        =   19
         Top             =   2505
         Width           =   1785
         _ExtentX        =   3149
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
      Begin rdActiveText.ActiveText txtBairro 
         Height          =   315
         Left            =   1200
         TabIndex        =   17
         Top             =   2130
         Width           =   4335
         _ExtentX        =   7646
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
      Begin rdActiveText.ActiveText txtDescCidade 
         Height          =   315
         Left            =   2715
         TabIndex        =   16
         Top             =   1755
         Width           =   5085
         _ExtentX        =   8969
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
      Begin rdActiveText.ActiveText vlrCodCidade 
         Height          =   315
         Left            =   1200
         TabIndex        =   14
         Top             =   1755
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
      Begin rdActiveText.ActiveText vlrNro 
         Height          =   315
         Left            =   6780
         TabIndex        =   13
         Top             =   1380
         Width           =   1020
         _ExtentX        =   1799
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
      Begin rdActiveText.ActiveText datCadastro 
         Height          =   315
         Left            =   6780
         TabIndex        =   5
         Top             =   255
         Width           =   1020
         _ExtentX        =   1799
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
      Begin rdActiveText.ActiveText vlrCPFCNPJ 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   615
         Width           =   2700
         _ExtentX        =   4763
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
         MaxLength       =   14
         TextMask        =   7
         RawText         =   7
         Mask            =   "###.###.###-##"
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin VB.CommandButton cmdPesqCidade 
         Caption         =   "..."
         Height          =   255
         Left            =   2385
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1815
         Width           =   270
      End
      Begin rdActiveText.ActiveText txtNome 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
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
      Begin VB.ComboBox cmbTipoPessoa 
         Height          =   315
         Left            =   3900
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1005
         Width           =   1650
      End
      Begin VB.ComboBox cmbSituacao 
         Height          =   315
         Left            =   6780
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1005
         Width           =   1020
      End
      Begin VB.ComboBox cmbUf 
         Height          =   315
         Left            =   6780
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2130
         Width           =   1020
      End
      Begin VB.CommandButton cmdConsultarCEP 
         Height          =   315
         Left            =   2400
         Picture         =   "Fornecedores.frx":0574
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Consultar CEP"
         Top             =   1005
         Width           =   330
      End
      Begin rdActiveText.ActiveText vlrIE 
         Height          =   315
         Left            =   5100
         TabIndex        =   7
         Top             =   630
         Width           =   2700
         _ExtentX        =   4763
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
         TextMask        =   9
         RawText         =   9
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtEndereco 
         Height          =   315
         Left            =   1200
         TabIndex        =   12
         Top             =   1380
         Width           =   4335
         _ExtentX        =   7646
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
      Begin rdActiveText.ActiveText txtFax 
         Height          =   315
         Left            =   3780
         TabIndex        =   20
         Top             =   2505
         Width           =   1755
         _ExtentX        =   3096
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
      Begin rdActiveText.ActiveText txtCelular 
         Height          =   315
         Left            =   6045
         TabIndex        =   21
         Top             =   2520
         Width           =   1755
         _ExtentX        =   3096
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
         MaxLength       =   14
         TextMask        =   5
         RawText         =   5
         Mask            =   "(###)####-####"
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtSite 
         Height          =   315
         Left            =   1200
         TabIndex        =   23
         Top             =   3270
         Width           =   6600
         _ExtentX        =   11642
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
      Begin VB.Label lblCelular 
         Caption         =   "Cel.:"
         Height          =   195
         Left            =   5640
         TabIndex        =   53
         Top             =   2550
         Width           =   885
      End
      Begin VB.Label lblFax 
         Caption         =   "Fax:"
         Height          =   195
         Left            =   3225
         TabIndex        =   52
         Top             =   2535
         Width           =   885
      End
      Begin VB.Label lblDataCadastro 
         Caption         =   "Dt. Cadastro:"
         Height          =   180
         Left            =   5640
         TabIndex        =   51
         Top             =   300
         Width           =   1395
      End
      Begin VB.Label lblTipoPessoa 
         Caption         =   "Tp. Pessoa:"
         Height          =   255
         Left            =   2910
         TabIndex        =   50
         Top             =   1035
         Width           =   1230
      End
      Begin VB.Label lblSituação 
         Caption         =   "Situação:"
         Height          =   240
         Left            =   5640
         TabIndex        =   49
         Top             =   1035
         Width           =   1050
      End
      Begin VB.Label lblSite 
         Caption         =   "Site:"
         Height          =   240
         Left            =   120
         TabIndex        =   48
         Top             =   3300
         Width           =   1110
      End
      Begin VB.Label lblEmail 
         Caption         =   "E-mail:"
         Height          =   210
         Left            =   120
         TabIndex        =   47
         Top             =   2925
         Width           =   1095
      End
      Begin VB.Label lblTelefone 
         Caption         =   "Telefone:"
         Height          =   195
         Left            =   90
         TabIndex        =   46
         Top             =   2535
         Width           =   1125
      End
      Begin VB.Label lblUF 
         Caption         =   "Estado:"
         Height          =   225
         Left            =   5640
         TabIndex        =   45
         Top             =   2160
         Width           =   990
      End
      Begin VB.Label lblCidade 
         Caption         =   "Cidade:"
         Height          =   300
         Left            =   120
         TabIndex        =   44
         Top             =   1785
         Width           =   1170
      End
      Begin VB.Label lblNumero 
         Caption         =   "Número:"
         Height          =   270
         Left            =   5640
         TabIndex        =   43
         Top             =   1425
         Width           =   1140
      End
      Begin VB.Label lblBairro 
         Caption         =   "Bairro:"
         Height          =   210
         Left            =   120
         TabIndex        =   42
         Top             =   2160
         Width           =   1140
      End
      Begin VB.Label lblEndereco 
         Caption         =   "Endereço:"
         Height          =   210
         Left            =   120
         TabIndex        =   41
         Top             =   1410
         Width           =   1155
      End
      Begin VB.Label lblCEP 
         Caption         =   "CEP"
         Height          =   210
         Left            =   120
         TabIndex        =   40
         Top             =   1035
         Width           =   1200
      End
      Begin VB.Label lblIE 
         Caption         =   "Ins. Estadual:"
         Height          =   225
         Left            =   4005
         TabIndex        =   39
         Top             =   660
         Width           =   1080
      End
      Begin VB.Label lblCNPJ 
         Caption         =   "CNPJ:"
         Height          =   225
         Left            =   120
         TabIndex        =   38
         Top             =   645
         Width           =   1125
      End
      Begin VB.Label lblNome 
         Caption         =   "Nome:"
         Height          =   285
         Left            =   120
         TabIndex        =   37
         Top             =   270
         Width           =   1110
      End
   End
   Begin VB.Frame fraIdentificacao 
      Caption         =   "Identificação"
      Height          =   660
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   7980
      Begin VB.CommandButton cmdPesqForn 
         Caption         =   "..."
         Height          =   255
         Left            =   2385
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   315
         Width           =   270
      End
      Begin rdActiveText.ActiveText vlrCod 
         Height          =   315
         Left            =   1185
         TabIndex        =   1
         Top             =   240
         Width           =   1140
         _ExtentX        =   2011
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
         Height          =   270
         Left            =   120
         TabIndex        =   36
         Top             =   270
         Width           =   1065
      End
   End
   Begin MSComctlLib.ListView lstLegenda 
      Height          =   1065
      Left            =   8100
      TabIndex        =   35
      Top             =   5280
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   1879
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList"
      SmallIcons      =   "ImageList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Descricao"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Label lblLegenda 
      Caption         =   "Legenda"
      Height          =   225
      Left            =   8100
      TabIndex        =   54
      Top             =   5055
      Width           =   1215
   End
End
Attribute VB_Name = "frmFornecedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lngSeqLst As Long
Private clsErro As INF_Erro.Funcoes
Private colContatos As Collection

Private Const CONST_LST_NOME As Integer = 0
Private Const CONST_LST_TELEFONE As Integer = 1
Private Const CONST_LST_EMAIL As Integer = 2

Private Const CONST_ICO_GRAVADO As String = "Gravado"
Private Const CONST_ICO_INSERIDO As String = "Inserido"
Private Const CONST_ICO_REMOVIDO As String = "Removido"
Private Const CONST_ICO_ALTERADO As String = "Alterado"

Private Sub cmbTipoPessoa_Click()
   
   Me.vlrIE.TextMask = [Custom Mask]
   Me.vlrCPFCNPJ.TextMask = [Custom Mask]
   
   If f1.CmbIndex(Me.cmbTipoPessoa) = 0 Then
      Me.lblCNPJ.Caption = "CPF:"
      Me.lblIE.Caption = "RG:"
      Me.vlrCPFCNPJ.Mask = "###.###.###-##"
      Me.vlrIE.Mask = "##.##.##.##-##"
   Else
      Me.lblCNPJ.Caption = "CNPJ:"
      Me.lblIE.Caption = "IE:"
      Me.vlrCPFCNPJ.Mask = "##.###.###/####-##"
      Me.vlrIE.Mask = "###.###.###.###"
   End If
End Sub

Private Sub cmdAdicionar_Click()
   On Error GoTo cmdAdicionar_Click_E
   
   Dim itemX As ListItem
   Dim frmModal As frmFornecedorContato
   Dim clsContainerContatos As sContainerFornContato
      
   Set frmModal = New frmFornecedorContato
   With frmModal
      .Show vbModal
      
      If Not .Cancelado Then
         lngSeqLst = lngSeqLst + 1
         
         Set clsContainerContatos = New sContainerFornContato
         clsContainerContatos.Sequencia = lngSeqLst
         clsContainerContatos.Contato = .Contato
         clsContainerContatos.Telefone = .Telefone
         clsContainerContatos.Email = .Email
         clsContainerContatos.Acao = "I"
         colContatos.Add clsContainerContatos, clsContainerContatos.Key
         
         Set itemX = Me.lstContatos.ListItems.Add(, clsContainerContatos.Key, clsContainerContatos.Contato)
         itemX.SubItems(CONST_LST_TELEFONE) = clsContainerContatos.Telefone
         itemX.SubItems(CONST_LST_EMAIL) = clsContainerContatos.Email
         itemX.SmallIcon = CONST_ICO_INSERIDO
         itemX.Selected = True
      End If
   End With
   
   GoTo DestruirObjetos
   
cmdAdicionar_Click_E:
   clsErro.Salvar Err
   Exibir clsErro, "cmdAdicionar_Click"
   
DestruirObjetos:
   If Not (frmModal Is Nothing) Then Unload frmModal
   Set frmModal = Nothing
End Sub

Private Sub cmdConsultar_Click()
   On Error GoTo cmdConsultar_Click_E
   
   Dim itemX As ListItem
   Dim clsCursor As INF_Cursor.Cursor
   Dim clsContainerContatos As sContainerFornContato
   
   If Not mCmpObrigatorio(clsErro, Me.vlrCod, "Código") Then GoTo Erro_Msg
   
   Set clsCursor = CreateObject("INF_Cursor.Cursor")
   With clsCursor
      .Inicializar clsConexao
      
      .SQL.Limpar
      .SQL.Mais " SELECT Empresa, Codigo, RazaoSocial, DataCad, CPFCNPJ, "
      .SQL.Mais "    RGIE, CEP, TpPessoa, Situacao, Endereco, Numero, CodCidade, "
      .SQL.Mais "    Bairro, Estado, Telefone, Fax, Cel, Email, Site "
      .SQL.Mais " FROM Fornecedores "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND Codigo = " & .Vlr(Me.vlrCod)
      
      If Not .Abrir(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         Exibir clsErro, "cmdConsultar_Click"
         GoTo DestruirObjetos
      End If
   
      If Not .EOF Then
         Me.txtNome = .Valor("RazaoSocial")
         Me.datCadastro = .Valor("DataCad")
         
         Me.cmbTipoPessoa.ListIndex = f1.CmbValor(Me.cmbTipoPessoa, .Valor("TpPessoa"))
         Select Case f1.CmbIndex(Me.cmbTipoPessoa)
         Case 0
            If Len(Trim(.Valor("CPFCNPJ"))) > 0 Then Me.vlrCPFCNPJ = Format(Mid(CStr(.Valor("CPFCNPJ")), 1, 9), "000,###,###") & "-" & Format(Mid(.Valor("CPFCNPJ"), 10, 2), "##")

            Me.vlrIE = .Valor("RGIE")
         Case 1
         End Select
         Me.vlrCEP = .Valor("CEP")
         Me.cmbSituacao.ListIndex = f1.CmbValor(Me.cmbSituacao, .Valor("Situacao"))
         Me.txtEndereco = .Valor("Endereco")
         Me.txtBairro = .Valor("Bairro")
         Me.vlrNro = .Valor("Numero")
         Me.vlrCodCidade = .Valor("CodCidade")
         Me.cmbUf.ListIndex = f1.CmbValor(Me.cmbUf, .Valor("Estado"))
         Me.txtTelefone = .Valor("Telefone")
         Me.txtFax = .Valor("Fax")
         Me.txtCelular = .Valor("Cel")
         Me.txtEmail = .Valor("Email")
         Me.txtSite = .Valor("Site")
         
         Me.cmdNovo.Enabled = False
         Me.cmdConsultar.Enabled = False
         
         Me.cmdSalvar.Caption = "&Alterar"
         If Not HabilitarBotao(clsErro, Me, Me.cmdSalvar) Then Exibir clsErro, "cmdConsultar_Click"
         If Not HabilitarBotao(clsErro, Me, Me.cmdExcluir) Then Exibir clsErro, "cmdConsultar_Click"
         
         Me.fraIdentificacao.Enabled = False
         Me.fraParametros.Enabled = True
         Me.fraContatos.Enabled = True
      Else
         mMsgInfo "Registro não localizado! Verifique."
         mFocus Me.vlrCod
         GoTo DestruirObjetos
      End If
      .Fechar
   End With
   
   With clsCursor
      .SQL.Limpar
      .SQL.Mais " SELECT Empresa, Codigo, CodFornecedor, Nome, Telefone, Email "
      .SQL.Mais " FROM FornecedoresContatos "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND CodFornecedor = " & .Vlr(Me.vlrCod)
      
      If Not .Abrir(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         Exibir clsErro, "cmdConsultar_Click"
         GoTo DestruirObjetos
      End If
      
      Do Until .EOF
         lngSeqLst = lngSeqLst + 1
         
         Set clsContainerContatos = New sContainerFornContato
         clsContainerContatos.Sequencia = lngSeqLst
         clsContainerContatos.Codigo = .Valor("Codigo")
         clsContainerContatos.Contato = .Valor("Nome")
         clsContainerContatos.Telefone = .Valor("Telefone")
         clsContainerContatos.Email = .Valor("Email")
         clsContainerContatos.Acao = "G"
         colContatos.Add clsContainerContatos, clsContainerContatos.Key
         
         Set itemX = Me.lstContatos.ListItems.Add(, clsContainerContatos.Key, clsContainerContatos.Contato)
         itemX.SubItems(CONST_LST_TELEFONE) = clsContainerContatos.Telefone
         itemX.SubItems(CONST_LST_EMAIL) = clsContainerContatos.Email
         itemX.SmallIcon = CONST_ICO_GRAVADO
         itemX.Selected = True
      
         .ProximoRegistro
      Loop
      .Fechar
   End With
      
   GoTo DestruirObjetos

cmdConsultar_Click_E:
   clsErro.Salvar Err
   Exibir clsErro, "cmdConsultar_Click"
   GoTo DestruirObjetos
   
Erro_Msg:
   Exibir clsErro, "cmdConsultar_Click"

DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Sub

Private Sub cmdConsultarCEP_Click()
   On Error GoTo cmdConsultarCEP_Click_E
   
   Dim url As String
   Dim clsCursor As INF_Cursor.Cursor
   Dim objXml As MSXML2.DOMDocument
   Dim objXmlNodes As MSXML2.IXMLDOMElement
   Dim ObjXmlElement As MSXML2.IXMLDOMElement
   
   If Not f1.VerificaConexaoInternet(clsErro) Then
      Exibir clsErro, "cmdConsultarCEP_Click"
      GoTo DestruirObjetos
   End If
   
   Set objXml = New MSXML2.DOMDocument

   url = "http://republicavirtual.com.br/web_cep.php?cep=" & Replace(Me.vlrCEP, "-", "") & "&formato=XML"
   objXml.async = False
   
   Call objXml.Load(url)

   If objXml.parseError.reason <> "" Then
      MsgBox objXml.parseError.reason
      Exit Sub
   End If

   Set objXmlNodes = objXml.documentElement

   For Each ObjXmlElement In objXmlNodes.childNodes
      Select Case ObjXmlElement.nodeName
      
      Case "resultado"
         If ObjXmlElement.nodeTypedValue = 0 Then
            mMsgInfo "CEP inválido! Verifique."
            mFocus Me.vlrCEP
            Exit Sub
         End If
      Case "cidade"
         
         Set clsCursor = CreateObject("INF_Cursor.Cursor")
         With clsCursor
            .Inicializar clsConexao
            
            .SQL.Limpar
            .SQL.Mais " SELECT Codigo, Municipio, UF "
            .SQL.Mais " FROM Municipios "
            .SQL.Mais " WHERE UPPER(Municipio) = " & .Txt(UCase(ObjXmlElement.nodeTypedValue))
            
            If Not .Abrir(.SQL.Texto) Then
               clsErro.Transferir = .TransferirErro
               Exibir clsErro, "cmdConsultarCEP_Click"
               GoTo DestruirObjetos
            End If
            
            If Not .EOF Then
               Me.vlrCodCidade = .Valor("Codigo")
               Me.txtDescCidade = .Valor("Municipio")
               Me.cmbUf.ListIndex = f1.CmbValor(Me.cmbUf, .Valor("UF"), enDescricao, 2)
            End If
            .Fechar
         End With
         
      Case "bairro"
         Me.txtBairro = ObjXmlElement.nodeTypedValue
         
      Case "tipo_logradouro"
         Me.txtEndereco = ObjXmlElement.nodeTypedValue
         
      Case "logradouro"
         Me.txtEndereco = Me.txtEndereco & " " & ObjXmlElement.nodeTypedValue
      End Select
   Next
   
   mFocus Me.vlrNro

   GoTo DestruirObjetos
   
cmdConsultarCEP_Click_E:
   clsErro.Salvar Err
   Exibir clsErro, "cmdConsultarCEP_Click"

DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
   Set objXml = Nothing
   Set objXmlNodes = Nothing
   Set ObjXmlElement = Nothing
End Sub

Private Sub cmdExcluir_Click()
   On Error GoTo cmdExcluir_Click_E
   
   Dim strBuffer As String
   Dim clsCursor As INF_Cursor.Cursor
   
   With clsCursor
      .Inicializar clsConexao
      
      .SQL.Limpar
      .SQL.Mais " SELECT CodFornecedor, CodProduto "
      .SQL.Mais " FROM ProdutosForn "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND CodFornecedor = " & .Vlr(Me.vlrCod)
      
      If Not .Abrir(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         Exibir clsErro, "cmdExcluir_Click"
         GoTo DestruirObjetos
      End If
      
      Do Until .EOF
         If Len(Trim(strBuffer)) > 0 Then strBuffer = strBuffer & ", "
         strBuffer = strBuffer & .Valor("CodProduto")
               
         .ProximoRegistro
      Loop
      .Fechar
   End With
   
   If Len(Trim(strBuffer)) > 0 Then
      mMsgInfo "Fornecedor vinculado ao(s) produto(s) abaixo: " & vbNewLine & strBuffer
      mFocus Me.cmdExcluir
      GoTo DestruirObjetos
   End If
   
   If vbNo = mMsgPerg("Você deseja realmente excluir este fornecedor?") Then Exit Sub
   
   With clsConexao
      .Begin
      
      .SQL.Limpar
      .SQL.Mais " DELETE FROM FornecedoresContatos "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND CodFornecedor = " & .Vlr(Me.vlrCod)
      
      If Not .Executar(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         clsConexao.RollBack
         Exibir clsErro, "cmdExcluir_Click"
         GoTo DestruirObjetos
      End If
      .Commit
   End With
   
   mMsgInfo "Registro excluido com sucesso! Verifique."
   cmdLimpar_Click
   
   GoTo DestruirObjetos
   
cmdExcluir_Click_E:
   clsErro.Salvar Err
   clsConexao.RollBack
   Exibir clsErro, "cmdExcluir_Click"

DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub cmdLimpar_Click()
   f1.Limpar Me
   
   Me.fraIdentificacao.Enabled = True
   Me.fraParametros.Enabled = False
   Me.fraContatos.Enabled = False
   
   Me.cmdSalvar.Caption = "&Salvar"
   Me.cmdSalvar.Enabled = False
   Me.cmdExcluir.Enabled = False
   
   f1.CollectionLimpar colContatos
   Me.lstContatos.ListItems.Clear
   lngSeqLst = 0
   
   If Not HabilitarBotao(clsErro, Me, Me.cmdNovo) Then Exibir clsErro, "cmdLimpar_Click"
   If Not HabilitarBotao(clsErro, Me, Me.cmdConsultar) Then Exibir clsErro, "cmdLimpar_Click"
   
   mFocus Me.vlrCod
End Sub

Private Sub cmdModificar_Click()
   On Error GoTo cmdModificar_Click_E
   
   Dim itemX As ListItem
   Dim frmModal As frmFornecedorContato
   Dim clsContainerContatos As sContainerFornContato
   
   Set itemX = Me.lstContatos.SelectedItem
   If itemX Is Nothing Then Exit Sub
   If Not itemX.Selected Then Exit Sub
   
   Set clsContainerContatos = colContatos(itemX.Key)
   Set frmModal = New frmFornecedorContato
   With frmModal
      .Contato = clsContainerContatos.Contato
      .Telefone = clsContainerContatos.Telefone
      .Email = clsContainerContatos.Email
      
      .Show vbModal
      
      If Not .Cancelado Then
         clsContainerContatos.Contato = .Contato
         clsContainerContatos.Telefone = .Telefone
         clsContainerContatos.Email = .Email
                  
         itemX.Text = clsContainerContatos.Contato
         itemX.SubItems(CONST_LST_TELEFONE) = clsContainerContatos.Telefone
         itemX.SubItems(CONST_LST_EMAIL) = clsContainerContatos.Email
         
         If Not clsContainerContatos.Acao = "I" Then
            clsContainerContatos.Acao = "A"
            itemX.SmallIcon = CONST_ICO_ALTERADO
         End If
         
         itemX.Selected = True
      End If
   End With
   
   GoTo DestruirObjetos

cmdModificar_Click_E:
   clsErro.Salvar Err
   Exibir clsErro, "cmdModificar_Click"

DestruirObjetos:
   If Not (frmModal Is Nothing) Then Unload frmModal
   Set frmModal = Nothing
End Sub

Private Sub cmdNovo_Click()
   cmdLimpar_Click
   Me.fraIdentificacao.Enabled = False
   Me.fraParametros.Enabled = True
   Me.fraContatos.Enabled = True
   
   Me.cmdNovo.Enabled = False
   Me.cmdConsultar.Enabled = False
   
   Me.cmdSalvar.Caption = "&Inserir"
   If Not HabilitarBotao(clsErro, Me, Me.cmdSalvar) Then Exibir clsErro, "cmdNovo_Click"
   
   mFocus Me.txtNome
End Sub

Private Sub cmdPesqCidade_Click()
   Dim frmModal As frmCG
   
   Set frmModal = New frmCG
   With frmModal
      .Codigo = Me.vlrCod
      .TpDefinicao = enMunicipios
      .Show vbModal
      
      If Not .Cancelado Then Me.vlrCodCidade = .Codigo
   End With
   
DestruirObjetos:
   If Not (frmModal Is Nothing) Then Unload frmModal
   Set frmModal = Nothing
End Sub

Private Sub cmdPesqForn_Click()
   Dim frmModal As frmCG
   
   Set frmModal = New frmCG
   With frmModal
      .Codigo = Me.vlrCod
      .TpDefinicao = enFornecedores
      .Show vbModal
      
      If Not .Cancelado Then Me.vlrCod = .Codigo
   End With
   
DestruirObjetos:
   If Not (frmModal Is Nothing) Then Unload frmModal
   Set frmModal = Nothing
End Sub

Private Sub cmdRemover_Click()
   On Error GoTo cmdRemover_Click_E
   
   Dim strKey As String
   Dim itemX As ListItem
   Dim frmModal As frmFornecedorContato
   Dim clsContainerContatos As sContainerFornContato
   
   Set itemX = Me.lstContatos.SelectedItem
   If itemX Is Nothing Then Exit Sub
   If Not itemX.Selected Then Exit Sub
   
   
   Set clsContainerContatos = colContatos(itemX.Key)
   
   If clsContainerContatos.Acao = "I" Then
      colContatos.Remove itemX.Key
      Me.lstContatos.ListItems.Remove itemX.Key
   Else
      clsContainerContatos.Acao = "E"
      itemX.SmallIcon = LST_ICO_REMOVIDO
   End If
      
   Exit Sub
   
cmdRemover_Click_E:
   clsErro.Salvar Err
   Exibir clsErro, "cmdRemover_Click"
End Sub

Private Sub cmdSalvar_Click()
   On Error GoTo cmdSalvar_Click_E
   
   Dim lngSeq As Long
   Dim strMsg As String
   Dim clsCursor As INF_Cursor.Cursor
   Dim clsContainerContatos As sContainerFornContato
   
   'Validações
   
   
   clsConexao.Begin
   
   Select Case Me.cmdSalvar.Caption
   Case "&Inserir"
      Set clsCursor = CreateObject("INF_Cursor.Cursor")
      With clsCursor
         .Inicializar clsConexao
         
         .SQL.Limpar
         .SQL.Mais " SELECT (ISNULL(MAX(Codigo), 0) + 1) AS MaxSeq "
         .SQL.Mais " FROM Fornecedores "
         .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
         
         If Not .Abrir(.SQL.Texto) Then
            clsErro.Transferir = .TransferirErro
            Exibir clsErro, "cmdSalvar_Click"
            GoTo DestruirObjetos
         End If
         
         lngSeq = .Valor("MaxSeq")
         
         .Fechar
      End With
      
      With clsConexao
         .SQL.Limpar
         .SQL.Mais " INSERT INTO Fornecedores "
         .SQL.Mais " (  Empresa, Codigo, RazaoSocial, DataCad, CPFCNPJ, "
         .SQL.Mais "    RGIE, CEP, TpPessoa, Situacao, Endereco, Numero, CodCidade, "
         .SQL.Mais "    Bairro, Estado, Telefone, Fax, Cel, Email, Site "
         .SQL.Mais " ) VALUES ( "
         .SQL.Mais .Txt(Prj.Sistema.IdEmpresa, True)
         .SQL.Mais .Vlr(lngSeq, True)
         .SQL.Mais .Txt(Me.txtNome, True)
         .SQL.Mais "GETDATE(),"
         .SQL.Mais .Txt(Me.vlrCPFCNPJ, True)
         .SQL.Mais .Txt(f1.SoNumeros(Me.vlrIE), True)
         .SQL.Mais .Txt(f1.SoNumeros(Me.vlrCEP), True)
         .SQL.Mais .Txt(f1.CmbParametro(Me.cmbTipoPessoa), True)
         .SQL.Mais .Txt(f1.CmbParametro(Me.cmbSituacao), True)
         .SQL.Mais .Txt(Me.txtEndereco, True)
         .SQL.Mais .Vlr(Me.vlrNro, True)
         .SQL.Mais .Vlr(Me.vlrCodCidade, True)
         .SQL.Mais .Txt(Me.txtBairro, True)
         .SQL.Mais .Txt(f1.CmbParametro(Me.cmbUf), True)
         .SQL.Mais .Txt(f1.SoNumeros(Me.txtTelefone), True)
         .SQL.Mais .Txt(f1.SoNumeros(Me.txtFax), True)
         .SQL.Mais .Txt(f1.SoNumeros(Me.txtCelular), True)
         .SQL.Mais .Txt(Me.txtEmail, True)
         .SQL.Mais .Txt(Me.txtSite)
         .SQL.Mais " ) "
         
         If Not .Executar(.SQL.Texto) Then
            clsErro.Transferir = .TransferirErro
            clsConexao.RollBack
            Exibir clsErro, "cmdSalvar_Click"
            GoTo DestruirObjetos
         End If
         
         strMsg = "Registro inserido com sucesso! Código: " & lngSeq
      End With
   
   Case "&Alterar"
      lngSeq = Me.vlrCod
      
      With clsConexao
         .SQL.Limpar
         .SQL.Mais " UPDATE Fornecedores SET "
         .SQL.Mais "    RazaoSocial = " & .Txt(Me.txtNome, True)
         .SQL.Mais "    CPFCNPJ = " & .Txt(f1.SoNumeros(Me.vlrCPFCNPJ), True)
         .SQL.Mais "    RGIE = " & .Txt(f1.SoNumeros(Me.vlrIE), True)
         .SQL.Mais "    CEP = " & .Txt(f1.SoNumeros(Me.vlrCEP), True)
         .SQL.Mais "    TpPessoa = " & .Txt(f1.CmbParametro(Me.cmbTipoPessoa), True)
         .SQL.Mais "    Situacao = " & .Txt(f1.CmbParametro(Me.cmbSituacao), True)
         .SQL.Mais "    Endereco = " & .Txt(Me.txtEndereco, True)
         .SQL.Mais "    Numero = " & .Vlr(Me.vlrNro, True)
         .SQL.Mais "    CodCidade = " & .Vlr(Me.vlrCodCidade, True)
         .SQL.Mais "    Bairro = " & .Txt(Me.txtBairro, True)
         .SQL.Mais "    Estado = " & .Txt(f1.CmbParametro(Me.cmbUf), True)
         .SQL.Mais "    Telefone = " & .Txt(f1.SoNumeros(Me.txtTelefone), True)
         .SQL.Mais "    Fax = " & .Txt(f1.SoNumeros(Me.txtFax), True)
         .SQL.Mais "    Cel = " & .Txt(f1.SoNumeros(Me.txtCelular), True)
         .SQL.Mais "    Email = " & .Txt(Me.txtEmail, True)
         .SQL.Mais "    Site = " & .Txt(Me.txtSite)
         .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
         .SQL.Mais " AND Codigo = " & .Vlr(lngSeq)
         
         If Not .Executar(.SQL.Texto) Then
            clsErro.Transferir = .TransferirErro
            clsConexao.RollBack
            Exibir clsErro, "cmdSalvar_Click"
            GoTo DestruirObjetos
         End If
         
         strMsg = "Registro alterado com sucesso!"
      End With
   End Select
   
   For Each clsContainerContatos In colContatos
      With clsConexao
         .SQL.Limpar
         
         Select Case clsContainerContatos.Acao
         Case "I"
            .SQL.Mais " INSERT INTO FornecedoresContatos "
            .SQL.Mais " ( Empresa, CodFornecedor, Nome, Telefone, Email "
            .SQL.Mais " ) VALUES ( "
            .SQL.Mais .Txt(Prj.Sistema.IdEmpresa, True)
            .SQL.Mais .Vlr(lngSeq, True)
            .SQL.Mais .Txt(clsContainerContatos.Contato, True)
            .SQL.Mais .Txt(f1.SoNumeros(clsContainerContatos.Telefone), True)
            .SQL.Mais .Txt(clsContainerContatos.Email)
            .SQL.Mais " ) "
         Case "A"
            .SQL.Mais " UPDATE FornecedoresContatos SET "
            .SQL.Mais "    Nome = " & .Txt(clsContainerContatos.Contato, True)
            .SQL.Mais "    Telefone = " & .Txt(f1.SoNumeros(clsContainerContatos.Telefone), True)
            .SQL.Mais "    Email = " & .Txt(clsContainerContatos.Email)
            .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
            .SQL.Mais " AND Codigo = " & .Vlr(clsContainerContatos.Codigo)
         Case "E"
            .SQL.Mais " DELETE FROM FornecedoresContatos "
            .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
            .SQL.Mais " AND Codigo = " & .Vlr(clsContainerContatos.Codigo)
         Case "G"
            'Sem ação
         End Select
         
         If .SQL.Tamanho > 0 Then
            If Not .Executar(.SQL.Texto) Then
               clsErro.Transferir = .TransferirErro
               clsConexao.RollBack
               Exibir clsErro, "cmdSalvar_Click"
               GoTo DestruirObjetos
            End If
         End If
      End With
   Next
   
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
   Set colContatos = New Collection
   Set clsErro = CreateObject("INF_Erro.Funcoes")

   f1.FormCentralizar Me

   If Not HabilitarBotao(clsErro, Me, Me.cmdNovo) Then Exibir clsErro, "Form_Load"
   If Not HabilitarBotao(clsErro, Me, Me.cmdConsultar) Then Exibir clsErro, "Form_Load"

   'Carrega combo com as UF - Unidades Federativas
   f1.CmbAdd Me.cmbUf, "AL", 0
   f1.CmbAdd Me.cmbUf, "AP", 1
   f1.CmbAdd Me.cmbUf, "AM", 2
   f1.CmbAdd Me.cmbUf, "BA", 3
   f1.CmbAdd Me.cmbUf, "CE", 4
   f1.CmbAdd Me.cmbUf, "DF", 5
   f1.CmbAdd Me.cmbUf, "ES", 6
   f1.CmbAdd Me.cmbUf, "GO", 7
   f1.CmbAdd Me.cmbUf, "MA", 8
   f1.CmbAdd Me.cmbUf, "MT", 9
   f1.CmbAdd Me.cmbUf, "MS", 10
   f1.CmbAdd Me.cmbUf, "MG", 11
   f1.CmbAdd Me.cmbUf, "PR", 12
   f1.CmbAdd Me.cmbUf, "PB", 13
   f1.CmbAdd Me.cmbUf, "PA", 14
   f1.CmbAdd Me.cmbUf, "PE", 15
   f1.CmbAdd Me.cmbUf, "PI", 16
   f1.CmbAdd Me.cmbUf, "RJ", 17
   f1.CmbAdd Me.cmbUf, "RN", 18
   f1.CmbAdd Me.cmbUf, "RS", 19
   f1.CmbAdd Me.cmbUf, "RO", 20
   f1.CmbAdd Me.cmbUf, "RR", 21
   f1.CmbAdd Me.cmbUf, "SC", 22
   f1.CmbAdd Me.cmbUf, "SE", 23
   f1.CmbAdd Me.cmbUf, "SP", 24
   f1.CmbAdd Me.cmbUf, "TO", 25
   Me.cmbUf.ListIndex = f1.CmbValor(Me.cmbUf, 19)
   
   'Carrega Combo Tipo de Pessoa
   '0 = PF
   '1 = PJ
   f1.CmbAdd Me.cmbTipoPessoa, "Pessoa Física", 0
   f1.CmbAdd Me.cmbTipoPessoa, "Pessoa Juridica", 1
   Me.cmbTipoPessoa.ListIndex = f1.CmbValor(Me.cmbTipoPessoa, 0)
   
   'Carrega Combo Situacao
   '0 = Ativo
   '1 = Inativo
   f1.CmbAdd Me.cmbSituacao, "Ativado", 1
   f1.CmbAdd Me.cmbSituacao, "Desativado", 2
   Me.cmbSituacao.ListIndex = f1.CmbValor(Me.cmbSituacao, 2)
   
   Me.lstLegenda.ListItems.Clear
   Me.lstLegenda.ListItems.Add , "K_1", CONST_ICO_GRAVADO, , CONST_ICO_GRAVADO
   Me.lstLegenda.ListItems.Add , "K_2", CONST_ICO_INSERIDO, , CONST_ICO_INSERIDO
   Me.lstLegenda.ListItems.Add , "K_3", CONST_ICO_ALTERADO, , CONST_ICO_ALTERADO
   Me.lstLegenda.ListItems.Add , "K_4", CONST_ICO_REMOVIDO, , CONST_ICO_REMOVIDO
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   f1.CollectionLimpar colContatos
   Set colContatos = Nothing
   Set clsErro = Nothing
End Sub
