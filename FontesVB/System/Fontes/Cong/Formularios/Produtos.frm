VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "InfinityControl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmProdutos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Produtos"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14490
   Icon            =   "Produtos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   14490
   Tag             =   "020200"
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   13140
      Top             =   2475
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   13110
      Top             =   3015
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
            Picture         =   "Produtos.frx":000C
            Key             =   "Inserido"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Produtos.frx":0166
            Key             =   "Gravado"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Produtos.frx":02C0
            Key             =   "Alterado"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Produtos.frx":041A
            Key             =   "Removido"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstLegenda 
      Height          =   1065
      Left            =   13110
      TabIndex        =   74
      Top             =   6060
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
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   345
      Left            =   13110
      TabIndex        =   73
      Top             =   1965
      Width           =   1320
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "&Limpar"
      Height          =   345
      Left            =   13110
      TabIndex        =   72
      Top             =   1590
      Width           =   1320
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Enabled         =   0   'False
      Height          =   345
      Left            =   13110
      TabIndex        =   71
      Tag             =   "6"
      Top             =   1215
      Width           =   1320
   End
   Begin VB.CommandButton cmdSalvar 
      Cancel          =   -1  'True
      Caption         =   "&Salvar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   13110
      TabIndex        =   70
      Top             =   840
      Width           =   1320
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "&Consultar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   13110
      TabIndex        =   69
      Top             =   465
      Width           =   1320
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Novo"
      Enabled         =   0   'False
      Height          =   345
      Left            =   13110
      TabIndex        =   68
      Top             =   90
      Width           =   1320
   End
   Begin VB.Frame fraIdentificacao 
      Caption         =   "Identifica��o"
      Height          =   660
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   12990
      Begin VB.CommandButton Command12 
         Caption         =   "..."
         Height          =   255
         Left            =   2115
         TabIndex        =   2
         Top             =   315
         Width           =   270
      End
      Begin rdActiveText.ActiveText vlrCod 
         Height          =   315
         Left            =   915
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
      Begin VB.Label lblCodigo 
         Caption         =   "C�digo:"
         Height          =   285
         Left            =   120
         TabIndex        =   107
         Top             =   285
         Width           =   1500
      End
   End
   Begin VB.Frame fraCadastro 
      Caption         =   "Informa��es de Cadastro"
      Enabled         =   0   'False
      Height          =   1845
      Left            =   8985
      TabIndex        =   35
      Top             =   660
      Width           =   4065
      Begin rdActiveText.ActiveText txtUsuarioCad 
         Height          =   315
         Left            =   1365
         TabIndex        =   37
         Top             =   615
         Width           =   2565
         _ExtentX        =   4524
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
      Begin rdActiveText.ActiveText datCad 
         Height          =   315
         Left            =   1365
         TabIndex        =   36
         Top             =   240
         Width           =   1005
         _ExtentX        =   1773
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
      Begin rdActiveText.ActiveText datUltCad 
         Height          =   315
         Left            =   1365
         TabIndex        =   38
         Top             =   990
         Width           =   1005
         _ExtentX        =   1773
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
      Begin rdActiveText.ActiveText txtUsuarioUltCad 
         Height          =   315
         Left            =   1365
         TabIndex        =   39
         Top             =   1365
         Width           =   2565
         _ExtentX        =   4524
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
      Begin VB.Label lblUserUltAlt 
         Caption         =   "Usu�rio Ult. Alt.:"
         Height          =   210
         Left            =   120
         TabIndex        =   108
         Top             =   1395
         Width           =   1230
      End
      Begin VB.Label lblDtUltCad 
         Caption         =   "Dt. Ult. Alt.:"
         Height          =   270
         Left            =   120
         TabIndex        =   106
         Top             =   1035
         Width           =   1305
      End
      Begin VB.Label lblUserCad 
         Caption         =   "Usu�rio Cad.:"
         Height          =   210
         Left            =   120
         TabIndex        =   105
         Top             =   645
         Width           =   1230
      End
      Begin VB.Label lblDtCad 
         Caption         =   "Data Cad.:"
         Height          =   270
         Left            =   120
         TabIndex        =   104
         Top             =   270
         Width           =   1305
      End
   End
   Begin VB.Frame fraImagem 
      Caption         =   "Imagem"
      Enabled         =   0   'False
      Height          =   1845
      Left            =   6450
      TabIndex        =   31
      Top             =   660
      Width           =   2475
      Begin VB.CommandButton cmdZoom 
         Height          =   315
         Left            =   2025
         Picture         =   "Produtos.frx":0574
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Zoom Imagem."
         Top             =   900
         Width           =   345
      End
      Begin VB.CommandButton cmdRemoverImg 
         Caption         =   ">>"
         Height          =   315
         Left            =   2025
         TabIndex        =   33
         ToolTipText     =   "Remover Imagem."
         Top             =   570
         Width           =   345
      End
      Begin VB.CommandButton cmdAdicionarImg 
         Caption         =   "<<"
         Height          =   315
         Left            =   2025
         TabIndex        =   32
         ToolTipText     =   "Adicionar Imagem."
         Top             =   240
         Width           =   345
      End
      Begin VB.Image Img 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1470
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1845
      End
   End
   Begin VB.Frame fraValores 
      Caption         =   "Valores"
      Enabled         =   0   'False
      Height          =   2805
      Left            =   9780
      TabIndex        =   57
      Top             =   4320
      Width           =   3270
      Begin VB.ComboBox cmbCustos 
         Height          =   315
         Left            =   2535
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   1710
         Width           =   630
      End
      Begin VB.ComboBox cmbMargem 
         Height          =   315
         Left            =   2535
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   1335
         Width           =   630
      End
      Begin VB.ComboBox cmbComissao 
         Height          =   315
         Left            =   2535
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   960
         Width           =   630
      End
      Begin VB.ComboBox cmbFrete 
         Height          =   315
         Left            =   2535
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   585
         Width           =   630
      End
      Begin rdActiveText.ActiveText vlrTributacao 
         Height          =   315
         Left            =   1440
         TabIndex        =   58
         Top             =   210
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
      Begin rdActiveText.ActiveText vlrFrete 
         Height          =   315
         Left            =   1440
         TabIndex        =   59
         Top             =   585
         Width           =   1035
         _ExtentX        =   1826
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
      Begin rdActiveText.ActiveText vlrComissao 
         Height          =   315
         Left            =   1440
         TabIndex        =   61
         Top             =   960
         Width           =   1035
         _ExtentX        =   1826
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
      Begin rdActiveText.ActiveText vlrMargem 
         Height          =   315
         Left            =   1440
         TabIndex        =   63
         Top             =   1335
         Width           =   1035
         _ExtentX        =   1826
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
      Begin rdActiveText.ActiveText vlrCusto 
         Height          =   315
         Left            =   1440
         TabIndex        =   65
         Top             =   1710
         Width           =   1035
         _ExtentX        =   1826
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
      Begin rdActiveText.ActiveText vlrVenda 
         Height          =   315
         Left            =   1440
         TabIndex        =   67
         Top             =   2085
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   556
         Alignment       =   1
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
         MaxLength       =   13
         Text            =   "0,00"
         TextMask        =   4
         RawText         =   4
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
         Locked          =   -1  'True
      End
      Begin VB.Label lblVenda 
         Caption         =   "Valor Venda:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   103
         Top             =   2130
         Width           =   1410
      End
      Begin VB.Label lblCustos 
         Caption         =   "Custos:"
         Height          =   300
         Left            =   120
         TabIndex        =   102
         Top             =   1755
         Width           =   1275
      End
      Begin VB.Label lblMargem 
         Caption         =   "Margem:"
         Height          =   300
         Left            =   120
         TabIndex        =   101
         Top             =   1365
         Width           =   1275
      End
      Begin VB.Label lblComissao 
         Caption         =   "Comiss�o:"
         Height          =   300
         Left            =   105
         TabIndex        =   100
         Top             =   990
         Width           =   1275
      End
      Begin VB.Label lblFrete 
         Caption         =   "Frete:"
         Height          =   300
         Left            =   120
         TabIndex        =   99
         Top             =   615
         Width           =   1275
      End
      Begin VB.Label lblTributos 
         Caption         =   "Tributos:"
         Height          =   300
         Left            =   120
         TabIndex        =   98
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.Frame fraFornecedor 
      Caption         =   "Fornecedor"
      Enabled         =   0   'False
      Height          =   1905
      Left            =   60
      TabIndex        =   27
      Top             =   5220
      Width           =   6330
      Begin VB.CommandButton cmdAdicionar 
         Caption         =   "<<"
         Height          =   315
         Left            =   5880
         TabIndex        =   29
         ToolTipText     =   "Adicionar Fornecedor."
         Top             =   255
         Width           =   345
      End
      Begin VB.CommandButton cmdRemover 
         Caption         =   ">>"
         Height          =   315
         Left            =   5880
         TabIndex        =   30
         ToolTipText     =   "Remover Fornecedor."
         Top             =   600
         Width           =   345
      End
      Begin MSComctlLib.ListView lstFornecedor 
         Height          =   1560
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   2752
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "C�digo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Raz�o Social"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Telefone"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Valor Compra"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Lucro"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Valor Lucro"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame fraTributacao 
      Caption         =   "Tributa��o"
      Enabled         =   0   'False
      Height          =   2805
      Left            =   6450
      TabIndex        =   49
      Top             =   4320
      Width           =   3285
      Begin VB.ComboBox cmbIPI 
         Height          =   315
         Left            =   2550
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   1365
         Width           =   645
      End
      Begin VB.ComboBox cmbPISCOFINS 
         Height          =   315
         Left            =   2550
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   990
         Width           =   645
      End
      Begin VB.ComboBox cmbICMS 
         Height          =   315
         Left            =   2550
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   615
         Width           =   645
      End
      Begin VB.ComboBox cmbTributacao 
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   240
         Width           =   1695
      End
      Begin rdActiveText.ActiveText vlrICMS 
         Height          =   315
         Left            =   1500
         TabIndex        =   51
         Top             =   615
         Width           =   1005
         _ExtentX        =   1773
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
      Begin rdActiveText.ActiveText vlrPISCOFINS 
         Height          =   315
         Left            =   1500
         TabIndex        =   53
         Top             =   990
         Width           =   1005
         _ExtentX        =   1773
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
      Begin rdActiveText.ActiveText vlrIPI 
         Height          =   315
         Left            =   1500
         TabIndex        =   55
         Top             =   1365
         Width           =   1005
         _ExtentX        =   1773
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
      Begin VB.Label lblIPI 
         Caption         =   "IPI de Venda:"
         Height          =   300
         Left            =   120
         TabIndex        =   89
         Top             =   1395
         Width           =   1335
      End
      Begin VB.Label lblPISCOFINS 
         Caption         =   "PIS / COFINS:"
         Height          =   300
         Left            =   120
         TabIndex        =   88
         Top             =   1020
         Width           =   1320
      End
      Begin VB.Label lblICMS 
         Caption         =   "ICMS:"
         Height          =   300
         Left            =   120
         TabIndex        =   87
         Top             =   645
         Width           =   1320
      End
      Begin VB.Label lblTributa��o 
         Caption         =   "Tributa��o:"
         Height          =   285
         Left            =   120
         TabIndex        =   85
         Top             =   270
         Width           =   1395
      End
   End
   Begin VB.Frame fraEstoque 
      Caption         =   "Estoque"
      Enabled         =   0   'False
      Height          =   1830
      Left            =   6450
      TabIndex        =   40
      Top             =   2490
      Width           =   6600
      Begin VB.ComboBox cmbVenderSemEst 
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   615
         Width           =   1035
      End
      Begin rdActiveText.ActiveText vlrEstMin 
         Height          =   315
         Left            =   4740
         TabIndex        =   45
         Top             =   240
         Width           =   1725
         _ExtentX        =   3043
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
         MaxLength       =   5
         Text            =   "0"
         TextMask        =   3
         RawText         =   3
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin VB.ComboBox cmbControlaEst 
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   240
         Width           =   1035
      End
      Begin rdActiveText.ActiveText vlrEstEntrada 
         Height          =   315
         Left            =   1500
         TabIndex        =   43
         Top             =   1005
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   556
         Alignment       =   1
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
         Text            =   "0"
         TextMask        =   3
         RawText         =   3
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText vlrEstSaida 
         Height          =   315
         Left            =   4740
         TabIndex        =   47
         Top             =   990
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         Alignment       =   1
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
         Text            =   "0"
         TextMask        =   3
         RawText         =   3
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
         Locked          =   -1  'True
      End
      Begin rdActiveText.ActiveText vlrEstReservado 
         Height          =   315
         Left            =   1500
         TabIndex        =   44
         Top             =   1380
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   556
         Alignment       =   1
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
         Text            =   "0"
         TextMask        =   3
         RawText         =   3
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText vlrCodDisponivel 
         Height          =   315
         Left            =   4740
         TabIndex        =   48
         Top             =   1365
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         Alignment       =   1
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
         Text            =   "0"
         TextMask        =   3
         RawText         =   3
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
         Locked          =   -1  'True
      End
      Begin rdActiveText.ActiveText vlrEstAtual 
         Height          =   315
         Left            =   4740
         TabIndex        =   46
         Top             =   615
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         Alignment       =   1
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
         Text            =   "0"
         TextMask        =   3
         RawText         =   3
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
         Locked          =   -1  'True
      End
      Begin VB.Label lblDisponivel 
         Caption         =   "Dispon�vel:"
         Height          =   300
         Left            =   3405
         TabIndex        =   86
         Top             =   1395
         Width           =   1335
      End
      Begin VB.Label lblReservado 
         Caption         =   "Reservado:"
         Height          =   300
         Left            =   120
         TabIndex        =   84
         Top             =   1410
         Width           =   945
      End
      Begin VB.Label lblSaida 
         Caption         =   "Sa�da:"
         Height          =   300
         Left            =   3405
         TabIndex        =   83
         Top             =   1020
         Width           =   1290
      End
      Begin VB.Label lblEntrada 
         Caption         =   "Entrada:"
         Height          =   300
         Left            =   120
         TabIndex        =   82
         Top             =   1035
         Width           =   945
      End
      Begin VB.Label lblEstoque 
         Caption         =   "Est. Atual:"
         Height          =   180
         Left            =   3390
         TabIndex        =   81
         Top             =   645
         Width           =   1380
      End
      Begin VB.Label lblVendaSemEst 
         Caption         =   "Vender Sem Est.:"
         Height          =   315
         Left            =   120
         TabIndex        =   80
         Top             =   645
         Width           =   1350
      End
      Begin VB.Label lblEstMin 
         Caption         =   "Est. M�nimo:"
         Height          =   300
         Left            =   3405
         TabIndex        =   79
         Top             =   270
         Width           =   1320
      End
      Begin VB.Label lblControlaEst 
         Caption         =   "Controla Est.:"
         Height          =   285
         Left            =   120
         TabIndex        =   78
         Top             =   270
         Width           =   1395
      End
   End
   Begin VB.Frame fraParametros 
      Caption         =   "Par�metros"
      Enabled         =   0   'False
      Height          =   4560
      Left            =   60
      TabIndex        =   3
      Top             =   660
      Width           =   6330
      Begin VB.CommandButton cmdPesqCodBarra 
         Height          =   315
         Left            =   3225
         Picture         =   "Produtos.frx":06BE
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   345
      End
      Begin VB.ComboBox cmbLucro 
         Height          =   315
         Left            =   5565
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   4125
         Width           =   645
      End
      Begin rdActiveText.ActiveText vlrLucroMin 
         Height          =   315
         Left            =   4455
         TabIndex        =   25
         Top             =   4125
         Width           =   1065
         _ExtentX        =   1879
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
      Begin VB.ComboBox cmbSituacao 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   4125
         Width           =   1080
      End
      Begin VB.TextBox txtObservacao 
         Height          =   825
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   3240
         Width           =   5010
      End
      Begin VB.CommandButton cmdConsUnidade 
         Caption         =   "..."
         Height          =   255
         Left            =   2340
         TabIndex        =   21
         Top             =   2925
         Width           =   270
      End
      Begin VB.CommandButton cmdConsModelo 
         Caption         =   "..."
         Height          =   255
         Left            =   2340
         TabIndex        =   18
         Top             =   2550
         Width           =   270
      End
      Begin VB.CommandButton cmdConsMarca 
         Caption         =   "..."
         Height          =   255
         Left            =   2340
         TabIndex        =   15
         Top             =   2175
         Width           =   270
      End
      Begin VB.CommandButton cmdConsSubGrupo 
         Caption         =   "..."
         Height          =   255
         Left            =   2340
         TabIndex        =   12
         Top             =   1800
         Width           =   270
      End
      Begin rdActiveText.ActiveText txtDescGrupo 
         Height          =   315
         Left            =   2685
         TabIndex        =   10
         Top             =   1365
         Width           =   3525
         _ExtentX        =   6218
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
      Begin VB.CommandButton cmdConsGrupo 
         Caption         =   "..."
         Height          =   255
         Left            =   2340
         TabIndex        =   9
         Top             =   1425
         Width           =   270
      End
      Begin rdActiveText.ActiveText vlrCodGrupo 
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         Top             =   1365
         Width           =   1080
         _ExtentX        =   1905
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
      Begin rdActiveText.ActiveText txtDescricao 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   615
         Width           =   5010
         _ExtentX        =   8837
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
         MaxLength       =   80
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtCodBarras 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   1965
         _ExtentX        =   3466
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
         Text            =   "0"
         TextMask        =   3
         RawText         =   3
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtAbreviatura 
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Top             =   990
         Width           =   5010
         _ExtentX        =   8837
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
      Begin rdActiveText.ActiveText vlrCodSubGrupo 
         Height          =   315
         Left            =   1200
         TabIndex        =   11
         Top             =   1740
         Width           =   1080
         _ExtentX        =   1905
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
      Begin rdActiveText.ActiveText txtDescSubGrupo 
         Height          =   315
         Left            =   2685
         TabIndex        =   13
         Top             =   1740
         Width           =   3525
         _ExtentX        =   6218
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
      Begin rdActiveText.ActiveText vlrCodMarca 
         Height          =   315
         Left            =   1200
         TabIndex        =   14
         Top             =   2115
         Width           =   1080
         _ExtentX        =   1905
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
      Begin rdActiveText.ActiveText txtDescMarca 
         Height          =   315
         Left            =   2685
         TabIndex        =   16
         Top             =   2115
         Width           =   3525
         _ExtentX        =   6218
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
      Begin rdActiveText.ActiveText vlrCodModelo 
         Height          =   315
         Left            =   1200
         TabIndex        =   17
         Top             =   2490
         Width           =   1080
         _ExtentX        =   1905
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
      Begin rdActiveText.ActiveText txtDescModelo 
         Height          =   315
         Left            =   2685
         TabIndex        =   19
         Top             =   2490
         Width           =   3525
         _ExtentX        =   6218
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
      Begin rdActiveText.ActiveText vlrCodUnidade 
         Height          =   315
         Left            =   1200
         TabIndex        =   20
         Top             =   2865
         Width           =   1080
         _ExtentX        =   1905
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
      Begin rdActiveText.ActiveText txtDescUnidade 
         Height          =   315
         Left            =   2685
         TabIndex        =   22
         Top             =   2865
         Width           =   3525
         _ExtentX        =   6218
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
      Begin VB.Label lblLucro 
         Caption         =   "Lucro M�nimo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3150
         TabIndex        =   97
         Top             =   4155
         Width           =   1305
      End
      Begin VB.Label lblSituacao 
         Caption         =   "Situa��o:"
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
         Left            =   105
         TabIndex        =   96
         Top             =   4155
         Width           =   1290
      End
      Begin VB.Label lblOs 
         Caption         =   "Observa��es:"
         Height          =   285
         Left            =   120
         TabIndex        =   95
         Top             =   3270
         Width           =   1020
      End
      Begin VB.Label lblUnidade 
         Caption         =   "Unidade:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   94
         Top             =   2895
         Width           =   1050
      End
      Begin VB.Label lblModelo 
         Caption         =   "Modelo:"
         Height          =   270
         Left            =   120
         TabIndex        =   93
         Top             =   2535
         Width           =   1050
      End
      Begin VB.Label lblMarca 
         Caption         =   "Marca:"
         Height          =   270
         Left            =   120
         TabIndex        =   92
         Top             =   2145
         Width           =   1050
      End
      Begin VB.Label lblSubGrupo 
         Caption         =   "Sub Gupo:"
         Height          =   270
         Left            =   120
         TabIndex        =   91
         Top             =   1770
         Width           =   1050
      End
      Begin VB.Label lblGrupo 
         Caption         =   "Gupo:"
         Height          =   270
         Left            =   120
         TabIndex        =   90
         Top             =   1395
         Width           =   1050
      End
      Begin VB.Label lblAbreviatura 
         Caption         =   "Abreviatura:"
         Height          =   315
         Left            =   120
         TabIndex        =   77
         Top             =   1020
         Width           =   1080
      End
      Begin VB.Label lblDescricao 
         Caption         =   "Descri��o:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   76
         Top             =   645
         Width           =   1005
      End
      Begin VB.Label lblCodBarras 
         Caption         =   "C�d. Barras:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   75
         Top             =   270
         Width           =   1065
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Legenda"
      Height          =   225
      Left            =   13110
      TabIndex        =   109
      Top             =   5835
      Width           =   1215
   End
End
Attribute VB_Name = "frmProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsErro As INF_Erro.Funcoes
Private colFornecedores As Collection

'Lista fornecedor colunas
Private Const LST_COL_CODFORN As Integer = 0
Private Const LST_COL_DSCFORN As Integer = 1
Private Const LST_COL_TELFORN As Integer = 2
Private Const LST_cOL_VALORCOMPRA As Integer = 3
Private Const LST_COL_LUCROFORN As Integer = 4
Private Const LST_COL_VALORLUCRO As Integer = 5

Private Sub cmbControlaEst_Change()
   If Left(Me.cmbControlaEst.Text, 1) = "S" Then
      Me.lblEstMin.FontBold = True
   Else
      Me.lblEstMin.FontBold = False
   End If
End Sub

Private Sub cmdAdicionar_Click()
   On Error GoTo cmdAdicionar_Click_E:
   
   Dim itemX As ListItem
   Dim curLucroPerc As Currency
   Dim curLucroVlr As Currency
   Dim frmModal As frmProdutosFornMod
   
   Dim ContainerProdForn As sContainerProdutosForn
   
   Set frmModal = New frmProdutosFornMod
   With frmModal
      .Show vbModal
      
      If Not .Cancelado Then
      
         For Each ContainerProdForn In colFornecedores
            If ContainerProdForn.CodFornecedor = .Codigo Then
               mMsgInfo "O fornecedor " & .Descricao & " j� consta na lista! Verifique."
               mFocus Me.lstFornecedor
               Exit Sub
            End If
         Next
         
         Set ContainerProdForn = New sContainerProdutosForn
         ContainerProdForn.CodFornecedor = .Codigo
         ContainerProdForn.DscFornecedor = .Descricao
         ContainerProdForn.Telefone = .Telefone
         ContainerProdForn.ValorCompra = .ValorCompra
         ContainerProdForn.Situacao = "I"
         ContainerProdForn.Lucro = 0
         ContainerProdForn.LucroValor = 0
         colFornecedores.Add ContainerProdForn, ContainerProdForn.Key
         
         'Verifica o maior fornecedor
         Call AtualizaLista
         
      End If
   End With
     
   Exit Sub

cmdAdicionar_Click_E:
   clsErro.Salvar Err
   Exibir clsErro, "cmdAdicionar_Click"
End Sub

Private Sub cmdAdicionarImg_Click()
   Me.CommonDialog.DialogTitle = "Infinity - Procura Arquivos .jpg" 'Titulo
   Me.CommonDialog.InitDir = "C:\" 'Caminho inicial
   Me.CommonDialog.Filter = "Imgs JPG(*.jpg)|*.jpg"
   Me.CommonDialog.FilterIndex = 1
   
   ' --------------------------------------------------------------------------- '
   ' cdlOFNFileMustExist: selecionar somente arquivos que existam                |
   ' cdlOFNHideReadOnly : exibe arquivos que possuem o atributo: Somente Leitura |
   ' cdlOFNLongNames    : exibe nome de arquivos longos                          |
   ' cdlOFNExplorer     : usa a interface estilo do Explorer                     |
   ' --------------------------------------------------------------------------- '

   Me.CommonDialog.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNExplorer
   'dipara um erro se nada for selecionado
   'Me.CommonDialog.CancelError = True

   Me.CommonDialog.ShowOpen

   If Len(Trim(Me.CommonDialog.FileName)) <> 0 Then Me.img.Picture = LoadPicture(Me.CommonDialog.FileName)
   
End Sub





Private Sub cmdConsultar_Click()
   On Error GoTo cmdConsultar_Click_E
   
   Dim itemX As ListItem
   Dim clsCursor As INF_Cursor.Cursor
   Dim ContainerProdForn As sContainerProdutosForn
   
   If Not mCmpObrigatorio(clsErro, Me.vlrCod, "C�digo") Then
      Exibir clsErro, "cmdConsultar_Click"
      Exit Sub
   End If
   
   'Verifica Cadastro do Produto
   Set clsCursor = CreateObject("INF_Cursor.Cursor")
   With clsCursor
      .Inicializar clsConexao
      
      .SQL.Limpar
      .SQL.Mais " SELECT Produtos.Empresa, Produtos.Codigo, Produtos.CodBarra, "
      .SQL.Mais "    Produtos.Descricao, Produtos.Abreviatura, Produtos.CodGrupo, "
      .SQL.Mais "    Produtos.CodSubGrupo, Produtos.CodMarca, Produtos.CodModelo, "
      .SQL.Mais "    Produtos.CodUnidade, Produtos.Obs, Produtos.IdSituacao, "
      .SQL.Mais "    Produtos.LucroMin, Produtos.IdLucroMin, Produtos.DtCad, "
      .SQL.Mais "    Produtos.CodUsuario, Produtos.DtUltAlt, Produtos.CodUsuarioUltAlt, "
      .SQL.Mais "    Produtos.IdControlaEst, Produtos.EstoqueMin, Produtos.IdVenderSemEst, "
      .SQL.Mais "    Produtos.IdTributacao , Produtos.ICMS, Produtos.IdICMS, "
      .SQL.Mais "    Produtos.PISCOFINS, Produtos.IdPISCOFINS, Produtos.IPI, "
      .SQL.Mais "    Produtos.IdIPI, Produtos.Tributos, Produtos.Frete, Produtos.IdFrete, "
      .SQL.Mais "    Produtos.Comissao, Produtos.IdComissao, Produtos.Margem, "
      .SQL.Mais "    Produtos.IdMargem, Produtos.Custos, Produtos.IdCustos, "
      .SQL.Mais "    Produtos.ValorVenda, "
      
      'Descri��es
      .SQL.Mais "    Marcas.Descricao AS DscMarca, "
      .SQL.Mais "    Modelos.Descricao AS DscModelo, "
      .SQL.Mais "    Unidades.Descricao AS DscUnidade, "
      .SQL.Mais "    Grupos.Descricao AS DscGrupo, "
      .SQL.Mais "    SubGrupos.Descricao AS DscSubGrupo, "
      .SQL.Mais "    Usuarios.Usuario AS NmUsuario "
      
      .SQL.Mais " FROM Produtos "
      .SQL.Mais " LEFT JOIN Grupos ON (" & rel_Grupos_Produtos & ")"
      .SQL.Mais " LEFT JOIN SubGrupos ON (" & rel_SubGrupos_Produtos & ")"
      .SQL.Mais " LEFT JOIN Marcas ON (" & rel_Marcas_Produtos & ")"
      .SQL.Mais " LEFT JOIN Modelos ON (" & rel_Modelos_Produtos & ")"
      .SQL.Mais " LEFT JOIN Unidades ON (" & rel_Unidades_Produtos & ")"
      .SQL.Mais " LEFT JOIN Usuarios ON (" & rel_Usuarios_Produtos & ")"
      
      .SQL.Mais " WHERE Produtos.Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND Produtos.Codigo = " & .Vlr(Me.vlrCod)
      
      If Not .Abrir(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         Exibir clsErro, "cmdConsultar_Click"
         GoTo DestruirObjetos
      End If
      
      If Not .EOF Then
         Me.txtCodBarras = .Valor("CodBarra")
         Me.txtDescricao = .Valor("Descricao")
         Me.txtAbreviatura = .Valor("Abreviatura")
         Me.vlrCodGrupo = .Valor("CodGrupo")
         Me.vlrCodSubGrupo = .Valor("CodSubGrupo")
         Me.vlrCodMarca = .Valor("CodMarca")
         Me.vlrCodModelo = .Valor("CodModelo")
         Me.vlrCodUnidade = .Valor("CodUnidade")
         Me.txtObservacao = .Valor("Obs")
         Me.cmbSituacao.ListIndex = f1.CmbValor(Me.cmbSituacao, .Valor("IdSituacao"))
         Me.vlrLucroMin = .Valor("LucroMin")
         Me.cmbLucro.ListIndex = f1.CmbValor(Me.cmbLucro, .Valor("IdLucroMin"))
         Me.datCad = .Valor("DtCad")
         Me.txtUsuarioCad = .Valor("NmUsuario")
'         Me.datUltCad = .Valor("DtUltAlt")
'         me.txtUsuarioUltCad =
         Me.cmbControlaEst.ListIndex = f1.CmbValor(Me.cmbControlaEst, .Valor("IdControlaEst"), enDescricao, 1)
         Me.vlrEstMin = .Valor("EstoqueMin")
         Me.cmbVenderSemEst.ListIndex = f1.CmbValor(Me.cmbVenderSemEst, .Valor("IdVenderSemEst"), enDescricao, 1)
         Me.cmbTributacao.ListIndex = f1.CmbValor(Me.cmbTributacao, .Valor("IdTributacao"))
         Me.vlrICMS = .Valor("ICMS")
         Me.cmbICMS.ListIndex = f1.CmbValor(Me.cmbICMS, .Valor("IdICMS"))
         Me.vlrPISCOFINS = .Valor("PISCOFINS")
         Me.cmbPISCOFINS.ListIndex = f1.CmbValor(Me.cmbPISCOFINS, .Valor("IdPISCOFINS"))
         Me.vlrIPI = .Valor("IPI")
         
         Me.cmdSalvar.Caption = "&Alterar"
         If Not HabilitarBotao(clsErro, Me, Me.cmdSalvar) Then Exibir clsErro, "cmdConsultar_Click"
         If Not HabilitarBotao(clsErro, Me, Me.cmdExcluir) Then Exibir clsErro, "cmdConsultar_Click"
         
         Me.fraIdentificacao.Enabled = False
         Me.fraCadastro.Enabled = True
         Me.fraEstoque.Enabled = True
         Me.fraImagem.Enabled = True
         Me.fraFornecedor.Enabled = True
         Me.fraParametros.Enabled = True
         Me.fraTributacao.Enabled = True
         Me.fraValores.Enabled = True
         
      Else
         mMsgInfo "Produto n�o localizado! Verifique."
         mFocus Me.vlrCod
         GoTo DestruirObjetos
      End If
      .Fechar
   End With
   
'   'Verifica o Estoque do Produto
'   With clsCursor
'      .SQL.Limpar
'      .SQL.Mais " SELECT EstEntrada, EstSaida, EstAtual "
'      .SQL.Mais " FROM VW_ProdutoEst "
'      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
'      .SQL.Mais " AND Codigo = " & .Vlr(Me.vlrCod)
'
'      If Not .Abrir(.SQL.Texto) Then
'         clsErro.Transferir = .TransferirErro
'         Exibir clsErro, "cmdConsultar_Click"
'         GoTo DestruirObjetos
'      End If
'
'      If Not .EOF Then
'         Me.vlrEstEntrada = .Valor("EstEntrada")
'         Me.vlrEstSaida = .Valor("EstSaida")
'         Me.vlrEstAtual = .Valor("EstAtual")
'      End If
'   End With
   
   'Verifica Fornecedores do Produto
   With clsCursor
      .SQL.Limpar
      .SQL.Mais " SELECT ProdutosForn.CodFornecedor, ProdutosForn.Lucro, "
      .SQL.Mais "    ProdutosForn.ValorCompra, ProdutosForn.ValorLucro, "
      .SQL.Mais "    Fornecedores.RazaoSocial, ISNULL(Fornecedores.Telefone, 0) AS Telefone "
      .SQL.Mais " FROM ProdutosForn "
      .SQL.Mais " INNER JOIN Fornecedores ON (" & rel_Fornecedores_ProdutosForn & ")"
      .SQL.Mais " WHERE ProdutosForn.Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND ProdutosForn.CodProduto = " & .Vlr(Me.vlrCod)
   
      If Not .Abrir(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         Exibir clsErro, "cmdConsultar_Click"
         GoTo DestruirObjetos
      End If
      
      Do Until .EOF
         'Armazena no container
         Set ContainerProdForn = New sContainerProdutosForn
         ContainerProdForn.CodFornecedor = .Valor("CodFornecedor")
         ContainerProdForn.DscFornecedor = .Valor("RazaoSocial")
         ContainerProdForn.Telefone = .Valor("Telefone")
         ContainerProdForn.Situacao = "G" 'Gravado
         colFornecedores.Add ContainerProdForn, ContainerProdForn.Key
      
         'Armazena na lista
         Set itemX = Me.lstFornecedor.ListItems.Add(, ContainerProdForn.Key, ContainerProdForn.CodFornecedor)
         itemX.SubItems(LST_COL_DSCFORN) = ContainerProdForn.DscFornecedor
         itemX.SubItems(LST_COL_TELFORN) = ContainerProdForn.Telefone
         itemX.SubItems(LST_COL_LUCROFORN) = ContainerProdForn.Lucro
         itemX.SubItems(LST_cOL_VALORCOMPRA) = ContainerProdForn.ValorCompra
         itemX.SubItems(LST_COL_VALORLUCRO) = ContainerProdForn.LucroValor
         itemX.SmallIcon = LST_ICO_GRAVADO
         itemX.Selected = True
         
         .ProximoRegistro
      Loop
      
      .Fechar
   End With
   
   'Verifica Imagem do Produto
   With clsCursor
      .SQL.Limpar
      .SQL.Mais " SELECT Empresa, CodProduto, PathImg "
      .SQL.Mais " FROM ProdutosPathImg "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND CodProduto = " & .Vlr(Me.vlrCod)
      
      If Not .Abrir(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         Exibir clsErro, "cmdConsultar_Click"
         GoTo DestruirObjetos
      End If
      
      If Not .EOF Then Me.img.Picture = LoadPicture(.Valor("PathImg"))

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
   
   Me.fraIdentificacao.Enabled = True
   Me.fraCadastro.Enabled = False
   Me.fraEstoque.Enabled = False
   Me.fraImagem.Enabled = False
   Me.fraFornecedor.Enabled = False
   Me.fraParametros.Enabled = False
   Me.fraTributacao.Enabled = False
   Me.fraValores.Enabled = False
   
   f1.CollectionLimpar colFornecedores
   Me.lstFornecedor.ListItems.Clear
   Me.img.Picture = LoadPicture(PATH_IMG & "SEM_IMG.JPG")
   DoEvents
   
   Me.cmdSalvar.Caption = "&Salvar"
   Me.cmdSalvar.Enabled = False
   Me.cmdExcluir.Enabled = False
   
   If Not HabilitarBotao(clsErro, Me, Me.cmdNovo) Then Exibir clsErro, "cmdNovo_Click"
   If Not HabilitarBotao(clsErro, Me, Me.cmdConsultar) Then Exibir clsErro, "cmdNovo_Click"
   
   mFocus Me.vlrCod
End Sub

Private Sub cmdNovo_Click()
   On Error GoTo cmdNovo_Click_E
   
   Me.fraIdentificacao.Enabled = False
   Me.fraCadastro.Enabled = True
   Me.fraEstoque.Enabled = True
   Me.fraImagem.Enabled = True
   Me.fraFornecedor.Enabled = True
   Me.fraParametros.Enabled = True
   Me.fraTributacao.Enabled = True
   Me.fraValores.Enabled = True
   
   Me.cmdNovo.Enabled = False
   Me.cmdConsultar.Enabled = False
   
   Me.cmdSalvar.Caption = "&Inserir"
   If Not HabilitarBotao(clsErro, Me, Me.cmdSalvar) Then Exibir clsErro, "cmdNovo_Click"
      
   Exit Sub

cmdNovo_Click_E:
   clsErro.Salvar Err
   Exibir clsErro, "cmdNovo_Click"
End Sub

Private Sub cmdRemoverImg_Click()
   If Len(Trim(Me.img.Picture)) > 0 Then
      If mMsgPerg("Voc� deseja realmente remover a imagem do produto?") Then
         Me.img.Picture = LoadPicture(PATH_IMG & "SEM_IMG.JPG")
         DoEvents
      End If
   End If
End Sub

Private Sub cmdSalvar_Click()
   On Error GoTo cmdSalvar_Click_E
   
   Dim strMsg As String
   Dim curSeqCod As Currency
   Dim clsCursor As INF_Cursor.Cursor
   Dim ContainerForn As sContainerProdutosForn
   
   Ampulheta True
   
   If Not mCmpObrigatorio(clsErro, Me.txtCodBarras, "C�digo de Barras") Then GoTo CmpObrig
   If Not mCmpObrigatorio(clsErro, Me.txtDescricao, "Descri��o") Then GoTo CmpObrig
   If Not mCmpObrigatorio(clsErro, Me.vlrCodUnidade, "C�digo da Unidade") Then GoTo CmpObrig
   If Not mCmpObrigatorio(clsErro, Me.vlrLucroMin, "Lucro M�nimo") Then GoTo CmpObrig
   
   If Left(Me.cmbControlaEst.Text, 1) = "S" Then
      If Not mCmpObrigatorio(clsErro, Me.vlrEstMin, "Estoque M�nimo") Then GoTo CmpObrig
   End If
         
   'Verifica Tipo de A��o
   Select Case Me.cmdSalvar.Caption
   Case "&Inserir"
      
      'Busca um novo c�digo para o produto
      Set clsCursor = CreateObject("INF_Cursor.Cursor")
      With clsCursor
         .Inicializar clsConexao
         
         .SQL.Limpar
         .SQL.Mais " SELECT (ISNULL(MAX(Codigo), 0) + 1) AS MaxSeq "
         .SQL.Mais " FROM Produtos "
         .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
         
         If Not .Abrir(.SQL.Texto) Then
            clsErro.Transferir = .TransferirErro
            Exibir clsErro, "cmdSalvar_Click"
            GoTo DestruirObjetos
         End If
         
         curSeqCod = .Valor("MaxSeq")
         
         .Fechar
      End With
      
      clsConexao.Begin
      
      'Produto
      With clsConexao
         .SQL.Limpar
         .SQL.Mais " INSERT INTO Produtos "
         .SQL.Mais " ( "
         .SQL.Mais "    Empresa, Codigo, CodBarra, Descricao, Abreviatura, CodGrupo, "
         .SQL.Mais "    CodSubGrupo, CodMarca, CodModelo, CodUnidade, Obs, IdSituacao, "
         .SQL.Mais "    LucroMin, IdLucroMin, DtCad, CodUsuario, "
         .SQL.Mais "    IdControlaEst, EstoqueMin, IdVenderSemEst, IdTributacao , ICMS, IdICMS, "
         .SQL.Mais "    PISCOFINS, IdPISCOFINS, IPI, IdIPI, Tributos, Frete, IdFrete, "
         .SQL.Mais "    Comissao, IdComissao, Margem, IdMargem, Custos, IdCustos, ValorVenda "
         .SQL.Mais " ) VALUES ( "
         .SQL.Mais .Txt(Prj.Sistema.IdEmpresa, True)
         .SQL.Mais .Vlr(curSeqCod, True)
         .SQL.Mais .Vlr(Me.txtCodBarras, True)
         .SQL.Mais .Txt(Me.txtDescricao, True)
         .SQL.Mais .Txt(Me.txtAbreviatura, True)
         .SQL.Mais .Vlr(Me.vlrCodGrupo, True)
         .SQL.Mais .Vlr(Me.vlrCodSubGrupo, True)
         .SQL.Mais .Vlr(Me.vlrCodMarca, True)
         .SQL.Mais .Vlr(Me.vlrCodModelo, True)
         .SQL.Mais .Vlr(Me.vlrCodUnidade, True)
         .SQL.Mais .Txt(Me.txtObservacao, True)
         .SQL.Mais .Vlr(f1.CmbIndex(Me.cmbSituacao), True)
         .SQL.Mais .Vlr(Me.vlrLucroMin, True)
         .SQL.Mais .Vlr(f1.CmbIndex(Me.cmbLucro), True)
         .SQL.Mais "GETDATE(), " 'Data do Servidor
         .SQL.Mais .Vlr(Prj.Sistema.IdUsuario, True)
         .SQL.Mais .Txt(Left(Me.cmbControlaEst, 1), True)
         .SQL.Mais .Vlr(Me.vlrEstMin, True)
         .SQL.Mais .Txt(Left(Me.cmbVenderSemEst, 1), True)
         .SQL.Mais .Vlr(f1.CmbIndex(Me.cmbTributacao), True)
         .SQL.Mais .Vlr(Me.vlrICMS, True)
         .SQL.Mais .Vlr(f1.CmbIndex(Me.cmbICMS), True)
         .SQL.Mais .Vlr(Me.vlrPISCOFINS, True)
         .SQL.Mais .Vlr(f1.CmbIndex(Me.cmbPISCOFINS), True)
         .SQL.Mais .Vlr(Me.vlrIPI, True)
         .SQL.Mais .Vlr(f1.CmbIndex(Me.cmbIPI), True)
         .SQL.Mais .Vlr(Me.vlrTributacao, True)
         .SQL.Mais .Vlr(Me.vlrFrete, True)
         .SQL.Mais .Vlr(f1.CmbIndex(Me.cmbFrete), True)
         .SQL.Mais .Vlr(Me.vlrComissao, True)
         .SQL.Mais .Vlr(f1.CmbIndex(Me.cmbComissao), True)
         .SQL.Mais .Vlr(Me.vlrMargem, True)
         .SQL.Mais .Vlr(f1.CmbIndex(Me.cmbMargem), True)
         .SQL.Mais .Vlr(Me.vlrCusto, True)
         .SQL.Mais .Vlr(f1.CmbIndex(Me.cmbCustos), True)
         .SQL.Mais .Vlr(Me.vlrVenda)
         .SQL.Mais " )"
         
         If Not .Executar(.SQL.Texto) Then
            clsErro.Transferir = .TransferirErro
            clsConexao.RollBack
            Exibir clsErro, "cmdSalvar_Click"
            GoTo DestruirObjetos
         End If
         
         strMsg = "Registro inserido com sucesso!" & vbNewLine & "C�digo: " & curSeqCod
      End With
         
   Case "&Alterar"
      curSeqCod = Me.vlrCod
      
      With clsConexao
         .SQL.Limpar
         .SQL.Mais " UPDATE Produtos SET "
         .SQL.Mais "    CodBarra = " & .Txt(Me.txtCodBarras, True)
         .SQL.Mais "    Descricao = " & .Txt(Me.txtDescricao, True)
         .SQL.Mais "    Abreviatura = " & .Txt(Me.txtAbreviatura, True)
         .SQL.Mais "    CodGrupo = " & .Vlr(Me.vlrCodGrupo, True)
         .SQL.Mais "    CodSubGrupo = " & .Vlr(Me.vlrCodSubGrupo, True)
         .SQL.Mais "    CodMarca = " & .Vlr(Me.vlrCodMarca, True)
         .SQL.Mais "    CodModelo = " & .Vlr(Me.vlrCodModelo, True)
         .SQL.Mais "    CodUnidade = " & .Vlr(Me.vlrCodUnidade, True)
         .SQL.Mais "    Obs = " & .Txt(Me.txtObservacao, True)
         .SQL.Mais "    IdSituacao = " & .Vlr(f1.CmbIndex(Me.cmbSituacao), True)
         .SQL.Mais "    LucroMin = " & .Vlr(Me.vlrLucroMin, True)
         .SQL.Mais "    IdLucroMin = " & .Vlr(f1.CmbIndex(Me.cmbLucro), True)
         .SQL.Mais "    DtUltAlt = GETDATE(), "
         .SQL.Mais "    CodUsuarioUltAlt = " & .Vlr(Prj.Sistema.IdUsuario, True)
         .SQL.Mais "    IdControlaEst = " & .Vlr(f1.CmbIndex(Me.cmbControlaEst), True)
         .SQL.Mais "    EstoqueMin = " & .Vlr(Me.vlrEstMin, True)
         .SQL.Mais "    IdVenderSemEst = " & .Vlr(f1.CmbIndex(Me.cmbVenderSemEst), True)
         .SQL.Mais "    IdTributacao = " & .Vlr(f1.CmbIndex(Me.cmbTributacao), True)
         .SQL.Mais "    ICMS = " & .Vlr(Me.vlrICMS, True)
         .SQL.Mais "    IdICMS = " & .Vlr(f1.CmbIndex(Me.cmbICMS), True)
         .SQL.Mais "    PISCOFINS = " & .Vlr(Me.vlrPISCOFINS, True)
         .SQL.Mais "    IdPISCOFINS = " & .Vlr(f1.CmbIndex(Me.cmbPISCOFINS), True)
         .SQL.Mais "    IPI = " & .Vlr(Me.vlrIPI, True)
         .SQL.Mais "    IdIPI = " & .Vlr(f1.CmbIndex(Me.cmbIPI), True)
         .SQL.Mais "    Tributos = " & .Vlr(Me.vlrTributacao, True)
         .SQL.Mais "    Frete = " & .Vlr(Me.vlrFrete, True)
         .SQL.Mais "    IdFrete = " & .Vlr(f1.CmbIndex(Me.cmbFrete), True)
         .SQL.Mais "    Comissao = " & .Vlr(Me.vlrComissao, True)
         .SQL.Mais "    IdComissao = " & .Vlr(f1.CmbIndex(Me.cmbComissao), True)
         .SQL.Mais "    Margem = " & .Vlr(Me.vlrMargem, True)
         .SQL.Mais "    IdMargem = " & .Vlr(f1.CmbIndex(Me.cmbMargem), True)
         .SQL.Mais "    Custos = " & .Vlr(Me.vlrCusto, True)
         .SQL.Mais "    IdCustos = " & .Vlr(f1.CmbIndex(Me.cmbCustos), True)
         .SQL.Mais "    ValorVenda = " & .Vlr(Me.vlrVenda)
         .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
         .SQL.Mais " AND Codigo = " & .Vlr(curSeqCod)
         
         If Not .Executar(.SQL.Texto) Then
            clsErro.Transferir = .TransferirErro
            clsConexao.RollBack
            Exibir clsErro, "cmdSalvar_Click"
            GoTo DestruirObjetos
         End If
         
         strMsg = "Registro alterado com sucesso!"
      End With
   End Select
   
   'Imagem (Excluir)
   With clsConexao
      .SQL.Limpar
      .SQL.Mais " DELETE FROM ProdutosPathImg "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND CodProduto = " & .Vlr(curSeqCod)
      
      If Not .Executar(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         clsConexao.RollBack
         Exibir clsErro, "cmdSalvar_Click"
         GoTo DestruirObjetos
      End If
   End With
   
   'Imagem (Inserir)
   With clsConexao
      .SQL.Limpar
      .SQL.Mais " INSERT INTO ProdutosPathImg ( "
      .SQL.Mais "    Empresa, CodProduto, PathImg "
      .SQL.Mais " ) VALUES ( "
      .SQL.Mais .Txt(Prj.Sistema.IdEmpresa, True)
      .SQL.Mais .Vlr(curSeqCod, True)
      .SQL.Mais .Txt(Me.img.Picture)
      .SQL.Mais " )"
      
      If Not .Executar(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         clsConexao.RollBack
         Exibir clsErro, "cmdSalvar_Click"
         GoTo DestruirObjetos
      End If
   End With
      
   'Atualiza a tabela de fornecedores
   For Each ContainerForn In colFornecedores
      With clsConexao
         .SQL.Limpar
         
         'Faz os procedimentos conforme situa��o atual do registro
         Select Case ContainerForn.Situacao
         Case "I" 'Inserido
            .SQL.Mais " INSERT INTO ProdutosForn ( "
            .SQL.Mais "    Empresa, CodFornecedor, CodProduto, Lucro, ValorCompra, ValorLucro "
            .SQL.Mais " ) VALUES ( "
            .SQL.Mais .Txt(Prj.Sistema.IdEmpresa, True)
            .SQL.Mais .Vlr(ContainerForn.CodFornecedor, True)
            .SQL.Mais .Vlr(curSeqCod, True)
            .SQL.Mais .Vlr(ContainerForn.Lucro, True)
            .SQL.Mais .Vlr(ContainerForn.ValorCompra, True)
            .SQL.Mais .Vlr(ContainerForn.LucroValor)
            .SQL.Mais " )"
            
         Case "A" 'Alterado
            .SQL.Mais " UPDATE ProdutosForn SET "
            .SQL.Mais "    Lucro = " & .Vlr(ContainerForn.Lucro, True)
            .SQL.Mais "    ValorCompra = " & .Vlr(ContainerForn.ValorCompra, True)
            .SQL.Mais "    ValorLucro = " & .Vlr(ContainerForn.LucroValor)
            .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
            .SQL.Mais " AND CodFornecedor = " & .Vlr(ContainerForn.CodFornecedor)
            .SQL.Mais " AND CodProduto = " & .Vlr(curSeqCod)
         
         Case "E" 'Removido
            .SQL.Mais " DELETE FROM ProdutosForn "
            .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
            .SQL.Mais " AND CodFornecedor = " & .Vlr(ContainerForn.CodFornecedor)
            .SQL.Mais " AND CodProduto = " & .Vlr(curSeqCod)
         
         Case "G" 'Gravado
            'N�o realiza nenhuma a��o para esta situa��o
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
   GoTo DestruirObjetos
      
CmpObrig:
   Exibir clsErro, "cmdSalvar_Click"

DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
   Set ContainerForn = Nothing
   Ampulheta False
End Sub

Private Sub Form_Load()
   On Error GoTo Form_Load_E
   
   Ampulheta True
   
   Set clsErro = CreateObject("INF_Erro.Funcoes")
   Set colFornecedores = New Collection
   
   f1.FormCentralizar Me
   
   If Not HabilitarBotao(clsErro, Me, Me.cmdNovo) Then Exibir clsErro, "Form_Load"
   If Not HabilitarBotao(clsErro, Me, Me.cmdConsultar) Then Exibir clsErro, "Form_Load"
   
   f1.CmbAdd Me.cmbSituacao, "Desativado", 0
   f1.CmbAdd Me.cmbSituacao, "Ativado", 1
   Me.cmbSituacao.ListIndex = f1.CmbValor(Me.cmbSituacao, 1)
   
   mCmbValorPerc Me.cmbLucro
   mCmbSimNao Me.cmbControlaEst
   mCmbSimNao Me.cmbVenderSemEst
   
   mCmbValorPerc Me.cmbICMS
   mCmbValorPerc Me.cmbPISCOFINS
   mCmbValorPerc Me.cmbIPI
   
   mCmbValorPerc Me.cmbFrete
   mCmbValorPerc Me.cmbComissao
   mCmbValorPerc Me.cmbMargem
   mCmbValorPerc Me.cmbCustos
   
   f1.CmbAdd Me.cmbTributacao, "Isento", 0
   f1.CmbAdd Me.cmbTributacao, "Tributado", 1
   Me.cmbTributacao.ListIndex = f1.CmbValor(Me.cmbTributacao, 0)
   
   Me.img.Picture = LoadPicture(PATH_IMG & "SEM_IMG.JPG")
   
   Me.lstLegenda.ListItems.Clear
   Me.lstLegenda.ListItems.Add , "K_1", "Gravado", , LST_ICO_GRAVADO
   Me.lstLegenda.ListItems.Add , "K_2", "Inserido", , LST_ICO_INSERIDO
   Me.lstLegenda.ListItems.Add , "K_3", "Alterado", , LST_ICO_ALTERADO
   Me.lstLegenda.ListItems.Add , "K_4", "Removido", , LST_ICO_REMOVIDO
      
   GoTo DestruirObjetos

Form_Load_E:
   clsErro.Salvar Err
   Exibir clsErro, "Form_Load"

DestruirObjetos:
   Ampulheta False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set clsErro = Nothing
   f1.CollectionLimpar colFornecedores
   Set colFornecedores = Nothing
End Sub

Private Sub AtualizaLista()
   On Error GoTo VerificaFornecedor_E
   
   Dim strKey As String
   Dim itemX As ListItem
   Dim curMaiorValor As Currency
   Dim clsCursor As INF_Cursor.Cursor
   Dim clsContainer As sContainerProdutosForn
   
   'Remove os itens da lista
   Me.lstFornecedor.ListItems.Clear
   
   'Verifica qual fornecedor, vende o produto mais caro
   For Each clsContainer In colFornecedores
      If curMaiorValor < clsContainer.ValorCompra Then
         curMaiorValor = clsContainer.ValorCompra
         strKey = clsContainer.Key
      End If
   Next
   
   'Seta o container do fornecedor mais caro
   Set clsContainer = colFornecedores(strKey)
   
   'Atribui o lucro minimo em cima do valor de compra do fornecedor mais caro
   Me.vlrVenda = ((clsContainer.ValorCompra * (IIf(Me.cmbLucro.Text = "%", (Me.vlrLucroMin / 100), Me.vlrLucroMin))) + clsContainer.ValorCompra)
   
   'Seta os valores para o fornecedor mais caro
   Set clsCursor = CreateObject("INF_Cursor.Cursor")
   With clsCursor
      .Inicializar clsConexao
      
      .SQL.Limpar
      .SQL.Mais " SELECT LucroPerc, LucroVlr "
      .SQL.Mais " FROM VerificaLucroFornecedor "
      .SQL.Mais " ( "
      .SQL.Mais .Vlr(Me.vlrVenda, True)
      .SQL.Mais .Vlr(clsContainer.ValorCompra)
      .SQL.Mais " ) "
      
      If Not .Abrir(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         clsErro.ModRotina = "VerificaFornecedor"
         GoTo DestruirObjetos
      End If
      
      If Not .EOF Then
         clsContainer.Lucro = .Valor("LucroPerc")
         clsContainer.LucroValor = .Valor("LucroVlr")
      End If
      .Fechar
   End With
   Set clsContainer = Nothing
   
   'Navega entre os registros para atualizar os valores de lucro e percentual
   For Each clsContainer In colFornecedores
   
      'S� atualiza os demais registros
      If Not clsContainer.Key = strKey Then
         With clsCursor
            .SQL.Limpar
            .SQL.Mais " SELECT LucroPerc, LucroVlr "
            .SQL.Mais " FROM VerificaLucroFornecedor "
            .SQL.Mais " ( "
            .SQL.Mais .Vlr(Me.vlrVenda, True)
            .SQL.Mais .Vlr(clsContainer.ValorCompra)
            .SQL.Mais " ) "
      
            If Not .Abrir(.SQL.Texto) Then
               clsErro.Transferir = .TransferirErro
               clsErro.ModRotina = "VerificaFornecedor"
               GoTo DestruirObjetos
            End If
      
            If Not .EOF Then
               clsContainer.Lucro = .Valor("LucroPerc")
               clsContainer.LucroValor = .Valor("LucroVlr")
            End If
            .Fechar
         End With
      End If
   Next
   
   'Adiciona tudo na lista
   For Each clsContainer In colFornecedores
      Set itemX = Me.lstFornecedor.ListItems.Add(, clsContainer.Key, clsContainer.CodFornecedor)
      itemX.SubItems(LST_COL_DSCFORN) = clsContainer.DscFornecedor
      itemX.SubItems(LST_COL_TELFORN) = clsContainer.Telefone
      itemX.SubItems(LST_COL_LUCROFORN) = Format(clsContainer.Lucro, "000.00") & " %"
      itemX.SubItems(LST_cOL_VALORCOMPRA) = Format(clsContainer.ValorCompra, "#.00")
      itemX.SubItems(LST_COL_VALORLUCRO) = Format(clsContainer.LucroValor, "#.00")
      
      Select Case clsContainer.Situacao
      Case "I"
         itemX.SmallIcon = LST_ICO_INSERIDO
      Case "E"
         itemX.SmallIcon = LST_ICO_REMOVIDO
      Case "A"
         itemX.SmallIcon = LST_ICO_ALTERADO
      Case "G"
         itemX.SmallIcon = LST_ICO_GRAVADO
      End Select
      
      itemX.Selected = True
   Next
      
   GoTo DestruirObjetos

VerificaFornecedor_E:
   clsErro.Salvar Err
   clsErro.ModRotina = "VerificaFornecedor"

DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Sub
