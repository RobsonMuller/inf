VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "InfinityControl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   9285
   Tag             =   "20402"
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8025
      Top             =   3105
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
            Picture         =   "Cliente.frx":0000
            Key             =   "Removido"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":015A
            Key             =   "Inserido"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":02B4
            Key             =   "Gravado"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":040E
            Key             =   "Alterado"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraContatos 
      Caption         =   "Contatos"
      Enabled         =   0   'False
      Height          =   1785
      Left            =   60
      TabIndex        =   28
      Top             =   4380
      Width           =   7770
      Begin VB.CommandButton cmdAdicionar 
         Caption         =   "<<"
         Height          =   315
         Left            =   7275
         TabIndex        =   30
         ToolTipText     =   "Adicionar Fornecedor."
         Top             =   240
         Width           =   345
      End
      Begin VB.CommandButton cmdRemover 
         Caption         =   ">>"
         Height          =   315
         Left            =   7275
         TabIndex        =   32
         ToolTipText     =   "Remover Fornecedor."
         Top             =   930
         Width           =   345
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "="
         Height          =   315
         Left            =   7275
         TabIndex        =   31
         ToolTipText     =   "Adicionar Fornecedor."
         Top             =   585
         Width           =   345
      End
      Begin MSComctlLib.ListView lstContatos 
         Height          =   1440
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   7065
         _ExtentX        =   12462
         _ExtentY        =   2540
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
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
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Novo"
      Enabled         =   0   'False
      Height          =   345
      Left            =   7905
      TabIndex        =   33
      Top             =   90
      Width           =   1320
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "&Consultar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   7905
      TabIndex        =   34
      Top             =   465
      Width           =   1320
   End
   Begin VB.CommandButton cmdSalvar 
      Cancel          =   -1  'True
      Caption         =   "&Salvar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   7905
      TabIndex        =   35
      Tag             =   "20102"
      Top             =   840
      Width           =   1320
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Enabled         =   0   'False
      Height          =   345
      Left            =   7905
      TabIndex        =   36
      Tag             =   "6"
      Top             =   1215
      Width           =   1320
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "&Limpar"
      Height          =   345
      Left            =   7905
      TabIndex        =   37
      Top             =   1590
      Width           =   1320
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   345
      Left            =   7905
      TabIndex        =   38
      Top             =   1965
      Width           =   1320
   End
   Begin VB.Frame fraParametros 
      Caption         =   "Parâmetros"
      Enabled         =   0   'False
      Height          =   3720
      Left            =   60
      TabIndex        =   3
      Top             =   660
      Width           =   7755
      Begin VB.ComboBox cmbTipoPessoa 
         Height          =   315
         Left            =   3735
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1005
         Width           =   1650
      End
      Begin VB.ComboBox cmbSituacao 
         Height          =   315
         Left            =   6585
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1005
         Width           =   1020
      End
      Begin VB.CommandButton cmdConsultarCEP 
         Height          =   315
         Left            =   2205
         Picture         =   "Cliente.frx":0568
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Consultar CEP"
         Top             =   1005
         Width           =   330
      End
      Begin VB.ComboBox cmbUf 
         Height          =   315
         Left            =   6585
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   2145
         Width           =   1020
      End
      Begin VB.CommandButton cmdPesqCidade 
         Caption         =   "..."
         Height          =   255
         Left            =   2205
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1830
         Width           =   270
      End
      Begin rdActiveText.ActiveText vlrCEP 
         Height          =   315
         Left            =   1005
         TabIndex        =   8
         Top             =   1005
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
      Begin rdActiveText.ActiveText txtBairro 
         Height          =   315
         Left            =   1005
         TabIndex        =   20
         Top             =   2145
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
         Left            =   2535
         TabIndex        =   16
         Top             =   1770
         Width           =   5070
         _ExtentX        =   8943
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
         Left            =   1005
         TabIndex        =   13
         Top             =   1770
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
         Left            =   6585
         TabIndex        =   12
         Top             =   1395
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
      Begin rdActiveText.ActiveText txtEndereco 
         Height          =   315
         Left            =   1005
         TabIndex        =   11
         Top             =   1395
         Width           =   4350
         _ExtentX        =   7673
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
      Begin rdActiveText.ActiveText datCadastro 
         Height          =   315
         Left            =   6585
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
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
         Left            =   1005
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtNome 
         Height          =   315
         Left            =   1005
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
      Begin rdActiveText.ActiveText vlrIE 
         Height          =   315
         Left            =   4905
         TabIndex        =   7
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtEmail 
         Height          =   315
         Left            =   1005
         TabIndex        =   26
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
         TextCase        =   2
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtTelefone 
         Height          =   315
         Left            =   1005
         TabIndex        =   23
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
         MaxLength       =   11
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtFax 
         Height          =   315
         Left            =   3585
         TabIndex        =   24
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
         MaxLength       =   11
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtCelular 
         Height          =   315
         Left            =   5850
         TabIndex        =   25
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
         MaxLength       =   11
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtSite 
         Height          =   315
         Left            =   1005
         TabIndex        =   27
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
         TextCase        =   2
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin VB.Label lblTipoPessoa 
         Caption         =   "Tp. Pessoa:"
         Height          =   255
         Left            =   2745
         TabIndex        =   54
         Top             =   1035
         Width           =   1230
      End
      Begin VB.Label lblTelefone 
         Caption         =   "Telefone:"
         Height          =   195
         Left            =   120
         TabIndex        =   52
         Top             =   2535
         Width           =   1125
      End
      Begin VB.Label lblEmail 
         Caption         =   "E-mail:"
         Height          =   210
         Left            =   120
         TabIndex        =   51
         Top             =   2925
         Width           =   1095
      End
      Begin VB.Label lblSite 
         Caption         =   "Site:"
         Height          =   240
         Left            =   120
         TabIndex        =   50
         Top             =   3300
         Width           =   1110
      End
      Begin VB.Label lblFax 
         Caption         =   "Fax:"
         Height          =   195
         Left            =   3030
         TabIndex        =   49
         Top             =   2535
         Width           =   885
      End
      Begin VB.Label lblCelular 
         Caption         =   "Cel.:"
         Height          =   195
         Left            =   5445
         TabIndex        =   48
         Top             =   2550
         Width           =   885
      End
      Begin VB.Label lblSituação 
         Caption         =   "Situação:"
         Height          =   240
         Left            =   5445
         TabIndex        =   47
         Top             =   1035
         Width           =   1050
      End
      Begin VB.Label lblNome 
         Caption         =   "Nome:"
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
         TabIndex        =   46
         Top             =   270
         Width           =   1110
      End
      Begin VB.Label lblCNPJ 
         Caption         =   "CPF:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   45
         Top             =   645
         Width           =   1125
      End
      Begin VB.Label lblIE 
         Caption         =   "RG:"
         Height          =   225
         Left            =   3810
         TabIndex        =   44
         Top             =   645
         Width           =   1080
      End
      Begin VB.Label lblDataCadastro 
         Caption         =   "Dt. Cadastro:"
         Height          =   180
         Left            =   5445
         TabIndex        =   43
         Top             =   270
         Width           =   1395
      End
      Begin VB.Label lblCEP 
         Caption         =   "CEP"
         Height          =   210
         Left            =   120
         TabIndex        =   41
         Top             =   1050
         Width           =   1200
      End
      Begin VB.Label lblEndereco 
         Caption         =   "Endereço:"
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
         Left            =   120
         TabIndex        =   40
         Top             =   1425
         Width           =   1155
      End
      Begin VB.Label lblBairro 
         Caption         =   "Bairro:"
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
         Left            =   120
         TabIndex        =   21
         Top             =   2175
         Width           =   1140
      End
      Begin VB.Label lblNumero 
         Caption         =   "Número:"
         Height          =   270
         Left            =   5445
         TabIndex        =   19
         Top             =   1440
         Width           =   1140
      End
      Begin VB.Label lblCidade 
         Caption         =   "Cidade:"
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
         TabIndex        =   18
         Top             =   1800
         Width           =   1170
      End
      Begin VB.Label lblUF 
         Caption         =   "Estado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5445
         TabIndex        =   17
         Top             =   2175
         Width           =   990
      End
   End
   Begin VB.Frame fraIdentificacao 
      Caption         =   "Identificação"
      Height          =   660
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   7755
      Begin VB.CommandButton cmdPesqForn 
         Caption         =   "..."
         Height          =   255
         Left            =   2070
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   315
         Width           =   270
      End
      Begin rdActiveText.ActiveText vlrCod 
         Height          =   315
         Left            =   855
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
         TabIndex        =   42
         Top             =   270
         Width           =   1065
      End
   End
   Begin MSComctlLib.ListView lstLegenda 
      Height          =   1065
      Left            =   7905
      TabIndex        =   39
      Top             =   5100
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
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
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
      Left            =   7905
      TabIndex        =   53
      Top             =   4875
      Width           =   1215
   End
End
Attribute VB_Name = "frmCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lngSeqLst As Long
Private colContatos As Collection
Private clsErro As INF_Erro.Funcoes

Private Const CONST_LST_NOME As Integer = 0
Private Const CONST_LST_TELEFONE As Integer = 1
Private Const CONST_LST_EMAIL As Integer = 2

Private Sub cmbTipoPessoa_Click()
   Select Case f1.CmbParametro(Me.cmbTipoPessoa)
   Case 0
      Me.vlrCPFCNPJ.MaxLength = 11
      Me.lblCNPJ.Caption = "CPF:"
   
      Me.lblIE.Caption = "RG:"
      Me.vlrIE.MaxLength = 10
   Case 1
      Me.vlrCPFCNPJ.Mask = 14
      Me.lblCNPJ.Caption = "CNPJ:"
      
      Me.vlrIE.MaxLength = 10
      Me.lblIE.Caption = "IE:"
   End Select
End Sub

Private Sub cmdAdicionar_Click()
   On Error GoTo cmdAdicionar_Click_E
   
   Dim itemX As ListItem
   Dim frmModal As frmClientesContato
   Dim clsContainerContatos As sContainerClientesContatos
      
   Set frmModal = New frmClientesContato
   With frmModal
      .Show vbModal
      
      If Not .Cancelado Then
         lngSeqLst = lngSeqLst + 1
         
         Set clsContainerContatos = New sContainerClientesContatos
         clsContainerContatos.Sequencial = lngSeqLst
         clsContainerContatos.Nome = .Contato
         clsContainerContatos.Telefone = .Telefone
         clsContainerContatos.Email = .Email
         clsContainerContatos.Acao = "I"
         colContatos.Add clsContainerContatos, clsContainerContatos.Key
         
         Set itemX = Me.lstContatos.ListItems.Add(, clsContainerContatos.Key, clsContainerContatos.Nome)
         itemX.SubItems(CONST_LST_TELEFONE) = clsContainerContatos.Telefone
         itemX.SubItems(CONST_LST_EMAIL) = clsContainerContatos.Email
         itemX.SmallIcon = LST_ICO_INSERIDO
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
   Dim clsContainer As sContainerClientesContatos

   If Not mCmpObrigatorio(clsErro, Me.vlrCod, "Código") Then GoTo Erro_Msg
   
   Set clsCursor = CreateObject("INF_Cursor.Cursor")
   With clsCursor
      .Inicializar clsConexao
      
      .SQL.Limpar
      .SQL.Mais " SELECT Empresa, Codigo, Nome, DataCad, CPFCNPJ, RGIE, CEP, TpPessoa, "
      .SQL.Mais "    Situacao, Endereco, Numero, CodCidade, Bairro, Estado, Telefone, "
      .SQL.Mais "    Fax, Cel, Email, Site "
      .SQL.Mais " FROM Clientes "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND Codigo = " & .Vlr(Me.vlrCod)
      
      If Not .Abrir(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         Exibir clsErro, "cmdConsultar_Click"
         GoTo DestruirObjetos
      End If
      
      If Not .EOF Then
         Me.txtNome = .Valor("Nome")
         Me.datCadastro = .Valor("DataCad")
         Me.vlrCPFCNPJ = .Valor("CPFCNPJ")
         Me.vlrIE = .Valor("RGIE")
         Me.vlrCEP = .Valor("CEP")
         Me.txtEndereco = .Valor("Endereco")
         Me.vlrNro = .Valor("Numero")
         Me.vlrCodCidade = .Valor("CodCidade")
         Me.txtBairro = .Valor("Bairro")
         Me.cmbUf.Text = .Valor("Estado")
         Me.txtTelefone = .Valor("Telefone")
         Me.txtFax = .Valor("Fax")
         Me.txtCelular = .Valor("Cel")
         Me.txtEmail = .Valor("Email")
         Me.txtSite = .Valor("Site")
         
         Me.cmbSituacao.ListIndex = f1.CmbValor(Me.cmbSituacao, .Valor("Situacao"))
         Me.cmbTipoPessoa.ListIndex = f1.CmbValor(Me.cmbTipoPessoa, .Valor("TpPessoa"))
         
         Me.fraIdentificacao.Enabled = False
         Me.fraParametros.Enabled = True
         Me.fraContatos.Enabled = True
         
         Me.cmdNovo.Enabled = False
         Me.cmdConsultar.Enabled = False
         
         Me.cmdSalvar.Caption = "&Alterar"
         If Not HabilitarBotao(clsErro, Me, Me.cmdSalvar) Then Exibir clsErro, "cmdConsultar_Click"
         If Not HabilitarBotao(clsErro, Me, Me.cmdExcluir) Then Exibir clsErro, "cmdConsultar_Click"
      Else
         mMsgInfo "Usuário não localizado! Verifique."
         mFocus Me.vlrCod
         GoTo DestruirObjetos
      End If
   End With
      
   With clsCursor
      .SQL.Limpar
      .SQL.Mais " SELECT Codigo, CodCliente, Nome, Telefone, Email "
      .SQL.Mais " FROM ClientesContatos "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND CodCliente = " & .Vlr(Me.vlrCod)
      
      If Not .Abrir(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         Exibir clsErro, "cmdConsultar_Click"
         GoTo DestruirObjetos
      End If
      
      Do Until .EOF
         lngSeqLst = lngSeqLst + 1
         
         'Armazena na collection
         Set clsContainer = New sContainerClientesContatos
         clsContainer.Sequencial = lngSeqLst
         clsContainer.CodContato = .Valor("Codigo")
         clsContainer.Nome = .Valor("Nome")
         clsContainer.Telefone = .Valor("Telefone")
         clsContainer.Email = .Valor("Email")
         colContatos.Add clsContainer, clsContainer.Key
         
         'Armazena na lista
         Set itemX = Me.lstContatos.ListItems.Add(, clsContainer.Key, clsContainer.Nome)
         itemX.SubItems(CONST_LST_TELEFONE) = clsContainer.Telefone
         itemX.SubItems(CONST_LST_EMAIL) = clsContainer.Email
         itemX.SmallIcon = LST_ICO_GRAVADO
         itemX.Selected = True
         
         .ProximoRegistro
      Loop
   End With
   
   mFocus Me.txtEndereco
   
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
   
   Dim strEndereco As String
   Dim url As String
   Dim clsCursor As INF_Cursor.Cursor
   Dim objXml As MSXML2.DOMDocument
   Dim objXmlNodes As MSXML2.IXMLDOMElement
   Dim ObjXmlElement As MSXML2.IXMLDOMElement
   
   If Len(Trim(Me.vlrCEP)) = 0 Then
      mMsgInfo "O número do CEP deve ser informado! Verifique."
      mFocus Me.vlrCEP
      GoTo DestruirObjetos
   End If
   
   If Not f1.VerificaConexaoInternet(clsErro) Then
      Exibir clsErro, "cmdConsultarCEP_Click"
      mFocus Me.vlrCEP
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
   
   If Not mMsgPerg("Você deseja excluir o registro atual ?") Then Exit Sub
   
   clsConexao.Begin
   With clsConexao
      .SQL.Limpar
      .SQL.Mais " DELETE FROM ClientesContatos "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND CodCliente = " & .Vlr(Me.vlrCod)
      
      If Not .Executar(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         clsConexao.RollBack
         Exibir clsErro, "cmdExcluir_Click"
         Exit Sub
      End If
   
      .SQL.Limpar
      .SQL.Mais " DELETE FROM Clientes "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND Codigo = " & .Vlr(Me.vlrCod)
      
      If Not .Executar(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         clsConexao.RollBack
         Exibir clsErro, "cmdExcluir_Click"
         Exit Sub
      End If
   End With
   
   clsConexao.Commit
   
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
   
   lngSeqLst = 0
   
   If Not HabilitarBotao(clsErro, Me, Me.cmdNovo) Then Exibir clsErro, "Form_Load"
   If Not HabilitarBotao(clsErro, Me, Me.cmdConsultar) Then Exibir clsErro, "Form_Load"
   
   Me.cmdSalvar.Caption = "&Salvar"
   Me.cmdSalvar.Enabled = False
   Me.cmdExcluir.Enabled = False
   
   Me.lstContatos.ListItems.Clear
   f1.CollectionLimpar colContatos
      
   Me.fraParametros.Enabled = False
   Me.fraIdentificacao.Enabled = True
   Me.fraContatos.Enabled = False
   
   mFocus Me.vlrCod
End Sub

Private Sub cmdModificar_Click()
   On Error GoTo cmdModificar_Click_E
   
   Dim itemX As ListItem
   Dim frmModal As frmClientesContato
   Dim clsContainerContatos As sContainerClientesContatos
   
   Set itemX = Me.lstContatos.SelectedItem
   If itemX Is Nothing Then Exit Sub
   If Not itemX.Selected Then Exit Sub
   
   Set clsContainerContatos = colContatos(itemX.Key)
   Set frmModal = New frmClientesContato
   With frmModal
      .Contato = clsContainerContatos.Nome
      .Telefone = clsContainerContatos.Telefone
      .Email = clsContainerContatos.Email
      
      .Show vbModal
      
      If Not .Cancelado Then
         clsContainerContatos.Nome = .Contato
         clsContainerContatos.Telefone = .Telefone
         clsContainerContatos.Email = .Email
                  
         itemX.Text = clsContainerContatos.Nome
         itemX.SubItems(CONST_LST_TELEFONE) = clsContainerContatos.Telefone
         itemX.SubItems(CONST_LST_EMAIL) = clsContainerContatos.Email
         
         If Not clsContainerContatos.Acao = "I" Then
            clsContainerContatos.Acao = "A"
            itemX.SmallIcon = LST_ICO_ALTERADO
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
   Me.vlrCod = 0
   
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
      .Codigo = Me.vlrCodCidade
      .TpDefinicao = enMunicipios
      .Show vbModal
      
      If Not .Cancelado Then Me.vlrCodCidade = .Codigo
      Me.txtDescCidade = .Descricao
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
      .TpDefinicao = enClientes
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
   Dim frmModal As frmClientesContato
   Dim clsContainerContatos As sContainerClientesContatos
   
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

   Dim strMsg As String
   Dim lngNewCod As Long
   Dim clsCursor As INF_Cursor.Cursor
   Dim clsContainer As sContainerClientesContatos
   
   If Not mCmpObrigatorio(clsErro, Me.txtNome, "Nome") Then GoTo Erro_Msg
   If Not mCmpObrigatorio(clsErro, Me.vlrCPFCNPJ, IIf(Left(Me.cmbTipoPessoa.Text, 1) = "F", "CPF", "CNPJ")) Then GoTo Erro_Msg
   If Not mCmpObrigatorio(clsErro, Me.txtEndereco, "Endereço") Then GoTo Erro_Msg
   If Not mCmpObrigatorio(clsErro, Me.vlrCodCidade, "Cidade") Then GoTo Erro_Msg
   If Not mCmpObrigatorio(clsErro, Me.txtBairro, "Bairro") Then GoTo Erro_Msg
   
   clsConexao.Begin
   Select Case Me.cmdSalvar.Caption
   Case "&Inserir"
      Set clsCursor = CreateObject("INF_Cursor.Cursor")
      With clsCursor
         .Inicializar clsConexao
         
         .SQL.Limpar
         .SQL.Mais " SELECT (ISNULL(MAX(Codigo), 0) + 1) AS MaxCod "
         .SQL.Mais " FROM Clientes "
         .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
         
         If Not .Abrir(.SQL.Texto) Then
            clsErro.Transferir = .TransferirErro
            clsConexao.RollBack
            Exibir clsErro, "cmdSalvar_Click"
            GoTo DestruirObjetos
         End If
         
         lngNewCod = .Valor("MaxCod")
         .Fechar
      End With
      
      With clsConexao
         .SQL.Limpar
         .SQL.Mais " INSERT INTO Clientes ( "
         .SQL.Mais "    Empresa, Codigo, Nome, DataCad, CPFCNPJ, RGIE, CEP, TpPessoa, Situacao, Endereco, "
         .SQL.Mais "    Numero, CodCidade, Bairro, Estado, Telefone, Fax, Cel, Email, Site "
         .SQL.Mais " ) VALUES ( "
         .SQL.Mais .Txt(Prj.Sistema.IdEmpresa, True)
         .SQL.Mais .Vlr(lngNewCod, True)
         .SQL.Mais .Txt(Me.txtNome, True)
         .SQL.Mais " GETDATE(), "
         .SQL.Mais .Txt(Me.vlrCPFCNPJ, True)
         .SQL.Mais .Txt(Me.vlrIE, True)
         .SQL.Mais .Txt(Me.vlrCEP, True)
         .SQL.Mais .Txt(f1.CmbParametro(Me.cmbTipoPessoa), True)
         .SQL.Mais .Txt(f1.CmbParametro(Me.cmbSituacao), True)
         .SQL.Mais .Txt(Me.txtEndereco, True)
         .SQL.Mais .Vlr(Me.vlrNro, True)
         .SQL.Mais .Vlr(Me.vlrCodCidade, True)
         .SQL.Mais .Txt(Me.txtBairro, True)
         .SQL.Mais .Txt(Me.cmbUf.Text, True)
         .SQL.Mais .Txt(f1.SoNumeros(Me.txtTelefone), True)
         .SQL.Mais .Txt(f1.SoNumeros(Me.txtFax), True)
         .SQL.Mais .Txt(f1.SoNumeros(Me.txtCelular), True)
         .SQL.Mais .Txt(Me.txtEmail, True)
         .SQL.Mais .Txt(Me.txtSite)
         .SQL.Mais ")"
         
         If Not .Executar(.SQL.Texto) Then
            clsErro.Transferir = .TransferirErro
            clsConexao.RollBack
            Exibir clsErro, "cmdExcluir_Click"
            GoTo DestruirObjetos
         End If
      End With
      
      strMsg = "Registro inserido com sucesso! Código: " & lngNewCod
      
   Case "&Alterar"
      With clsConexao
         lngNewCod = Me.vlrCod
         
         .SQL.Limpar
         .SQL.Mais " UPDATE Clientes SET "
         .SQL.Mais "    Nome = " & .Txt(Me.txtNome, True)
         .SQL.Mais "    CPFCNPJ = " & .Txt(Me.vlrCPFCNPJ, True)
         .SQL.Mais "    RGIE = " & .Txt(Me.vlrIE, True)
         .SQL.Mais "    CEP = " & .Txt(Me.vlrCEP, True)
         .SQL.Mais "    TpPessoa = " & .Txt(f1.CmbParametro(Me.cmbTipoPessoa), True)
         .SQL.Mais "    Situacao = " & .Txt(f1.CmbParametro(Me.cmbSituacao), True)
         .SQL.Mais "    Endereco = " & .Txt(Me.txtEndereco, True)
         .SQL.Mais "    Numero = " & .Vlr(Me.vlrNro, True)
         .SQL.Mais "    CodCidade = " & .Vlr(Me.vlrCodCidade, True)
         .SQL.Mais "    Bairro = " & .Txt(Me.txtBairro, True)
         .SQL.Mais "    Estado = " & .Txt(Me.cmbUf.Text, True)
         .SQL.Mais "    Telefone = " & .Txt(f1.SoNumeros(Me.txtTelefone), True)
         .SQL.Mais "    Fax = " & .Txt(f1.SoNumeros(Me.txtFax), True)
         .SQL.Mais "    Cel = " & .Txt(f1.SoNumeros(Me.txtCelular), True)
         .SQL.Mais "    Email = " & .Txt(Me.txtEmail, True)
         .SQL.Mais "    Site = " & .Txt(Me.txtSite)
         .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
         .SQL.Mais " AND Codigo = " & .Vlr(lngNewCod)
         
         If Not .Executar(.SQL.Texto) Then
            clsErro.Transferir = .TransferirErro
            clsConexao.RollBack
            Exibir clsErro, "cmdExcluir_Click"
            GoTo DestruirObjetos
         End If
      End With
      
      strMsg = "Registro alterado com sucesso!"
   End Select
   
   For Each clsContainer In colContatos
      With clsConexao
         .SQL.Limpar
         
         Select Case clsContainer.Acao
         Case "I"
            .SQL.Mais " INSERT INTO ClientesContatos ( "
            .SQL.Mais "    Empresa, CodCliente, Nome, Telefone, Email "
            .SQL.Mais " ) VALUES ( "
            .SQL.Mais .Txt(Prj.Sistema.IdEmpresa, True)
            .SQL.Mais .Vlr(lngNewCod, True)
            .SQL.Mais .Txt(clsContainer.Nome, True)
            .SQL.Mais .Txt(f1.SoNumeros(clsContainer.Telefone), True)
            .SQL.Mais .Txt(clsContainer.Email)
            .SQL.Mais ")"
         Case "A"
            .SQL.Mais " UPDATE ClientesContatos SET "
            .SQL.Mais "    Nome = " & .Txt(clsContainer.Nome, True)
            .SQL.Mais "    Telefone = " & .Txt(clsContainer.Telefone, True)
            .SQL.Mais "    Email = " & .Txt(clsContainer.Email)
            .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
            .SQL.Mais " AND Codigo = " & .Vlr(clsContainer.CodContato)
            .SQL.Mais " AND CodCliente = " & .Vlr(lngNewCod)
         Case "E"
            .SQL.Mais " DELETE FROM ClientesContatos "
            .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
            .SQL.Mais " AND Codigo = " & .Vlr(clsContainer.CodContato)
            .SQL.Mais " AND CodCliente = " & .Vlr(lngNewCod)
         Case "G"
            'Gravado - Sem Ação
         End Select
         
         If Len(Trim(.SQL.Texto)) > 0 Then
            If Not .Executar(.SQL.Texto) Then
               clsErro.Transferir = .TransferirErro
               clsConexao.RollBack
               Exibir clsErro, "cmdExcluir_Click"
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

Erro_Msg:
   Exibir clsErro, "cmdSalvar_Click"

DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Sub

Private Sub Form_Load()
   Set clsErro = New INF_Erro.Funcoes
   Set colContatos = New Collection
   
   f1.FormCentralizar Me
   
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
   Me.cmbTipoPessoa.ListIndex = f1.CmbValor(Me.cmbTipoPessoa, 0, enCodigo)
   
   'Carrega Combo Situacao
   '0 = Ativo
   '1 = Inativo
   f1.CmbAdd Me.cmbSituacao, "Ativo", 0
   f1.CmbAdd Me.cmbSituacao, "Inativo", 1
   Me.cmbSituacao.ListIndex = f1.CmbValor(Me.cmbSituacao, 0)
      
   Me.lstLegenda.ListItems.Clear
   Me.lstLegenda.ListItems.Add , "K_1", "Gravado", , LST_ICO_GRAVADO
   Me.lstLegenda.ListItems.Add , "K_2", "Inserido", , LST_ICO_INSERIDO
   Me.lstLegenda.ListItems.Add , "K_3", "Alterado", , LST_ICO_ALTERADO
   Me.lstLegenda.ListItems.Add , "K_4", "Removido", , LST_ICO_REMOVIDO
         
   If Not HabilitarBotao(clsErro, Me, Me.cmdNovo) Then Exibir clsErro, "Form_Load"
   If Not HabilitarBotao(clsErro, Me, Me.cmdConsultar) Then Exibir clsErro, "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set clsErro = Nothing
   f1.CollectionLimpar colContatos
   Set colContatos = Nothing
End Sub

Private Sub vlrCodCidade_LostFocus()
   If Not mVerificaCidade(clsErro, Me.vlrCodCidade, Me.txtDescCidade, Me.cmbUf) Then
      Exibir clsErro, "vlrCodCidade_LostFocus"
      mFocus Me.vlrCodCidade
   End If
End Sub
