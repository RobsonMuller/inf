VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "InfinityControl.ocx"
Begin VB.Form frmEntrada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16530
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   16530
   Tag             =   "20301"
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Novo"
      Enabled         =   0   'False
      Height          =   345
      Left            =   15105
      TabIndex        =   79
      Top             =   105
      Width           =   1320
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "&Consultar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   15105
      TabIndex        =   78
      Top             =   480
      Width           =   1320
   End
   Begin VB.CommandButton cmdSalvar 
      Cancel          =   -1  'True
      Caption         =   "&Salvar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   15105
      TabIndex        =   77
      Top             =   855
      Width           =   1320
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Enabled         =   0   'False
      Height          =   345
      Left            =   15105
      TabIndex        =   76
      Tag             =   "6"
      Top             =   1230
      Width           =   1320
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "&Limpar"
      Height          =   345
      Left            =   15105
      TabIndex        =   75
      Top             =   1605
      Width           =   1320
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   345
      Left            =   15105
      TabIndex        =   74
      Top             =   1980
      Width           =   1320
   End
   Begin VB.Frame Frame4 
      Caption         =   "Informações de Cadastro"
      Height          =   645
      Left            =   8130
      TabIndex        =   68
      Top             =   15
      Width           =   6900
      Begin rdActiveText.ActiveText datCad 
         Height          =   315
         Left            =   1035
         TabIndex        =   69
         Top             =   225
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
         Locked          =   -1  'True
      End
      Begin rdActiveText.ActiveText txtUsuarioCad 
         Height          =   315
         Left            =   4185
         TabIndex        =   71
         Top             =   225
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
      Begin VB.Label lblUserCad 
         Caption         =   "Usuário Cad.:"
         Height          =   210
         Left            =   3075
         TabIndex        =   72
         Top             =   255
         Width           =   1230
      End
      Begin VB.Label lblDtCad 
         Caption         =   "Data Cad.:"
         Height          =   270
         Left            =   120
         TabIndex        =   70
         Top             =   255
         Width           =   1305
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dados Adicionais"
      Height          =   1320
      Left            =   8130
      TabIndex        =   66
      Top             =   6390
      Width           =   6900
      Begin VB.TextBox Text1 
         Height          =   900
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   67
         Text            =   "Entrada.frx":0000
         Top             =   270
         Width           =   6630
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cálculo do ISSQN"
      Height          =   1035
      Left            =   8130
      TabIndex        =   57
      Top             =   5370
      Width           =   6900
      Begin rdActiveText.ActiveText ActiveText15 
         Height          =   315
         Left            =   1875
         TabIndex        =   58
         Top             =   240
         Width           =   1545
         _ExtentX        =   2725
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
      Begin rdActiveText.ActiveText ActiveText16 
         Height          =   315
         Left            =   5340
         TabIndex        =   60
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
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
      Begin rdActiveText.ActiveText ActiveText17 
         Height          =   315
         Left            =   1875
         TabIndex        =   62
         Top             =   615
         Width           =   1545
         _ExtentX        =   2725
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
      Begin rdActiveText.ActiveText ActiveText18 
         Height          =   315
         Left            =   5340
         TabIndex        =   64
         Top             =   615
         Width           =   1365
         _ExtentX        =   2408
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
      Begin VB.Label Label17 
         Caption         =   "Valordo ISSQN:"
         Height          =   300
         Left            =   3555
         TabIndex        =   65
         Top             =   645
         Width           =   2175
      End
      Begin VB.Label Label16 
         Caption         =   "Base Cálculo ISSQN:"
         Height          =   300
         Left            =   120
         TabIndex        =   63
         Top             =   645
         Width           =   2175
      End
      Begin VB.Label Label15 
         Caption         =   "Valor total dos Serviços:"
         Height          =   300
         Left            =   3540
         TabIndex        =   61
         Top             =   270
         Width           =   2175
      End
      Begin VB.Label Label14 
         Caption         =   "Inscrição Mun.:"
         Height          =   285
         Left            =   120
         TabIndex        =   59
         Top             =   270
         Width           =   1110
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produtos / Serviços"
      Height          =   4680
      Left            =   8115
      TabIndex        =   52
      Top             =   690
      Width           =   6915
      Begin VB.CommandButton Command5 
         Caption         =   "="
         Height          =   315
         Left            =   6450
         TabIndex        =   56
         ToolTipText     =   "Remover Fornecedor."
         Top             =   600
         Width           =   345
      End
      Begin VB.CommandButton Command4 
         Caption         =   ">>"
         Height          =   315
         Left            =   6450
         TabIndex        =   55
         ToolTipText     =   "Remover Fornecedor."
         Top             =   945
         Width           =   345
      End
      Begin VB.CommandButton Command3 
         Caption         =   "<<"
         Height          =   315
         Left            =   6450
         TabIndex        =   54
         ToolTipText     =   "Adicionar Fornecedor."
         Top             =   270
         Width           =   345
      End
      Begin MSComctlLib.ListView lstProdutos 
         Height          =   4290
         Left            =   120
         TabIndex        =   53
         Top             =   270
         Width           =   6285
         _ExtentX        =   11086
         _ExtentY        =   7567
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   15
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "NCM/SH"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "CSO SN"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "CFOP"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Unid."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Quant."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Valor Unitário"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Valor Total"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Valor Desc."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "B. Cálc. ICMS"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Valor ICMS"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Valor IPI"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Alíq. ICMS"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Alíq. IPI"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame fraTransportadora 
      Caption         =   "Tranportadora"
      Height          =   1035
      Left            =   60
      TabIndex        =   43
      Top             =   6660
      Width           =   7965
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1350
         TabIndex        =   49
         Text            =   "Combo1"
         Top             =   615
         Width           =   1830
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   255
         Left            =   2535
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   300
         Width           =   270
      End
      Begin rdActiveText.ActiveText ActiveText12 
         Height          =   315
         Left            =   2880
         TabIndex        =   44
         Top             =   240
         Width           =   4950
         _ExtentX        =   8731
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
      Begin rdActiveText.ActiveText ActiveText13 
         Height          =   315
         Left            =   1350
         TabIndex        =   46
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
      Begin rdActiveText.ActiveText ActiveText14 
         Height          =   315
         Left            =   7020
         TabIndex        =   51
         Top             =   615
         Width           =   810
         _ExtentX        =   1429
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
         MaxLength       =   7
         Text            =   "AAA9999"
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin VB.Label Label13 
         Caption         =   "Placa do Veículo:"
         Height          =   270
         Left            =   5460
         TabIndex        =   50
         Top             =   645
         Width           =   1530
      End
      Begin VB.Label Label12 
         Caption         =   "Frete por Conta: "
         Height          =   270
         Left            =   120
         TabIndex        =   48
         Top             =   645
         Width           =   1245
      End
      Begin VB.Label Label11 
         Caption         =   "Código:"
         Height          =   285
         Left            =   120
         TabIndex        =   47
         Top             =   270
         Width           =   1110
      End
   End
   Begin VB.Frame fraImpostos 
      Caption         =   "Impostos"
      Height          =   2535
      Left            =   60
      TabIndex        =   20
      Top             =   4125
      Width           =   7965
      Begin rdActiveText.ActiveText vlrICMS 
         Height          =   315
         Left            =   2565
         TabIndex        =   21
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
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
      Begin rdActiveText.ActiveText ActiveText2 
         Height          =   315
         Left            =   6450
         TabIndex        =   23
         Top             =   255
         Width           =   1365
         _ExtentX        =   2408
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
      Begin rdActiveText.ActiveText ActiveText3 
         Height          =   315
         Left            =   2565
         TabIndex        =   25
         Top             =   615
         Width           =   1365
         _ExtentX        =   2408
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
      Begin rdActiveText.ActiveText ActiveText4 
         Height          =   315
         Left            =   6450
         TabIndex        =   27
         Top             =   630
         Width           =   1365
         _ExtentX        =   2408
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
      Begin rdActiveText.ActiveText ActiveText5 
         Height          =   315
         Left            =   2565
         TabIndex        =   29
         Top             =   990
         Width           =   1365
         _ExtentX        =   2408
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
      Begin rdActiveText.ActiveText ActiveText6 
         Height          =   315
         Left            =   6450
         TabIndex        =   31
         Top             =   990
         Width           =   1365
         _ExtentX        =   2408
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
      Begin rdActiveText.ActiveText ActiveText7 
         Height          =   315
         Left            =   2565
         TabIndex        =   33
         Top             =   1365
         Width           =   1365
         _ExtentX        =   2408
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
      Begin rdActiveText.ActiveText ActiveText8 
         Height          =   315
         Left            =   6450
         TabIndex        =   35
         Top             =   1365
         Width           =   1365
         _ExtentX        =   2408
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
      Begin rdActiveText.ActiveText ActiveText9 
         Height          =   315
         Left            =   2565
         TabIndex        =   37
         Top             =   1740
         Width           =   1365
         _ExtentX        =   2408
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
      Begin rdActiveText.ActiveText ActiveText10 
         Height          =   315
         Left            =   6450
         TabIndex        =   39
         Top             =   1740
         Width           =   1365
         _ExtentX        =   2408
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
      Begin rdActiveText.ActiveText ActiveText11 
         Height          =   315
         Left            =   2565
         TabIndex        =   41
         Top             =   2115
         Width           =   1365
         _ExtentX        =   2408
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
      Begin VB.Label Label10 
         Caption         =   "Valor Total da Nota:"
         Height          =   300
         Left            =   120
         TabIndex        =   42
         Top             =   2145
         Width           =   2175
      End
      Begin VB.Label Label9 
         Caption         =   "Valor do IPI:"
         Height          =   300
         Left            =   4260
         TabIndex        =   40
         Top             =   1770
         Width           =   2055
      End
      Begin VB.Label Label8 
         Caption         =   "Outras Despesas:"
         Height          =   300
         Left            =   120
         TabIndex        =   38
         Top             =   1770
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "Valor do Desconto:"
         Height          =   300
         Left            =   4260
         TabIndex        =   36
         Top             =   1395
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Valor do Seguro:"
         Height          =   300
         Left            =   120
         TabIndex        =   34
         Top             =   1395
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Valor do Frete:"
         Height          =   300
         Left            =   4260
         TabIndex        =   32
         Top             =   1020
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Valor Total dos Produtos:"
         Height          =   300
         Left            =   120
         TabIndex        =   30
         Top             =   1020
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Valor do ICMS Substituto:"
         Height          =   300
         Left            =   4260
         TabIndex        =   28
         Top             =   660
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Base de Calculo do ICMS ST:"
         Height          =   300
         Left            =   120
         TabIndex        =   26
         Top             =   645
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Valor do ICMS:"
         Height          =   300
         Left            =   4260
         TabIndex        =   24
         Top             =   285
         Width           =   1500
      End
      Begin VB.Label lblICMS 
         Caption         =   "Base de Calculo do ICMS:"
         Height          =   300
         Left            =   120
         TabIndex        =   22
         Top             =   270
         Width           =   2175
      End
   End
   Begin VB.Frame fraFaturasDuplicatas 
      Caption         =   "Faturas/Duplicatas"
      Height          =   2400
      Left            =   60
      TabIndex        =   13
      Top             =   1725
      Width           =   7980
      Begin VB.CommandButton cmdModificar 
         Caption         =   "="
         Height          =   315
         Left            =   7500
         TabIndex        =   19
         ToolTipText     =   "Remover Fornecedor."
         Top             =   960
         Width           =   345
      End
      Begin VB.CommandButton cmdRemover 
         Caption         =   ">>"
         Height          =   315
         Left            =   7485
         TabIndex        =   18
         ToolTipText     =   "Remover Fornecedor."
         Top             =   1305
         Width           =   345
      End
      Begin VB.CommandButton cmdAdicionar 
         Caption         =   "<<"
         Height          =   315
         Left            =   7500
         TabIndex        =   17
         ToolTipText     =   "Adicionar Fornecedor."
         Top             =   615
         Width           =   345
      End
      Begin MSComctlLib.ListView lstFaturasDuplicatas 
         Height          =   1665
         Left            =   120
         TabIndex        =   16
         Top             =   615
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   2937
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fatura / Duplicata"
            Object.Width           =   2716
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Vencimento"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Valor"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.ComboBox cmbDefinicao 
         Height          =   315
         Left            =   1230
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   240
         Width           =   2625
      End
      Begin VB.Label lblDefinicao 
         Caption         =   "Definição:"
         Height          =   270
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   1140
      End
   End
   Begin VB.Frame fraFornecedor 
      Caption         =   "Fornecedor"
      Height          =   1050
      Left            =   60
      TabIndex        =   4
      Top             =   675
      Width           =   7980
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   300
         Width           =   270
      End
      Begin rdActiveText.ActiveText vlrCPFCNPJ 
         Height          =   315
         Left            =   1215
         TabIndex        =   5
         Top             =   615
         Width           =   2625
         _ExtentX        =   4630
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
         MaxLength       =   14
         TextMask        =   7
         RawText         =   7
         Mask            =   "###.###.###-##"
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
         Locked          =   -1  'True
      End
      Begin rdActiveText.ActiveText txtNome 
         Height          =   315
         Left            =   2730
         TabIndex        =   6
         Top             =   240
         Width           =   5100
         _ExtentX        =   8996
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
      Begin rdActiveText.ActiveText vlrIE 
         Height          =   315
         Left            =   5115
         TabIndex        =   7
         Top             =   615
         Width           =   2700
         _ExtentX        =   4763
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
      Begin rdActiveText.ActiveText ActiveText1 
         Height          =   315
         Left            =   1200
         TabIndex        =   12
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
      Begin VB.Label lblNome 
         Caption         =   "Razão Social:"
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   270
         Width           =   1110
      End
      Begin VB.Label lblCNPJ 
         Caption         =   "CNPJ:"
         Height          =   225
         Left            =   120
         TabIndex        =   9
         Top             =   645
         Width           =   1125
      End
      Begin VB.Label lblIE 
         Caption         =   "Ins. Estadual:"
         Height          =   225
         Left            =   3960
         TabIndex        =   8
         Top             =   645
         Width           =   1080
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
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   300
         Width           =   270
      End
      Begin rdActiveText.ActiveText vlrCod 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
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
         TabIndex        =   3
         Top             =   270
         Width           =   1065
      End
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   15105
      Top             =   3030
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
            Picture         =   "Entrada.frx":0006
            Key             =   "Inserido"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Entrada.frx":0160
            Key             =   "Gravado"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Entrada.frx":02BA
            Key             =   "Alterado"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Entrada.frx":0414
            Key             =   "Removido"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstLegenda 
      Height          =   1065
      Left            =   15105
      TabIndex        =   73
      Top             =   6630
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
      Left            =   15105
      TabIndex        =   80
      Top             =   6405
      Width           =   1215
   End
End
Attribute VB_Name = "frmEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFechar_Click()
   Unload Me
End Sub
