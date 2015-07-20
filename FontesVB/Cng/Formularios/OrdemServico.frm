VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "InfinityControl.ocx"
Begin VB.Form frmOrdemServico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Odem de Serviço"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Lançamentos Adicionais"
      Height          =   1065
      Left            =   60
      TabIndex        =   36
      Top             =   4530
      Width           =   7770
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informações de Cadastro"
      Height          =   1065
      Left            =   60
      TabIndex        =   35
      Top             =   3465
      Width           =   7770
   End
   Begin VB.Frame fraCliente 
      Caption         =   "Cliente"
      Height          =   1755
      Left            =   60
      TabIndex        =   19
      Top             =   1710
      Width           =   7785
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   255
         Left            =   2595
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   255
         Width           =   270
      End
      Begin rdActiveText.ActiveText ActiveText1 
         Height          =   315
         Left            =   1395
         TabIndex        =   21
         Top             =   210
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
      Begin rdActiveText.ActiveText txtNome 
         Height          =   315
         Left            =   2940
         TabIndex        =   23
         Top             =   210
         Width           =   4695
         _ExtentX        =   8281
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
         Left            =   1395
         TabIndex        =   24
         Top             =   1335
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
         Left            =   2580
         TabIndex        =   25
         Top             =   960
         Width           =   5055
         _ExtentX        =   8916
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
         Left            =   1395
         TabIndex        =   26
         Top             =   960
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
         Left            =   6615
         TabIndex        =   27
         Top             =   585
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
         Left            =   1395
         TabIndex        =   28
         Top             =   585
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
      Begin rdActiveText.ActiveText ActiveText2 
         Height          =   315
         Left            =   6615
         TabIndex        =   34
         Top             =   1335
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
      Begin VB.Label lblEndereco 
         Caption         =   "Endereço:"
         Height          =   210
         Left            =   120
         TabIndex        =   33
         Top             =   615
         Width           =   1155
      End
      Begin VB.Label lblBairro 
         Caption         =   "Bairro:"
         Height          =   210
         Left            =   120
         TabIndex        =   32
         Top             =   1365
         Width           =   1140
      End
      Begin VB.Label lblNumero 
         Caption         =   "Número:"
         Height          =   270
         Left            =   5835
         TabIndex        =   31
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label lblCidade 
         Caption         =   "Cidade:"
         Height          =   300
         Left            =   120
         TabIndex        =   30
         Top             =   990
         Width           =   1170
      End
      Begin VB.Label lblUF 
         Caption         =   "Estado:"
         Height          =   225
         Left            =   5835
         TabIndex        =   29
         Top             =   1365
         Width           =   990
      End
      Begin VB.Label Label3 
         Caption         =   "Código:"
         Height          =   270
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Frame fraParametros 
      Caption         =   "Parâmetros"
      Height          =   1050
      Left            =   60
      TabIndex        =   10
      Top             =   660
      Width           =   7785
      Begin VB.CommandButton cmdConsGrupo 
         Caption         =   "..."
         Height          =   255
         Left            =   2535
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   675
         Width           =   270
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   6135
         TabIndex        =   14
         Text            =   "Combo1"
         Top             =   240
         Width           =   1530
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1395
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   240
         Width           =   2055
      End
      Begin rdActiveText.ActiveText txtDescGrupo 
         Height          =   315
         Left            =   2895
         TabIndex        =   15
         Top             =   615
         Width           =   4005
         _ExtentX        =   7064
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
      Begin rdActiveText.ActiveText vlrCodGrupo 
         Height          =   315
         Left            =   1395
         TabIndex        =   17
         Top             =   615
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
      Begin VB.Label lblGrupo 
         Caption         =   "Operador:"
         Height          =   270
         Left            =   120
         TabIndex        =   18
         Top             =   645
         Width           =   1050
      End
      Begin VB.Label Label2 
         Caption         =   "Situação:"
         Height          =   255
         Left            =   5280
         TabIndex        =   13
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Ordem:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   270
         Width           =   1545
      End
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Novo"
      Enabled         =   0   'False
      Height          =   345
      Left            =   7935
      TabIndex        =   9
      Top             =   90
      Width           =   1320
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "&Consultar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   7935
      TabIndex        =   8
      Top             =   465
      Width           =   1320
   End
   Begin VB.CommandButton cmdSalvar 
      Cancel          =   -1  'True
      Caption         =   "&Salvar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   7935
      TabIndex        =   7
      Top             =   840
      Width           =   1320
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Enabled         =   0   'False
      Height          =   345
      Left            =   7935
      TabIndex        =   6
      Tag             =   "6"
      Top             =   1215
      Width           =   1320
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "&Limpar"
      Height          =   345
      Left            =   7935
      TabIndex        =   5
      Top             =   1590
      Width           =   1320
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   345
      Left            =   7935
      TabIndex        =   4
      Top             =   1965
      Width           =   1320
   End
   Begin VB.Frame fraIdentificacao 
      Caption         =   "Identificação"
      Height          =   660
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   7785
      Begin VB.CommandButton cmdPesqForn 
         Caption         =   "..."
         Height          =   255
         Left            =   2385
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   315
         Width           =   270
      End
      Begin rdActiveText.ActiveText vlrCod 
         Height          =   315
         Left            =   1185
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
End
Attribute VB_Name = "frmOrdemServico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
