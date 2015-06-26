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
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
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
