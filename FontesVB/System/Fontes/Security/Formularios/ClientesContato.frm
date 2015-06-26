VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "InfinityControl.ocx"
Begin VB.Form frmClientesContato 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Contatos"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   5595
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   345
      Left            =   2835
      TabIndex        =   4
      Top             =   1500
      Width           =   1320
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   4215
      TabIndex        =   5
      Top             =   1500
      Width           =   1320
   End
   Begin VB.Frame fraParametros 
      Caption         =   "Parâmetros"
      Height          =   1425
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   5475
      Begin rdActiveText.ActiveText txtNome 
         Height          =   315
         Left            =   1050
         TabIndex        =   1
         Top             =   240
         Width           =   4275
         _ExtentX        =   7541
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
         Left            =   1050
         TabIndex        =   2
         Top             =   615
         Width           =   1275
         _ExtentX        =   2249
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
      Begin rdActiveText.ActiveText txtEmail 
         Height          =   315
         Left            =   1050
         TabIndex        =   3
         Top             =   990
         Width           =   4290
         _ExtentX        =   7567
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
      Begin VB.Label lblEmail 
         Caption         =   "E-mail:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblTelefone 
         Caption         =   "Telefone:"
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   645
         Width           =   1365
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
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   270
         Width           =   1320
      End
   End
End
Attribute VB_Name = "frmClientesContato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private blnCancelado As Boolean

Private Sub cmdCancelar_Click()
   blnCancelado = True
   Me.Hide
End Sub

Private Sub cmdOk_Click()
   If Len(Trim(Me.txtNome)) = 0 Then
      mMsgInfo "O nome do contato deve ser informado! Verifique."
      mFocus Me.txtNome
      Exit Sub
   End If
   
   Me.Hide
End Sub

Public Property Get Cancelado() As Boolean
   Cancelado = blnCancelado
End Property

Public Property Get Contato() As String
   Contato = Me.txtNome
End Property

Public Property Let Contato(ByVal vNewValue As String)
   Me.txtNome = vNewValue
End Property

Public Property Get Telefone() As String
   Telefone = Me.txtTelefone
End Property

Public Property Let Telefone(ByVal vNewValue As String)
   Me.txtTelefone = vNewValue
End Property

Public Property Get Email() As String
   Email = Me.txtEmail
End Property

Public Property Let Email(ByVal vNewValue As String)
   Me.txtEmail = vNewValue
End Property
