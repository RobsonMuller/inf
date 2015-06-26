VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "InfinityControl.ocx"
Begin VB.Form frmBarCode 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Leitor de Código de Barras ..."
   ClientHeight    =   900
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5700
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   345
      Left            =   4320
      TabIndex        =   13
      Top             =   495
      Width           =   1320
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "&Limpar"
      Height          =   345
      Left            =   2955
      TabIndex        =   12
      Top             =   495
      Width           =   1320
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "&Confirmar"
      Height          =   345
      Left            =   1590
      TabIndex        =   11
      Top             =   495
      Width           =   1320
   End
   Begin rdActiveText.ActiveText txtCode 
      Height          =   315
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   4
      RawText         =   0
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCode 
      Height          =   315
      Index           =   1
      Left            =   570
      TabIndex        =   1
      Top             =   90
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   4
      RawText         =   0
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCode 
      Height          =   315
      Index           =   2
      Left            =   1080
      TabIndex        =   2
      Top             =   90
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   4
      RawText         =   0
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCode 
      Height          =   315
      Index           =   3
      Left            =   1590
      TabIndex        =   3
      Top             =   90
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   4
      RawText         =   0
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCode 
      Height          =   315
      Index           =   4
      Left            =   2100
      TabIndex        =   4
      Top             =   90
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   4
      RawText         =   0
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCode 
      Height          =   315
      Index           =   5
      Left            =   2610
      TabIndex        =   5
      Top             =   90
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   4
      RawText         =   0
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCode 
      Height          =   315
      Index           =   6
      Left            =   3120
      TabIndex        =   6
      Top             =   90
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   4
      RawText         =   0
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCode 
      Height          =   315
      Index           =   7
      Left            =   3630
      TabIndex        =   7
      Top             =   90
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   4
      RawText         =   0
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCode 
      Height          =   315
      Index           =   8
      Left            =   4140
      TabIndex        =   8
      Top             =   90
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   4
      RawText         =   0
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCode 
      Height          =   315
      Index           =   9
      Left            =   4650
      TabIndex        =   9
      Top             =   90
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   4
      RawText         =   0
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCode 
      Height          =   315
      Index           =   10
      Left            =   5160
      TabIndex        =   10
      Top             =   105
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   4
      RawText         =   0
      FontSize        =   8,25
   End
End
Attribute VB_Name = "frmBarCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private blnCancelado As Boolean
Private txtBarCode As String

Private Sub cmdConfirmar_Click()
   Dim intCount As Integer
   
   txtBarCode = ""
   
   For intCount = 0 To 10
      If Me.txtCode(intCount) = "" Then
         mMsgInfo "Todos os campos devem ser preenchidos! Verifique."
         txtBarCode = ""
         mFocus Me.txtCode(intCount)
         Exit Sub
      End If
      txtBarCode = txtBarCode & Me.txtCode(intCount)
   Next
   
   blnCancelado = False
   Unload Me
End Sub

Public Property Get BarCode() As String
   BarCode = txtBarCode
End Property

Public Property Get Cancelado() As Boolean
   Cancelado = blnCancelado
End Property

Private Sub cmdFechar_Click()
   blnCancelado = True
   Unload Me
End Sub

Private Sub cmdLimpar_Click()
   f1.Limpar Me
   
   mFocus Me.txtCode(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Me.Hide
End Sub

Private Sub txtCode_Change(Index As Integer)
   If Len(Me.txtCode(Index)) = 4 Then
      If Index = 10 Then
         mFocus Me.cmdConfirmar
      Else
         mFocus Me.txtCode(Index + 1)
      End If
   End If
End Sub
