VERSION 5.00
Begin VB.Form frmConnectMail 
   BorderStyle     =   0  'None
   Caption         =   "Conexão de E-mail"
   ClientHeight    =   1065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1065
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblStatus 
      Caption         =   "Label1"
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
      Left            =   1260
      TabIndex        =   0
      Top             =   360
      Width           =   4320
   End
End
Attribute VB_Name = "frmConnectMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private POPServerPorta As Long
Private POPServer As String
Private POPAuthenticateSSL As Boolean
Private UserName As String
Private Password As String

Public Property Let Porta(ByVal vNewValue As Long)
   POPServerPorta = vNewValue
End Property

Public Property Let POP3(ByVal vNewValue As String)
   POPServer = vNewValue
End Property

Public Property Let SSL(ByVal vNewValue As Boolean)
   POPAuthenticateSSL = vNewValue
End Property

Public Property Let Usuario(ByVal vNewValue As String)
   UserName = vNewValue
End Property

Public Property Let Senha(ByVal vNewValue As String)
   Password = vNewValue
End Property

Private Sub Form_Load()
   f1.FormCentralizar Me
End Sub

Public Function Conectar(objErro As Object, objGetMail As Object) As Boolean
   On Error GoTo Conectar_E
   
   Conectar = False
   Me.lblStatus.Caption = "Conectando com o provedor de e-mail ..."
   With objGetMail
      DoEvents
   
      'Estabelece a conexão com o provedor de e-mail
      If Not .OpenPOP3(POPServer, POPServerPorta, UserName, Password, True) Then Exit Function
      
      DoEvents
      Me.lblStatus.Caption = "Conectado ..."
   End With
   
   Conectar = True
   Unload Me
   
   Exit Function
   
Conectar_E:
   objErro.Salvar Err
   objErro.ModRotina = "Conectar"
   Unload Me
End Function
