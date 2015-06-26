VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMapa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Localização ..."
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdZoomIn 
      Height          =   450
      Left            =   10350
      Picture         =   "Mapa.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   555
      Width           =   480
   End
   Begin VB.CommandButton cmdZoomOut 
      Height          =   450
      Left            =   10350
      Picture         =   "Mapa.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      Width           =   480
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   8250
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   10230
      ExtentX         =   18045
      ExtentY         =   14552
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intZoom As Integer
Private strEnd As String

Public Property Let Endereco(ByVal vNewValue As String)
   strEnd = vNewValue
End Property

Public Sub LocalizarEndereco()
   intZoom = 15
   WebBrowser1.Navigate ("http://maps.googleapis.com/maps/api/staticmap?center=" & strEnd & _
      "&zoom=" & intZoom & "&size=1024x512&maptype=roadmap&markers=size:mid%7Ccolor:red%7C" & strEnd & "&sensor=true")
End Sub

Private Sub Command1_Click()
   
End Sub

Private Sub cmdZoomIn_Click()
   intZoom = intZoom + 1
   WebBrowser1.Navigate ("about:blank")
   WebBrowser1.Navigate ("http://maps.googleapis.com/maps/api/staticmap?center=" & strEnd & _
      "&zoom=" & intZoom & "&size=1024x512&maptype=roadmap&markers=size:mid%7Ccolor:red%7C" & strEnd & "&sensor=true")
End Sub

Private Sub cmdZoomOut_Click()
   intZoom = intZoom - 1
   WebBrowser1.Navigate ("about:blank")
   WebBrowser1.Navigate ("http://maps.googleapis.com/maps/api/staticmap?center=" & strEnd & _
      "&zoom=" & intZoom & "&size=1024x512&maptype=roadmap&markers=size:mid%7Ccolor:red%7C" & strEnd & "&sensor=true")
End Sub

