VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSession 
   Caption         =   "OstroSoft POP3 Component"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   440
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   580
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtStatus 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   19
      Text            =   "frmSession.frx":0000
      Top             =   4200
      Width           =   8655
   End
   Begin VB.Frame fmSession 
      Caption         =   "Session"
      Height          =   1260
      Left            =   0
      TabIndex        =   13
      Top             =   60
      Width           =   8655
      Begin VB.CheckBox chkSSL 
         Caption         =   "Use SSL"
         Height          =   255
         Left            =   1440
         TabIndex        =   22
         Top             =   810
         Width           =   975
      End
      Begin VB.TextBox txtPort 
         Height          =   315
         Left            =   3600
         TabIndex        =   20
         Text            =   "110"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtServer 
         Height          =   315
         Left            =   660
         TabIndex        =   0
         Text            =   "localhost"
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtLogin 
         Height          =   315
         Left            =   5040
         TabIndex        =   1
         Text            =   "test@localhost"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7200
         PasswordChar    =   "*"
         TabIndex        =   2
         Text            =   "test"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Default         =   -1  'True
         Height          =   315
         Left            =   3600
         TabIndex        =   3
         Top             =   780
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Port"
         Height          =   195
         Left            =   3200
         TabIndex        =   21
         Top             =   420
         Width           =   285
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Server"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   420
         Width           =   465
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Login"
         Height          =   195
         Left            =   4440
         TabIndex        =   15
         Top             =   420
         Width           =   390
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Password"
         Height          =   195
         Left            =   6360
         TabIndex        =   14
         Top             =   420
         Width           =   690
      End
   End
   Begin VB.Frame fmMessages 
      Caption         =   "Messages"
      Height          =   2730
      Left            =   0
      TabIndex        =   12
      Top             =   1350
      Width           =   8655
      Begin VB.TextBox txtNumber 
         Alignment       =   1  'Right Justify
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   4680
         TabIndex        =   5
         Text            =   "0"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtSize 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2160
         TabIndex        =   4
         Text            =   "0"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh Message List"
         Height          =   315
         Left            =   6840
         TabIndex        =   6
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset Session"
         Height          =   315
         Left            =   6840
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete Message"
         Height          =   315
         Left            =   6840
         TabIndex        =   10
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton cmdHeaders 
         Caption         =   "View Headers"
         Height          =   315
         Left            =   6840
         TabIndex        =   9
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "View Message"
         Height          =   315
         Left            =   6840
         TabIndex        =   8
         Top             =   960
         Width           =   1695
      End
      Begin MSComctlLib.ListView lvwMessages 
         Height          =   1815
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3201
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Size"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "UIDL"
            Object.Width           =   4762
         EndProperty
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Number of emails"
         Height          =   195
         Left            =   3360
         TabIndex        =   18
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Mailbox size, octets (bytes)"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   420
         Width           =   1890
      End
   End
End
Attribute VB_Name = "frmSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Visual Basic example for OstroSoft POP3 Component
'written by Igor Ostrovsky (OstroSoft)
'
'For more information about OstroSoft POP3 Component go to
'http://www.ostrosoft.com/ospop3.aspx
'
'Questions, suggestions, comments - email to info@ostrosoft.com
'or submit a form at http://www.ostrosoft.com/contact.asp

Option Explicit

Public WithEvents oSession As OSPOP3_Plus.Session
Attribute oSession.VB_VarHelpID = -1
Public bHeadersOnly As Boolean
Dim bItem As Boolean

Private Sub cmdConnect_Click()
  cmdConnect.Enabled = False
  If cmdConnect.Caption = "Connect" Then
    txtStatus = ""
    With oSession
        .UseSSL = (chkSSL.Value = 1)
        .OpenPOP3 txtServer.Text, Val(txtPort.Text), txtLogin.Text, txtPassword.Text
    End With
  Else
    oSession.ClosePOP3
  End If
End Sub

Private Sub cmdDelete_Click()
  If lvwMessages.SelectedItem Is Nothing Then Exit Sub
  If MsgBox("Would you like to delete a selected email?", vbYesNo, "Confirm") = vbYes Then
    oSession.DeleteMessage lvwMessages.SelectedItem.Text
    lvwMessages.ListItems.Remove lvwMessages.SelectedItem.Index
  End If
End Sub

Private Sub cmdHeaders_Click()
  If lvwMessages.SelectedItem Is Nothing Then Exit Sub
  bHeadersOnly = True
  frmMessage.Show vbModal
End Sub

Private Sub cmdRefresh_Click()
  GetStats
End Sub

Private Sub cmdReset_Click()
  oSession.ResetSession
  GetStats
End Sub

Private Sub cmdView_Click()
  If lvwMessages.SelectedItem Is Nothing Then Exit Sub
  bHeadersOnly = False
  frmMessage.Show vbModal
End Sub

Private Sub Form_Load()
  Set oSession = New OSPOP3_Plus.Session
  fmMessages.Enabled = False
End Sub

Private Sub lvwMessages_DblClick()
  If bItem Then
    If Not lvwMessages.SelectedItem Is Nothing Then
      bHeadersOnly = True
      frmMessage.Show vbModal
    End If
    bItem = False
  End If
End Sub

Private Sub lvwMessages_ItemClick(ByVal Item As MSComctlLib.ListItem)
  bItem = True
End Sub

Private Sub oSession_Closed()
  txtSize = 0
  txtNumber = 0
  lvwMessages.ListItems.Clear
  
  cmdConnect.Caption = "Connect"
  cmdConnect.Enabled = True
  fmMessages.Enabled = False
End Sub

Private Sub oSession_Connected()
  GetStats
  cmdConnect.Caption = "Disconnect"
  cmdConnect.Enabled = True
  fmMessages.Enabled = True
End Sub

Private Sub oSession_ErrorPOP3(ByVal Number As Long, ByVal Description As String)
  'MsgBox "Error " & Number & ": " & Description
    If oSession.State = OSPOP3_Plus.StateConstants_popClosed Then
        cmdConnect.Caption = "Connect"
        fmMessages.Enabled = False
    Else
        cmdConnect.Caption = "Disconnect"
    End If
    cmdConnect.Enabled = True
End Sub

Private Sub oSession_StatusChanged(ByVal Status As String, ByVal StatusType As OSPOP3_Plus.StatusTypeConstants)
  Dim sPrompt As String
  Dim sTemp As String
  Select Case StatusType
  Case OSPOP3_Plus.StatusTypeConstants_stPOP3Request: sPrompt = "< "
  Case OSPOP3_Plus.StatusTypeConstants_stPOP3Response: sPrompt = "> "
  Case OSPOP3_Plus.StatusTypeConstants_stError: sPrompt = "! "
  Case OSPOP3_Plus.StatusTypeConstants_stState: sPrompt = "# "
  Case Else: sPrompt = "? "
  End Select
  
  sTemp = txtStatus & sPrompt & Status & vbCrLf
  If Len(sTemp) > 32000 Then sTemp = Right(sTemp, 32000)
  
  txtStatus = sTemp & vbCrLf
End Sub

Private Sub GetStats()
  oSession.GetMailboxSize
  txtSize = oSession.MailboxSize
  txtNumber = oSession.MessageCount
  
  Dim oMessageList 'As ArrayList
  Dim oMLE As MessageListEntry
  Dim li As ListItem

  lvwMessages.ListItems.Clear
  Set oMessageList = oSession.GetMessageList
  For Each oMLE In oMessageList
    Set li = lvwMessages.ListItems.Add(, , oMLE.ID)
    li.SubItems(1) = oMLE.Size
    li.SubItems(2) = oMLE.UIDL
  Next
End Sub

Private Sub txtStatus_Change()
  txtStatus.SelStart = Len(txtStatus.Text)
End Sub
