VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMessage 
   Caption         =   "Message"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4965
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   4965
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkAttachments 
      Caption         =   "include attachments"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   4320
      Width           =   1695
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   7435
      _Version        =   393217
      Indentation     =   529
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
   hWndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type
  
Dim m As OSPOP3_Plus.Message
      
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
  Dim h As OSPOP3_Plus.Header
  Dim s
  Dim nd As Node
  Dim e As OSPOP3_Plus.Email
  Dim a As OSPOP3_Plus.Attachment
  Dim iCount As Integer
  Dim i As Integer
  Dim j As Integer

On Error Resume Next
  tvw.Nodes.Clear
  i = frmSession.lvwMessages.SelectedItem.Text
  
  If frmSession.bHeadersOnly Then
    Set m = frmSession.oSession.GetMessageHeaders(i)
  Else
    Set m = frmSession.oSession.GetMessage(i)
  End If
  tvw.Nodes.Add , , "m" & i, "message " & i
  
  tvw.Nodes.Add "m" & i, tvwChild, "m" & i & "from", "from"
  tvw.Nodes.Add "m" & i & "from", tvwChild, , "name: " & m.Sender.Name
  tvw.Nodes.Add "m" & i & "from", tvwChild, , "address: " & m.Sender.Address
  
  tvw.Nodes.Add "m" & i, tvwChild, "m" & i & "to", "to"
  For Each e In m.Recipients
    Set nd = tvw.Nodes.Add("m" & i & "to", tvwChild, , "recipient")
    tvw.Nodes.Add nd.Index, tvwChild, , "name: " & e.Name
    tvw.Nodes.Add nd.Index, tvwChild, , "address: " & e.Address
  Next
  
  tvw.Nodes.Add "m" & i, tvwChild, "m" & i & "cc", "cc"
  For Each e In m.RecipientsCC
    Set nd = tvw.Nodes.Add("m" & i & "cc", tvwChild, , "recipient")
    tvw.Nodes.Add nd.Index, tvwChild, , "name: " & e.Name
    tvw.Nodes.Add nd.Index, tvwChild, , "address: " & e.Address
  Next
  
  tvw.Nodes.Add "m" & i, tvwChild, "m" & i & "subject", "subject"
  tvw.Nodes.Add "m" & i & "subject", tvwChild, , m.Subject
  
  tvw.Nodes.Add "m" & i, tvwChild, "m" & i & "date", "date"
  tvw.Nodes.Add "m" & i & "date", tvwChild, , m.DateSent
  
  tvw.Nodes.Add "m" & i, tvwChild, "m" & i & "ContentType", "Content-Type"
  tvw.Nodes.Add "m" & i & "ContentType", tvwChild, , m.ContentType
  
  tvw.Nodes.Add "m" & i, tvwChild, "m" & i & "ContentTransferEncoding", "Content-Transfer-Encoding"
  tvw.Nodes.Add "m" & i & "ContentTransferEncoding", tvwChild, , m.ContentTransferEncoding
  
  tvw.Nodes.Add "m" & i, tvwChild, "m" & i & "char", "Charset"
  tvw.Nodes.Add "m" & i & "char", tvwChild, , m.Charset
  
  tvw.Nodes.Add "m" & i, tvwChild, "m" & i & "UIDL", "UIDL"
  tvw.Nodes.Add "m" & i & "UIDL", tvwChild, , m.UIDL
  
  tvw.Nodes.Add "m" & i, tvwChild, "m" & i & "h", "headers"
  For Each h In m.Headers
    tvw.Nodes.Add "m" & i & "h", tvwChild, "m" & i & h.Name, h.Name
    For Each s In h.Values
      tvw.Nodes.Add "m" & i & h.Name, tvwChild, , s
    Next
  Next
  
  If Not frmSession.bHeadersOnly Then
    tvw.Nodes.Add "m" & i, tvwChild, "m" & i & "body", "body"
    tvw.Nodes.Add "m" & i & "body", tvwChild, , Left(m.Body, 1000)
    
    tvw.Nodes.Add "m" & i, tvwChild, "m" & i & "html", "html"
    tvw.Nodes.Add "m" & i & "html", tvwChild, , Left(m.HTMLBody, 1000)
    
    tvw.Nodes.Add "m" & i, tvwChild, "m" & i & "a", "attachments"
    For Each a In m.Attachments
      j = j + 1
      tvw.Nodes.Add "m" & i & "a", tvwChild, "m" & i & j, a.AttachmentName
      tvw.Nodes.Add "m" & i & j, tvwChild, , "ContentDisposition: " & a.ContentDisposition
      tvw.Nodes.Add "m" & i & j, tvwChild, , "ContentTransferEncoding: " & a.ContentTransferEncoding
      tvw.Nodes.Add "m" & i & j, tvwChild, , "ContentType: " & a.ContentType
      tvw.Nodes.Add "m" & i & j, tvwChild, , "FileName: " & a.fileName
      tvw.Nodes.Add "m" & i & j, tvwChild, , "Body: " & Left(a.Body, 1000)
      'a.Save App.Path & "\" & a.FileName 'uncomment to save an attachment
    Next
    
    'm.Save App.Path & "\" & m.UIDL & ".eml" 'uncomment to save the message
  End If
  
  For Each nd In tvw.Nodes
    If Right(nd.Key, 1) <> "h" Then nd.Expanded = True 'do not expand headers
  Next
  tvw.Nodes(1).Selected = True
End Sub

Private Sub btnSave_Click()
    Dim path As String
    path = SelectPath
    m.Save (path + "\" + m.UIDL + ".eml")
    If (chkAttachments.Value = 1) Then
        Dim a As OSPOP3_Plus.Attachment
        For Each a In m.Attachments
            Dim fileName As String
            Dim attachmentCounter As Integer
            attachmentCounter = 1
            If a.AttachmentName <> "" Then
                fileName = Replace(a.AttachmentName, Chr(34), "") 'attachment name may contain quotes!
            Else
                fileName = "attachment_" + CStr(attachmentCounter) 'file name can't be blank
                If a.ContentType = "message/rfc822" Then
                    fileName = fileName + ".eml"
                End If
            End If

            a.Save (path + "\" + fileName)
            attachmentCounter = attachmentCounter + 1
        Next
    End If
    MsgBox "message saved"
End Sub

Private Function SelectPath() As String
  'Opens a Treeview control that displays the directories in a computer

     Dim lpIDList As Long
     Dim sBuffer As String
     Dim szTitle As String
     Dim tBrowseInfo As BrowseInfo

     szTitle = "This is the title"
     With tBrowseInfo
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
     End With

     lpIDList = SHBrowseForFolder(tBrowseInfo)

     If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        SelectPath = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
     End If
  End Function
