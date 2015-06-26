Private Enum POP3States
    POP3_Connect
    POP3_USER
    POP3_PASS
    POP3_STAT
    POP3_RETR
    POP3_DELE
    POP3_QUIT
End Enum

Private m_State As POP3States

Private Sub cmd_conect_Click()
    Winsock1.Close
    Winsock1.LocalPort = 0
    'Winsock1.Connect "125.125.125.2", 9456
    Winsock1.Connect "pop.clix.pt", 110
    m_State = POP3_Connect
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Winsock1.SendData "Quit" & vbCrLf
End Sub

Private Sub List1_Click()
    Text1.Text = List1.List(List1.ListIndex)
    Extrair
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim x As Integer
    Dim strdata As String
    Dim pos, t, numMsg As Integer
    
    Winsock1.GetData strdata
    List1.AddItem strdata
    
                
    Text4 = Text4 + strdata
                
    'If Left$(strdata, 1) = "+" Then
        Select Case m_State
        Case POP3_Connect
            m_State = POP3_USER
            Winsock1.SendData "USER x0110339" & vbCrLf
            List1.AddItem strdata
        Case POP3_USER
            m_State = POP3_PASS
            Winsock1.SendData "PASS " & txt_pass & vbCrLf
            List1.AddItem strdata
        Case POP3_PASS
            m_State = POP3_STAT
            Winsock1.SendData "STAT" & vbCrLf
            List1.AddItem strdata
        Case POP3_STAT
            k = k + 1
            If k = 1 Then
                pos = InStr(5, strdata, " ", 1) 'Procura espaço
                t = pos - 5 'Determina tamanho da string, com o nº de msg
                numMsg = Mid$(strdata, 5, t) 'Numero de mensagens
            End If
            
            If numMsg > 0 Then MsgBox "Tem " & numMsg & " novas mensagens"
                
                
            For x = 1 To numMsg
                Winsock1.SendData "RETR " & x & vbCrLf
                DoEvents
                List1.AddItem strdata
                
            Next x
            
        
       ' Case POP3_R
       '         k = k + 1
       '         MsgBox k
       '         If k <= numMsg Then
       '             Winsock1.SendData "RETR " & k & vbCrLf
       '             DoEvents
       '         End If
       '         List1.AddItem strdata
       '         m_State = POP3_R
        Case Else
            
        End Select
            
    'End If
    Text3 = k
    
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If Number = 11001 Then
        MsgBox "Host não encontrado", , Number
    Else
        MsgBox Description, , Number
    End If
End Sub

Public Sub Extrair()
Dim p1, p2, p3, p4 As Integer
Dim c1, c2 As Integer
Dim k1, k2, k3 As Integer
Dim s1, s2, s3 As String
Dim nome, de, para, assunto, data As String


'From
p1 = InStr(1, Text1, "From:")

c1 = Len(Text1)
k1 = c1 - p1
s1 = Right(Text1, k1 + 1)
p2 = InStr(1, s1, "<")

nome = Mid(s1, 7, p2 - 7)
txt_name = nome

p3 = InStr(1, s1, ">")
de = Mid(s1, p2 + 1, p3 - p2 - 1)
txt_from = de


'To
p1 = InStr(1, s1, "To:")
p2 = InStr(1, s1, "Subject:")
para = Mid(s1, p1 + 5, p2 - p1 - 8)
txt_to = para

'Subject
p1 = InStr(1, s1, "Subject:")
p2 = InStr(1, s1, "Date:")
assunto = Mid(s1, p1 + 9, p2 - p1 - 11)
txt_subject = assunto

'Date
p1 = InStr(1, s1, "Date:")
data = Mid(s1, p1 + 6, 22)
txt_data = data

End Sub
