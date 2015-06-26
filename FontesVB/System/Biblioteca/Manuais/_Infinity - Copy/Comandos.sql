verifica versão Me.Caption = "Lydians v" & Format(App.Major, "#00") & "." & Format(App.Minor, "000") & "." & Format(App.Revision, "0000")

Public Function sFormCentralizar(ByRef frmName As Form)
   On Error GoTo sFormCentralizar_E

   Dim llngAreaLivreH As Long
   Dim llngAreaLivreW As Long
   
   If Not frmName.MDIChild Then Exit Function
    
   frmName.WindowState = 0
                      
   llngAreaLivreW = Prj.Projeto.MDIForm.ScaleWidth
   llngAreaLivreH = Prj.Projeto.MDIForm.ScaleHeight
                      
   If (frmName.Height > llngAreaLivreH) Then
      frmName.Top = 0
   Else
      frmName.Top = ((llngAreaLivreH - frmName.Height) \ 2)
   End If
 
   If (frmName.Width > llngAreaLivreW) Then
      frmName.Left = 0
   Else
      frmName.Left = ((llngAreaLivreW - frmName.Width) \ 2)
   End If
   
   Exit Function
   
sFormCentralizar_E:
   MsgBox Err.Description
End Function

Public Function FormataDiaSemana(Data As Date, Optional Extenso As Boolean = True) As String
   Dim strDia As String
   Dim intDia As Integer
   
   FormataDiaSemana = ""
   
   If CStr(Data) = "0" Or Len(Trim(CStr(Data))) = 0 Then Exit Function
   
   intDia = WeekDay(Data)
   
   If Extenso Then
      Select Case intDia
      Case 1
         strDia = "Domingo"
      Case 2
         strDia = "Segunda-feira"
      Case 3
         strDia = "Terça-feira"
      Case 4
         strDia = "Quarta-feira"
      Case 5
         strDia = "Quinta-feira"
      Case 6
         strDia = "Sexta-feira"
      Case 7
         strDia = "Sábado"
      End Select
      FormataDiaSemana = strDia
   Else
      FormataDiaSemana = CStr(intDia)
   End If
   
End Function

Public Function FormataCICCGC(ByVal strDocumento As String) As String
   Dim strD As String
   Dim I As Integer
   strD = Trim(strDocumento)
   I = Len(strD)
   If I = 11 Then 'CIC
      FormataCICCGC = Format(strD, "@@@.@@@.@@@-@@")
   ElseIf I = 14 Then  'CGC
      FormataCICCGC = Format(strD, "@@.@@@.@@@/@@@@-@@")
   Else
      FormataCICCGC = strD
   End If

End Function

Public Function ObtemPrimeiroDiaMes(Data As Date) As Date
   ObtemPrimeiroDiaMes = DateSerial(Year(Data), Month(Data), 1)
End Function

Public Function ObtemPrimeiroDiaMesAnt(Data As Date) As Date
   Dim x As Date
   x = Me.ObtemUltimoDiaMesAnt(Data)
   ObtemPrimeiroDiaMesAnt = Me.ObtemPrimeiroDiaMes(x)
End Function

Public Function ObtemUltimoDiaMes(Data As Date) As Date
   Dim x As Date
   
   x = DateAdd("m", 1, Data)
   ObtemUltimoDiaMes = DateAdd("d", -1, DateSerial(Year(x), Month(x), 1))

End Function

Public Function ObtemUltimoDiaMesAnt(Data As Date) As Date
   Dim x As Date
   
   x = Me.ObtemPrimeiroDiaMes(Data)
   
   ObtemUltimoDiaMesAnt = DateAdd("d", -1, x)

End Function


Function VerificaCGC(ByVal CGC As String) As Boolean
   Dim intD1 As Integer
   Dim intD4 As Integer
   Dim intXX As Integer
   Dim intConta As Integer
   Dim intResto As Integer
   Dim intDigito As Integer
   
   VerificaCGC = False
   
'#ANG20130709 - ATD6647
'   CGC = Trim$(CGC)
   CGC = Trim$(SoNumeros(CGC))
'#FANG20130709 - ATD6647
   
   If Len(CGC) <> 14 Then Exit Function

   intD1 = 0
   intD4 = 0
   intXX = 1

   For intConta = 1 To 12
      intD1 = intD1 + Val(Mid(CGC, intConta, 1)) * (IIf(intXX < 5, 6, 14) - intXX)
      intD4 = intD4 + Val(Mid(CGC, intConta, 1)) * (IIf(intXX < 6, 7, 15) - intXX)
      intXX = intXX + 1
   Next

   intResto = intD1 - (Int(intD1 / 11) * 11)
   intDigito = IIf(intResto < 2, 0, (11 - intResto))
   intD4 = intD4 + 2 * intDigito
   intResto = intD4 - (Int(intD4 / 11) * 11)
   intDigito = Val(Mid(intDigito, 1) + Format$(IIf(intResto < 2, 0, 11 - intResto), "0"))

   If intDigito = Val(Mid(CGC, Len(CGC) - 1, 2)) Then
      VerificaCGC = True
   End If

End Function

Function VerificaCPF(ByVal CPF As String) As Boolean
   Dim intD1 As Long
   Dim intD2 As Long
   Dim intXX As Long
   Dim intConta As Long
   Dim intResto As Long
   Dim intDigito As Long
   
   'Calculo do Dígito do Cpf

   VerificaCPF = False

'#ANG20130709 - ATD6647
'   CPF = Trim$(CPF)
   CPF = Trim$(SoNumeros(CPF))
'#FANG20130709 - ATD6647
      
   If Len(CPF) <> 11 Then Exit Function

   intD1 = 0
   intD2 = 0
   intXX = 1

   For intConta = 1 To Len(CPF) - 2
      intD1 = intD1 + (11 - intXX) * Val(Mid(CPF, intConta, 1))
      intD2 = intD2 + (12 - intXX) * Val(Mid(CPF, intConta, 1))
      intXX = intXX + 1
   Next

   intResto = intD1 - Int(intD1 / 11) * 11
   intDigito = IIf(intResto < 2, 0, 11 - intResto)
   intD2 = intD2 + 2 * intDigito
   intResto = intD2 - Int(intD2 / 11) * 11

   intDigito = Val(Format$(intDigito, "0") + Format$(IIf(intResto < 2, 0, 11 - intResto), "0"))

   If intDigito = Val(Mid(CPF, Len(CPF) - 1, 2)) Then
      VerificaCPF = True
   End If

End Function

Public Function RegLer(Sistema As String, ByRef Chave As String, Optional Default As Variant) As String
   
   If Not IsMissing(Default) Then
      RegLer = GetSetting(Sistema, "Sistema", Chave, Default)
   Else
      RegLer = GetSetting(Sistema, "Sistema", Chave)
   End If
   
End Function

Public Sub RegApagar(Sistema As String, ByRef Chave As String)
   
   DeleteSetting Sistema, "Sistema", Chave
   
End Sub

Public Sub RegGravar(Sistema As String, ByRef Chave As String, ByRef Valor As String)
   
   SaveSetting Sistema, "Sistema", Chave, Valor

End Sub

Public Function SoNumeros(ByVal strPalavra As String) As String
   Dim I As Integer
   Dim Max As Integer
   Dim strNova As String
   Dim strCar As String * 1
   
   strNova = ""
   strPalavra = Trim(strPalavra)
   Max = Len(strPalavra)
   
   For I = 1 To Max
      strCar = Mid(strPalavra, I, 1)
      If IsNumeric(strCar) Then strNova = strNova & strCar
   Next
   
   SoNumeros = strNova
   
End Function
