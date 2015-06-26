VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sFuncoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum EnPesqCmb
   enCodigo = 0
   enDescricao = 1
End Enum

Public Function NmComp() As String
   Dim strNome As String
   Dim lretorno As Long

   strNome = String(255, Chr$(0))
   lretorno = GetComputerName(strNome, 255)
   strNome = Left$(strNome, InStr(1, strNome, Chr$(0)) - 1)
   NmComp = strNome
   Exit Function
End Function

Public Function FormCentralizar(ByRef frmName As Form)
   On Error GoTo FormCentralizar_E

   Dim llngAreaLivreH As Long
   Dim llngAreaLivreW As Long
   
   If Not frmName.MDIChild Then Exit Function
    
   frmName.WindowState = 0
                      
   llngAreaLivreW = MDIInfinity.ScaleWidth
   llngAreaLivreH = MDIInfinity.ScaleHeight
                      
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
   
   'Carrega o icone no form
   frmName.Icon = LoadResPicture(ICO_SYSTEM, vbResIcon)
   
   Exit Function
   
FormCentralizar_E:
   MsgBox Err.Description, vbCritical, "Infinity - Funçoes (Erro)"
End Function

Public Function FormataDiaSemana(Data As Date, Optional Extenso As Boolean = True) As String
   Dim strDia As String
   Dim intDia As Integer
   
   FormataDiaSemana = ""
   
   If CStr(Data) = "0" Or Len(Trim(CStr(Data))) = 0 Then Exit Function
   
   intDia = Weekday(Data)
   
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
   CGC = Trim$(SoNumeros(CGC))
   
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

   VerificaCPF = False
   CPF = Trim$(SoNumeros(CPF))
      
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

Public Function MaiorEntre(vlrA As Variant, vlrB As Variant) As Variant
   MaiorEntre = vlrA
   If vlrA < vlrB Then MaiorEntre = vlrB
End Function

Public Function MenorEntre(vlrA As Variant, vlrB As Variant) As Variant
   MenorEntre = vlrB
   If vlrA > vlrB Then MenorEntre = vlrA
End Function

Public Sub CollectionLimpar(objCollection As Collection)
   On Error GoTo CollectionLimpar_E
   
   Dim lngCount As Long
   
   For lngCount = 1 To objCollection.Count
      objCollection.Remove 1
   Next lngCount
   
   Exit Sub

CollectionLimpar_E:
   MsgBox Err.Description, vbCritical, "Infinity - Funçoes (Erro)"
End Sub

Public Sub Limpar(objForm As Form)
   On Error GoTo Limpar_E
   
   Dim Cnt As Control
   For Each Cnt In objForm.Controls
      If TypeOf Cnt Is TextBox Then Cnt.Text = ""
      If TypeOf Cnt Is ActiveText Then Cnt.Text = ""
      If TypeOf Cnt Is ListBox Then Cnt.Clear
   Next
   
   Exit Sub

Limpar_E:
   MsgBox Err.Description, vbCritical, "Infinity - Funçoes (Erro)"
End Sub

Public Sub CmbAdd(objCmb As Object, Descricao As String, Codigo As Long)
   On Error GoTo CmbAdd_E
   
   If Not TypeName(objCmb) = "ComboBox" Then Exit Sub
   
   objCmb.AddItem (Descricao)
   objCmb.ItemData(objCmb.NewIndex) = Codigo
   
   Exit Sub
   
CmbAdd_E:
   MsgBox Err.Description, vbCritical, "Infinity - Funções (Erro)"
End Sub

Public Function CmbValor(objCmp As Object, Valor As Variant, Optional TpPesquisa As EnPesqCmb = enCodigo, Optional Tamanho As Long) As Integer
   On Error GoTo CmbValor_E
   
   Dim intCount As Integer
   Dim lngTamanhoAux As Long
   Dim blnValorEncontrado As Boolean
   
   If Not TypeName(objCmp) = "ComboBox" Then Exit Function
   
   objCmp.ListIndex = -1
   
   blnValorEncontrado = False
   For intCount = 0 To objCmp.ListCount - 1
      objCmp.ListIndex = intCount
      
      Select Case TpPesquisa
      Case EnPesqCmb.enCodigo
         If IsNumeric(Valor) Then
            If CInt(Trim(objCmp.ItemData(intCount))) = CInt(Valor) Then
               blnValorEncontrado = True
               Exit For
            End If
         End If
      Case EnPesqCmb.enDescricao
         If Tamanho = 0 Then lngTamanhoAux = Len(objCmp.Text)
            
         If UCase(Left(objCmp.Text, IIf(Tamanho = 0, lngTamanhoAux, Tamanho))) = UCase(Valor) Then
            blnValorEncontrado = True
            Exit For
         End If
         lngTamanhoAux = 0
      End Select
   Next
   
   Select Case blnValorEncontrado
   Case True
      CmbValor = intCount
   Case False
      If objCmp.ListCount > 0 Then
         CmbValor = 0
      Else
         CmbValor = -1
      End If
   End Select
   
   Exit Function

CmbValor_E:
   MsgBox Err.Description, vbCritical, "Infinity - Funçoes (Erro)"   'O erro é exibido de dentro da classe por que a função não possui retorno
End Function

Public Function CmbParametro(objCmp As ComboBox) As Long
   On Error GoTo CmbParametro_E
   
   CmbParametro = objCmp.ItemData(objCmp.ListIndex)
   
   Exit Function

CmbParametro_E:
   MsgBox Err.Description, vbCritical, "Infinity - Funçoes (Erro)" 'O erro é exibido de dentro da classe por que a função não possui retorno
End Function

Public Function CmbIndex(objCmp As ComboBox) As Integer
   On Error GoTo CmbIndex_E
   
   CmbIndex = objCmp.ListIndex
   
   Exit Function

CmbIndex_E:
   MsgBox Err.Description, vbCritical, "Infinity - Funçoes (Erro)" 'O erro é exibido de dentro da classe por que a função não possui retorno
End Function

Public Function VerificaConexaoInternet(objErro As INF_Erro.Funcoes) As Boolean
   On Error GoTo VerificaConexaoInternet_E
   
   Dim Ret As Long
   Dim sConnType As String * 255
   
   VerificaConexaoInternet = False
   
   Ret = InternetGetConnectedStateEx(Ret, sConnType, 254, 0)
   
   If Ret = 1 Then
      objErro.Salvar Err, 1, "Você está conectado a Internet via " & sConnType
   Else
      objErro.Salvar Err, 1, "Você não está conectado a internet."
   End If
   
   objErro.ModRotina = "VerificaConexaoInternet"
   
   VerificaConexaoInternet = True
   
   Exit Function

VerificaConexaoInternet_E:
   objErro.Salvar Err
   objErro.ModRotina = "VerificaConexaoInternet"
End Function

