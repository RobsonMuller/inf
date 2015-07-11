VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public SQL As CmdSQL
Private clsErro As INF_Erro.Funcoes

'Microsoft ActiveX Data Object 2.8 Library - msado28.tlb
Private ADOCon As ADODB.Connection
Private ADORst As ADODB.Recordset
Private RstOpen As Boolean

Private Sub Class_Initialize()
   Set SQL = New CmdSQL
   Set clsErro = CreateObject("INF_Erro.Funcoes")
End Sub

Private Sub Class_Terminate()
   Set SQL = Nothing
   Set clsErro = Nothing
   Set ADOCon = Nothing
End Sub

Public Sub Inicializar(objCon As Object)
   Set ADOCon = objCon.Connect
End Sub

Public Property Get TransferirErro() As Object
   Set TransferirErro = clsErro.Transferir
End Property

Public Function Abrir(CmdSQL As String) As Boolean
   On Error GoTo Abrir_E
   
   Abrir = False
   
   Set ADORst = CreateObject("ADODB.RecordSet")
   
   ADORst.ActiveConnection = ADOCon
   ADORst.CursorType = adOpenForwardOnly
   ADORst.CursorLocation = adUseClient
      
   ADORst.Open CmdSQL
   
   RstOpen = True
   
   Abrir = True
   
   Exit Function

Abrir_E:
   clsErro.Salvar Err, , , CmdSQL
   clsErro.ModRotina = "Abrir"
End Function

Public Sub Fechar()
   On Error Resume Next
   If RstOpen Then ADORst.Close
   RstOpen = False
End Sub

Public Property Get EOF() As Boolean
   EOF = ADORst.EOF
End Property

Public Sub ProximoRegistro()
   ADORst.MoveNext
End Sub

Public Property Get Valor(NmCmp As String) As String
   On Error GoTo Valor_E
   
   Valor = ADORst.Fields(NmCmp)
   
   Exit Property

Valor_E:
   'Erros tratados
   Select Case Err.Number
   Case 94 '(Invalid use of Null)
      Valor = ""
      Exit Property
   End Select
End Property

Public Property Get Vlr(Valor As Variant, Optional ComVirgula As Boolean = False) As String
   Vlr = Replace(CStr(Valor), ",", ".") & IIf(ComVirgula, ",", "")
End Property

Public Property Get Txt(Texto As String, Optional ComVirgula As Boolean = False) As String
   Txt = "'" & Texto & "'" & IIf(ComVirgula, ", ", "")
End Property

Public Property Get Dt(Data As Variant, Optional ComVirgula As Boolean = False, Optional ComHora As Boolean = False) As String
   If Not IsDate(Data) Then
      Dt = "CONVERT(datetime, NULL)"
   ElseIf IsDate(Data) Then
      If Year(Data) < 1900 Then
         Dt = "CONVERT(datetime, NULL)"
      Else
         If ComHora Then
            Dt = "{ts '" & Format$(Data, "yyyy-mm-dd hh:mm:ss") & "'}"
         Else
            Dt = "{d '" & Format(Data, "yyyy-mm-dd") & "'}"
         End If
      End If
   End If
   
   If ComVirgula Then Dt = Dt & ", "
End Property

'COMPABILIDADE COM O VISUAL BASIC 5
'VB6 JÁ POSSUI O REPLACE
Private Function Replace(Source As String, Find As String, ReplaceStr As String, _
    Optional ByVal Start As Long = 1, Optional Count As Long = -1, _
    Optional Compare As VbCompareMethod = vbBinaryCompare) As String

   Dim findLen As Long
   Dim replaceLen As Long
   Dim index As Long
   Dim counter As Long
   
   findLen = Len(Find)
   replaceLen = Len(ReplaceStr)
   If findLen = 0 Then Err.Raise 5
   
   If Start < 1 Then Start = 1
   index = Start
   Replace = Source
   Do
      index = InStr(index, Replace, Find, Compare)
      If index = 0 Then Exit Do
      If findLen = replaceLen Then
         Mid$(Replace, index, findLen) = ReplaceStr
      Else
         Replace = Left$(Replace, index - 1) & ReplaceStr & Mid$(Replace, index + findLen)
      End If
      index = index + replaceLen
      counter = counter + 1
   Loop Until counter = Count
   If Start > 1 Then Replace = Mid$(Replace, Start)
End Function


