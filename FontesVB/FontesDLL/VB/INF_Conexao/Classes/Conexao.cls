VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Conexao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public SQL As CmdSQL
Public FB As FunctionBase
Private clsErro As INF_Erro.Funcoes

Private strConnectionString As String
Private ConexOpen As Boolean
Private TransOpen As Boolean
Private ADOCon As ADODB.Connection
Private ADOCommand As ADODB.Command

Public Property Get ConnectionString() As String
   ConnectionString = strConnectionString
End Property

Public Property Get Connect() As Object
   Set Connect = ADOCon
End Property

Public Function Abrir(Connect As String) As Boolean
   On Error GoTo Abrir_E
   
   Abrir = False
   strConnectionString = Connect
   ADOCon.ConnectionString = strConnectionString
   ADOCon.Open
   
   ConexOpen = True
   Abrir = True
   
   Exit Function

Abrir_E:
   clsErro.Salvar Err
   clsErro.ModRotina = "Abrir"
End Function

Public Sub Begin()
   If Not TransOpen Then ADOCon.BeginTrans
   TransOpen = True
End Sub

Public Sub Commit()
   ADOCon.CommitTrans
   TransOpen = False
End Sub

Public Sub RollBack()
   On Error Resume Next
   ADOCon.RollbackTrans
   TransOpen = False
End Sub

Private Sub Class_Initialize()
   Set ADOCon = CreateObject("ADODB.Connection")
   Set clsErro = CreateObject("INF_Erro.Funcoes")
   Set SQL = New CmdSQL
   Set FB = New FunctionBase
End Sub

Private Sub Class_Terminate()
   If ConexOpen Then ADOCon.Close
   Set ADOCon = Nothing
   Set clsErro = Nothing
   Set SQL = Nothing
End Sub

Public Function Executar(CmdSQL As String) As Boolean
   On Error GoTo Executar_E
   
   Executar = False
   
   Set ADOCommand = CreateObject("ADODB.Command")
   ADOCommand.ActiveConnection = ADOCon
   ADOCommand.Prepared = True
   ADOCommand.CommandText = CmdSQL
   ADOCommand.Execute
   
   Executar = True
   
   Exit Function
   
Executar_E:
   clsErro.Salvar Err
   clsErro.ModRotina = "Executar"
End Function

Public Property Get TransferirErro() As Object
   Set TransferirErro = clsErro.Transferir
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
