VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private blnRstOpen As Boolean
Private ADOCon As ADODB.Connection
Private ADORst As ADODB.Recordset
Public SQL As New sSQL

Public Sub Inicializar(Conexao As Object)
   Set ADOCon = Conexao
End Sub

Public Function Abrir(SQLTexto As String) As Boolean
   On Error GoTo Abrir_E
   
   Abrir = False
   
   If Not blnRstOpen Then Set ADORst = New ADODB.Recordset
   ADORst.CursorLocation = adUseClient
   ADORst.Open SQLTexto, ADOCon, adOpenStatic, adLockReadOnly
   blnRstOpen = True
   
   Abrir = True
   
   Exit Function

Abrir_E:
   clsErro.Salvar Err
   clsErro.ModRotina "Abrir"
End Function

Public Sub Fechar()
   If blnRstOpen Then ADORst.Close
   blnRstOpen = False
End Sub

Public Function Vlr(Valor As Variant) As Variant
   Vlr = Replace(Valor, ",", ".")
End Function

Public Function Txt(Texto As String) As String
   Txt = "'" & Texto & "'"
End Function

Public Function Dt(Data As Date) As String
   If Not IsDate(Data) Then
      Dt = "NULL"
   Else
      Dt = "{d '" & Format(Data, "YYYY-MM-DD") & "'"
   End If
End Function

Public Function EOF() As Boolean
   EOF = ADORst.EOF
End Function

Public Sub ProximoRegistro()
   ADORst.MoveNext
End Sub

Public Function Valor(CmpName As String) As Variant
   Valor = ADORst.Fields(CmpName)
   If IsNull(Valor) Then Valor = ""
End Function
