VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sConexao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public SQL As New sSQL
Private ADOCmd As ADODB.Command
Private ADOCon As ADODB.Connection

Public Sub Begin()
   ADOCon.BeginTrans
End Sub

Public Sub Commit()
   ADOCon.CommitTrans
End Sub

Public Sub RollBack()
   ADOCon.RollbackTrans
End Sub

Public Function Vlr(Valor As Variant, Optional ComVirgula As Boolean = False) As Variant
   Dim vlrValor As Variant
   vlrValor = Replace(Valor, ",", ".")
   If vlrValor = "" Then vlrValor = 0
   If ComVirgula Then vlrValor = vlrValor & ","
   Vlr = vlrValor
End Function

Public Function Txt(Texto As String, Optional ComVirgula As Boolean = False) As String
   Dim strTexto As String
   If Texto = "" Then Texto = "NULL"
   strTexto = "'" & Texto & "'"
   If ComVirgula Then strTexto = strTexto & ","
   Txt = strTexto
End Function

Public Function Dt(Data As String, Optional ComVirgula As Boolean = False) As String
   If Not IsDate(Data) Then
      Dt = "NULL"
   Else
      Dt = "{d '" & Format(Data, "YYYY-MM-DD") & "'"
   End If
   If ComVirgula Then Dt = Dt & ", "
End Function

Public Function DataServer(Optional ComVirgula As Boolean = False) As String
   DataServer = "GetDate()"
   If ComVirgula Then DataServer = DataServer & ", "
End Function

Public Function Executar(SQLTexto As String) As Boolean
   On Error GoTo Executar_E
   
   Executar = False
   Set ADOCmd = New ADODB.Command
   ADOCmd.ActiveConnection = ADOCon
   ADOCmd.CommandText = SQLTexto
   ADOCmd.Execute
   Executar = True
   
   Exit Function

Executar_E:
   clsErro.Salvar Err
   clsErro.ModRotina "Executar"
End Function

Private Sub Class_Initialize()
   Set ADOCon = AdoCn
End Sub

Private Sub Class_Terminate()
   Set SQL = Nothing
   Set ADOCon = Nothing
End Sub
