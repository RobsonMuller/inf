VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sErro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strRot As String
Private strCab As String
Private strNumber As String
Private strDesc As String

Public Sub Salvar(objErro As ErrObject)
   strCab = ""
   strNumber = ""
   strDesc = ""
   
   strCab = "Infinity - Erro (" & Now & ")"
   strNumber = "Erro Number: " & objErro.Number
   strDesc = "Erro Description: " & objErro.Description
End Sub

Public Sub ModRotina(strRotina As String)
   If Len(Trim(strRot)) > 0 Then
      strRot = strRot & "\"
   Else
      strRot = strRot & "Erro Rotina: "
   End If
   strRot = strRot & strRotina
End Sub

Public Sub Exibir()
   Dim fso As Object
   Dim out As TextStream
   Dim FileName As String
   Dim strMsg As String
   
   Set fso = CreateObject("Scripting.FileSystemObject")
   
   FileName = Dir("C:\_Infinity\Teste.txt", vbDirectory)
   
   If Len(Trim(FileName)) = 0 Then
      Set out = fso.CreateTextFile("C:\_Infinity\Teste.txt", True, True)
   Else
      Set out = fso.OpenTextFile(FileName, ForAppending, True, -1)
   End If
   
   out.WriteLine ""
   out.WriteLine strCab
   out.WriteLine strNumber
   out.WriteLine strDesc
   out.WriteLine strRot
   
   out.Close
   
   strMsg = ""
   strMsg = strMsg & strNumber & vbNewLine
   strMsg = strMsg & strDesc & vbNewLine
   strMsg = strMsg & strRot
   
   MsgBox strMsg, vbCritical, "Infinity - Erro"
   
End Sub
