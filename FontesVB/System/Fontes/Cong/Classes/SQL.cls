VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strTexto As String

Public Sub Mais(SQLTexto As String)
   strTexto = strTexto & SQLTexto
End Sub

Public Function Texto() As String
   Texto = strTexto
End Function

Public Sub Limpar()
   strTexto = ""
End Sub

Public Function Tamanho() As Currency
   Tamanho = Len(strTexto)
End Function
