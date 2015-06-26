VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sContainerCG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private lngCod As Long
Private strDesc As String

Public Property Get Key() As String
   Key = "K_" & Codigo
End Property

Public Property Get Codigo() As Long
   Codigo = lngCod
End Property

Public Property Let Codigo(ByVal vNewValue As Long)
   lngCod = vNewValue
End Property

Public Property Get Descricao() As String
   Descricao = strDesc
End Property

Public Property Let Descricao(ByVal vNewValue As String)
   strDesc = vNewValue
End Property
