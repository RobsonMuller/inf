VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sContainerPermissoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private curCodInterface As Currency
Private lngCodButton As Long
Private strDscButton As String
Private blnSelecionado As Boolean

Public Property Get Key() As String
   Key = "K_" & IDInterface & IDButton
End Property

Public Property Get IDInterface() As Currency
   IDInterface = curCodInterface
End Property

Public Property Let IDInterface(ByVal vNewValue As Currency)
   curCodInterface = vNewValue
End Property

Public Property Get DescButton() As String
   DescButton = strDscButton
End Property

Public Property Let DescButton(ByVal vNewValue As String)
   strDscButton = vNewValue
End Property

Public Property Get IDButton() As Long
   IDButton = lngCodButton
End Property

Public Property Let IDButton(ByVal vNewValue As Long)
   lngCodButton = vNewValue
End Property

Public Property Get Selecionado() As Boolean
   Selecionado = blnSelecionado
End Property

Public Property Let Selecionado(ByVal vNewValue As Boolean)
   blnSelecionado = vNewValue
End Property
