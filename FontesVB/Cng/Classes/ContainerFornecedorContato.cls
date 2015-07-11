VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sContainerFornContato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private lngSeq As Long
Private lngCodigo As Long
Private strContato As String
Private strTelefone As String
Private strEmail As String
Private strAcao As String

Public Property Get Key() As String
   Key = "K_" & Sequencia
End Property

Public Property Get Sequencia() As Long
   Sequencia = lngSeq
End Property

Public Property Let Sequencia(ByVal vNewValue As Long)
   lngSeq = vNewValue
End Property

Public Property Get Contato() As String
   Contato = strContato
End Property

Public Property Let Contato(ByVal vNewValue As String)
   strContato = vNewValue
End Property

Public Property Get Telefone() As String
   Telefone = strTelefone
End Property

Public Property Let Telefone(ByVal vNewValue As String)
   strTelefone = vNewValue
End Property

Public Property Get Email() As String
   Email = strEmail
End Property

Public Property Let Email(ByVal vNewValue As String)
   strEmail = vNewValue
End Property

Public Property Get Acao() As String
   Acao = strAcao
End Property

Public Property Let Acao(ByVal vNewValue As String)
   strAcao = vNewValue
End Property

Public Property Get Codigo() As Long
   Codigo = lngCodigo
End Property

Public Property Let Codigo(ByVal vNewValue As Long)
   lngCodigo = vNewValue
End Property
