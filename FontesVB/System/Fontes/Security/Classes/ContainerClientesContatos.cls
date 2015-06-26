VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sContainerClientesContatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strAcao As String
Private lngSequencial As Long
Private lngCodContato As Long
Private strNome As String
Private strTelefone As String
Private strEmail As String

Public Property Get Key() As String
   Key = "K_" & Sequencial
End Property

Public Property Get Sequencial() As Long
   Sequencial = lngSequencial
End Property

Public Property Let Sequencial(ByVal vNewValue As Long)
   lngSequencial = vNewValue
End Property

Public Property Get CodContato() As Long
   CodContato = lngCodContato
End Property

Public Property Let CodContato(ByVal vNewValue As Long)
   lngCodContato = vNewValue
End Property

Public Property Get Nome() As String
   Nome = strNome
End Property

Public Property Let Nome(ByVal vNewValue As String)
   strNome = vNewValue
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
