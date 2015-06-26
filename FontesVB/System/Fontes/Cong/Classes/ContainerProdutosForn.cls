VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sContainerProdutosForn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strSituacao As String 'I - Inserido, A - Alterado, G - Gravado, E - Excluido
Private lngCodFornecedor As Long
Private strDscFornecedor As String
Private strTelefone As String
Private curLucro As Currency
Private curValorCompra As Currency
Private curLucroValor As Currency

Public Property Get Key() As String
   Key = "K_" & CodFornecedor
End Property

Public Property Get CodFornecedor() As Long
   CodFornecedor = lngCodFornecedor
End Property

Public Property Let CodFornecedor(ByVal vNewValue As Long)
   lngCodFornecedor = vNewValue
End Property

Public Property Get Situacao() As String
   Situacao = strSituacao
End Property

Public Property Let Situacao(ByVal vNewValue As String)
   strSituacao = vNewValue
End Property

Public Property Get DscFornecedor() As String
   DscFornecedor = strDscFornecedor
End Property

Public Property Let DscFornecedor(ByVal vNewValue As String)
   strDscFornecedor = vNewValue
End Property

Public Property Get Telefone() As String
   Telefone = strTelefone
End Property

Public Property Let Telefone(ByVal vNewValue As String)
   strTelefone = vNewValue
End Property

Public Property Get Lucro() As Currency
   Lucro = curLucro
End Property

Public Property Let Lucro(ByVal vNewValue As Currency)
   curLucro = vNewValue
End Property

Public Property Get ValorCompra() As Currency
   ValorCompra = curValorCompra
End Property

Public Property Let ValorCompra(ByVal vNewValue As Currency)
   curValorCompra = vNewValue
End Property

Public Property Get LucroValor() As Currency
   LucroValor = curLucroValor
End Property

Public Property Let LucroValor(ByVal vNewValue As Currency)
   curLucroValor = vNewValue
End Property
