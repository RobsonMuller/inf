VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBasicCad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Esta classe é utilizada para os cadastro básicos

'Marca
'Modelo
'Unidade
'Categoria
'Grupo
'SubGrupo

'Tais cadastros são chamados apartir da interface BasicCad.
Private clsErro As INF_Erro.Funcoes
Private lngCod As Long
Private strDesc As String
Private strDescAbrev As String
Private strSit As String
Private strTable As String

Public Property Get Cod() As Long
   Cod = lngCod
End Property

Public Property Let Cod(ByVal vNewValue As Long)
   lngCod = vNewValue
End Property

Public Property Get Desc() As String
   Desc = strDesc
End Property

Public Property Let Desc(ByVal vNewValue As String)
   strDesc = vNewValue
End Property

Public Property Get DescAbrev() As String
   DescAbrev = strDescAbrev
End Property

Public Property Let DescAbrev(ByVal vNewValue As String)
   strDescAbrev = vNewValue
End Property

Public Property Get Sit() As String
   Sit = strSit
End Property

Public Property Let Sit(ByVal vNewValue As String)
   strSit = vNewValue
End Property

Public Property Get TransferirErro() As Object
   Set TransferirErro = clsErro.Transferir
End Property

Private Sub Class_Initialize()
   Set clsErro = CreateObject("INF_Erro.Funcoes")
End Sub

Private Sub Class_Terminate()
   Set clsErro = Nothing
End Sub

Public Function CurBasicCad() As Boolean
   On Error GoTo CurBasicCad_E
   
   CurBasicCad = False
   
   Dim clsCursor As INF_Cursor.Cursor
   
   If Cod = 0 Then
      clsErro.Salvar Err, 1, "Código não informado! Verifique."
      clsErro.ModRotina = "CurBasicCad"
      GoTo DestruirObjetos
   End If
   
   If Len(Trim(strTable)) = 0 Then
      clsErro.Salvar Err, 1, "Erro de desenvolvimento. Tabela não informado! Comunique o administrador do sistema."
      clsErro.ModRotina = "CurBasicCad"
      GoTo DestruirObjetos
   End If
   
   Set clsCursor = CreateObject("INF_Cursor.Cursor")
   With clsCursor
      .Inicializar clsConexao
      
      .SQL.Limpar
      .SQL.Mais " SELECT Empresa, Codigo, Descricao, Abreviatura, Situacao "
      .SQL.Mais " FROM " & strTable
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND Codigo = " & .Vlr(Cod)
      
      If Not .Abrir(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         clsErro.ModRotina = "CurBasicCad"
         GoTo DestruirObjetos
      End If
      
      If .EOF Then
         clsErro.Salvar Err, 2, "Registro não localizado! Verifique."
         clsErro.ModRotina = "CurBasicCad"
         GoTo DestruirObjetos
      Else
         Desc = .Valor("Descricao")
         DescAbrev = .Valor("Abreviatura")
         Sit = .Valor("Situacao")
      End If
      .Fechar
   End With
   
   CurBasicCad = True
   
   GoTo DestruirObjetos

CurBasicCad_E:
   clsErro.Salvar Err
   clsErro.ModRotina = "CurBasicCad"

DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Function

Public Property Let Table(ByVal vNewValue As String)
   strTable = vNewValue
End Property
