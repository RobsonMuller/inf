Attribute VB_Name = "mDeclaration"
Option Explicit

Public f1 As sFuncoes
Public clsErro As INF_Erro.Funcoes
Public clsConexao As INF_Conexao.Conexao
Public Const NM_APP As String = "Infinity"
Public Prj As tyPrj

Public Type tySystem
   ConexActive As Boolean
   IdEmpresa As String
   IdUsuario As Long
   DtMovimento As Date
   RazaoSocial As String
   NomeFantasia As String
   CPFCNPJ As String
End Type

Public Type tyServer
   ConnectionString As String
End Type

Public Type tyPrj
   Sistema As tySystem
   Servidor As tyServer
End Type

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long

Public Const LST_ICO_GRAVADO As String = "Gravado"
Public Const LST_ICO_INSERIDO As String = "Inserido"
Public Const LST_ICO_ALTERADO As String = "Alterado"
Public Const LST_ICO_REMOVIDO As String = "Removido"

Public Const PATH_IMG As String = "C:\_Infinity\Imagens\"
Public Const LINK_EMAIL As String = "http://schemas.microsoft.com/cdo/configuration/"

'RESOURCE
Public Const ICO_SYSTEM As Integer = 101

