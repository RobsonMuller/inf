Attribute VB_Name = "mDeclarations"
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

Public Const rel_Marcas_Produtos As String = "Marcas.Empresa = Produtos.Empresa AND Marcas.Codigo = Produtos.CodMarca"
Public Const rel_Modelos_Produtos As String = "Modelos.Empresa = Produtos.Empresa AND Modelos.Codigo = Produtos.CodModelo"
Public Const rel_Unidades_Produtos As String = "Unidades.Empresa = Produtos.Empresa AND Unidades.Codigo = Produtos.CodUnidade"
Public Const rel_Grupos_Produtos As String = "Grupos.Empresa = Produtos.Empresa AND Grupos.Codigo = Produtos.CodGrupo"
Public Const rel_SubGrupos_Produtos As String = "SubGrupos.Empresa = Produtos.Empresa AND SubGrupos.Codigo = Produtos.CodSubGrupo"
Public Const rel_Usuarios_Produtos As String = "Usuarios.Empresa = Produtos.Empresa AND Usuarios.Codigo = Produtos.CodUsuario"
Public Const rel_Fornecedores_ProdutosForn As String = "Fornecedores.Empresa = ProdutosForn.Empresa AND Fornecedores.Codigo = ProdutosForn.CodFornecedor"
Public Const rel_PathArq_Produtos As String = "PathArq.Empresa = Produtos.Empresa AND PathArq.CodProduto = Produtos.Codigo"

Public Const LST_ICO_GRAVADO As String = "Gravado"
Public Const LST_ICO_INSERIDO As String = "Inserido"
Public Const LST_ICO_ALTERADO As String = "Alterado"
Public Const LST_ICO_REMOVIDO As String = "Removido"

Public Const PATH_IMG As String = "C:\_Infinity\Imagens\"
Public Const LINK_EMAIL As String = "http://schemas.microsoft.com/cdo/configuration/"

'RESOURCE
Public Const ICO_SYSTEM As Integer = 101
