Attribute VB_Name = "mVerifica"
Option Explicit

Public Function mVerificaCidade(pclsErro As INF_Erro.Funcoes, Codigo As ActiveText, retMunicipio As ActiveText, Optional retUF As ComboBox = Nothing) As Boolean
   On Error GoTo mVerificaCidade_E
   
   Dim clsCursor As INF_Cursor.Cursor
   
   mVerificaCidade = False
   
   If Len(Trim(Codigo)) > 0 Then
      If IsNumeric(Codigo) Then
         Set clsCursor = CreateObject("INF_Cursor.Cursor")
         With clsCursor
            .Inicializar clsConexao
            
            .SQL.Limpar
            .SQL.Mais " SELECT Municipio, UF "
            .SQL.Mais " FROM Municipios "
            .SQL.Mais " WHERE Codigo = " & .Vlr(Codigo)
                        
            If Not .Abrir(.SQL.Texto) Then
               pclsErro.Transferir = .TransferirErro
               pclsErro.ModRotina = "mVerificaCidade"
               GoTo DestruirObjetos
            End If
            
            If Not .EOF Then
               retMunicipio = .Valor("Municipio")
               On Error Resume Next
               retUF.Text = .Valor("UF")
               On Error GoTo mVerificaCidade_E
            Else
               pclsErro.Salvar Err, 1, "Município não localizado! Verifique."
               pclsErro.ModRotina = "mVerificaCidade"
               GoTo DestruirObjetos
            End If
            .Fechar
         End With
      End If
   End If
   
   mVerificaCidade = True
   
   GoTo DestruirObjetos
         
mVerificaCidade_E:
   pclsErro.Salvar Err
   pclsErro.ModRotina = "mVerificaCidade"

DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Function

Public Function VerificaUsuario(objErro As Object, ID As ActiveText, Dsc As ActiveText, Optional Ativo As Boolean = True) As Boolean
   On Error GoTo VerificaUsuario_E
   
   Dim clsCursor As INF_Cursor.Cursor
   
   VerificaUsuario = False
   
   If ID = 0 Then
      VerificaUsuario = True
      GoTo DestruirObjetos
   End If
   
   Set clsCursor = CreateObject("INF_Cursor.Cursor")
   With clsCursor
      .Inicializar clsConexao
      
      .SQL.Limpar
      .SQL.Mais " SELECT Usuario "
      .SQL.Mais " FROM Usuarios "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND Codigo = " & .Vlr(ID)
      .SQL.Mais " AND IdSituacao = " & .Vlr(IIf(Ativo, 1, 0))
      
      If Not .Abrir(.SQL.Texto) Then
         objErro.TransferirErro = .TransferirErro
         objErro.ModRotina = "VerificaUsuario"
         GoTo DestruirObjetos
      End If
      
      If Not .EOF Then
         Dsc = .Valor("Usuario")
      Else
         objErro.Salvar Err, 1, "Usuário não localizado! Verifique."
         objErro.ModRotina = "VerificaUsuario"
         GoTo DestruirObjetos
      End If
      .Fechar
   End With
   
   VerificaUsuario = True
   
   GoTo DestruirObjetos
   
VerificaUsuario_E:
   objErro.Salvar Err
   objErro.ModRotina = "VerificaUsuario"

DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Function

