Attribute VB_Name = "mVerifica"
Option Explicit

Public Function VerificaUsuario(objErro As Object, ID As ActiveText, Dsc As ActiveText, Optional Situacao As String = "") As Boolean
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
      
      If Len(Trim(Situacao)) > 0 Then .SQL.Mais " AND Situacao  = " & .Txt(Situacao)
      
      If Not .Abrir(.SQL.Texto) Then
         objErro.Transferir = .TransferirErro
         objErro.ModRotina = "VerificaUsuario"
         GoTo DestruirObjetos
      End If
      
      If Not .EOF Then
         Dsc = .Valor("Usuario")
      Else
         objErro.Salvar Err, 1, "Usuário não localizado! Verifique."
         objErro.ModRotina = "VerificaUsuario"
         mFocus ID
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

Public Function VerificaGrupo(objErro As Object, ID As ActiveText, Dsc As ActiveText, Optional Situacao As String = "") As Boolean
   On Error GoTo VerificaGrupo_E
   
   Dim clsCursor As INF_Cursor.Cursor
   
   VerificaGrupo = False
   
   If ID = 0 Then
      VerificaGrupo = True
      GoTo DestruirObjetos
   End If
   
   Set clsCursor = CreateObject("INF_Cursor.Cursor")
   With clsCursor
      .Inicializar clsConexao
      
      .SQL.Limpar
      .SQL.Mais " SELECT Descricao "
      .SQL.Mais " FROM Grupos "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND Codigo = " & .Vlr(ID)
      
      If Len(Trim(Situacao)) > 0 Then .SQL.Mais " AND Situacao  = " & .Txt(Situacao)
      
      If Not .Abrir(.SQL.Texto) Then
         objErro.Transferir = .TransferirErro
         objErro.ModRotina = "VerificaGrupo"
         GoTo DestruirObjetos
      End If
      
      If Not .EOF Then
         Dsc = .Valor("Descricao")
      Else
         objErro.Salvar Err, 1, "Grupo não localizado! Verifique."
         objErro.ModRotina = "VerificaGrupo"
         mFocus ID
         GoTo DestruirObjetos
      End If
      .Fechar
   End With
   
   VerificaGrupo = True
   
   GoTo DestruirObjetos
   
VerificaGrupo_E:
   objErro.Salvar Err
   objErro.ModRotina = "VerificaGrupo"

DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Function

Public Function VerificaSubGrupo(objErro As Object, ID As ActiveText, Dsc As ActiveText, Optional Situacao As String = "") As Boolean
   On Error GoTo VerificaSubGrupo_E
   
   Dim clsCursor As INF_Cursor.Cursor
   
   VerificaSubGrupo = False
   
   If ID = 0 Then
      VerificaSubGrupo = True
      GoTo DestruirObjetos
   End If
   
   Set clsCursor = CreateObject("INF_Cursor.Cursor")
   With clsCursor
      .Inicializar clsConexao
      
      .SQL.Limpar
      .SQL.Mais " SELECT Descricao "
      .SQL.Mais " FROM SubGrupo "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND Codigo = " & .Vlr(ID)
      
      If Len(Trim(Situacao)) > 0 Then .SQL.Mais " AND Situacao  = " & .Txt(Situacao)
      
      If Not .Abrir(.SQL.Texto) Then
         objErro.Transferir = .TransferirErro
         objErro.ModRotina = "VerificaSubGrupo"
         GoTo DestruirObjetos
      End If
      
      If Not .EOF Then
         Dsc = .Valor("Descricao")
      Else
         objErro.Salvar Err, 1, "SubGrupo não localizado! Verifique."
         objErro.ModRotina = "VerificaSubGrupo"
         mFocus ID
         GoTo DestruirObjetos
      End If
      .Fechar
   End With
   
   VerificaSubGrupo = True
   
   GoTo DestruirObjetos
   
VerificaSubGrupo_E:
   objErro.Salvar Err
   objErro.ModRotina = "VerificaSubGrupo"

DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Function

Public Function VerificaMarca(objErro As Object, ID As ActiveText, Dsc As ActiveText, Optional Situacao As String = "") As Boolean
   On Error GoTo VerificaMarca_E
   
   Dim clsCursor As INF_Cursor.Cursor
   
   VerificaMarca = False
   
   If ID = 0 Then
      VerificaMarca = True
      GoTo DestruirObjetos
   End If
   
   Set clsCursor = CreateObject("INF_Cursor.Cursor")
   With clsCursor
      .Inicializar clsConexao
      
      .SQL.Limpar
      .SQL.Mais " SELECT Descricao "
      .SQL.Mais " FROM Marcas "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND Codigo = " & .Vlr(ID)
      
      If Len(Trim(Situacao)) > 0 Then .SQL.Mais " AND Situacao  = " & .Txt(Situacao)
      
      If Not .Abrir(.SQL.Texto) Then
         objErro.Transferir = .TransferirErro
         objErro.ModRotina = "VerificaMarca"
         GoTo DestruirObjetos
      End If
      
      If Not .EOF Then
         Dsc = .Valor("Descricao")
      Else
         objErro.Salvar Err, 1, "Marca não localizado! Verifique."
         objErro.ModRotina = "VerificaMarca"
         mFocus ID
         GoTo DestruirObjetos
      End If
      .Fechar
   End With
   
   VerificaMarca = True
   
   GoTo DestruirObjetos
   
VerificaMarca_E:
   objErro.Salvar Err
   objErro.ModRotina = "VerificaMarca"

DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Function

Public Function VerificaModelo(objErro As Object, ID As ActiveText, Dsc As ActiveText, Optional Situacao As String = "") As Boolean
   On Error GoTo VerificaModelo_E
   
   Dim clsCursor As INF_Cursor.Cursor
   
   VerificaModelo = False
   
   If ID = 0 Then
      VerificaModelo = True
      GoTo DestruirObjetos
   End If
   
   Set clsCursor = CreateObject("INF_Cursor.Cursor")
   With clsCursor
      .Inicializar clsConexao
      
      .SQL.Limpar
      .SQL.Mais " SELECT Descricao "
      .SQL.Mais " FROM Modelos "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND Codigo = " & .Vlr(ID)
      
      If Len(Trim(Situacao)) > 0 Then .SQL.Mais " AND Situacao  = " & .Txt(Situacao)
      
      If Not .Abrir(.SQL.Texto) Then
         objErro.Transferir = .TransferirErro
         objErro.ModRotina = "VerificaModelo"
         GoTo DestruirObjetos
      End If
      
      If Not .EOF Then
         Dsc = .Valor("Descricao")
      Else
         objErro.Salvar Err, 1, "Modelo não localizado! Verifique."
         objErro.ModRotina = "VerificaModelo"
         mFocus ID
         GoTo DestruirObjetos
      End If
      .Fechar
   End With
   
   VerificaModelo = True
   
   GoTo DestruirObjetos
   
VerificaModelo_E:
   objErro.Salvar Err
   objErro.ModRotina = "VerificaModelo"

DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Function

Public Function VerificaUnidade(objErro As Object, ID As ActiveText, Dsc As ActiveText, Optional Situacao As String = "") As Boolean
   On Error GoTo VerificaUnidade_E
   
   Dim clsCursor As INF_Cursor.Cursor
   
   VerificaUnidade = False
   
   If ID = 0 Then
      VerificaUnidade = True
      GoTo DestruirObjetos
   End If
   
   Set clsCursor = CreateObject("INF_Cursor.Cursor")
   With clsCursor
      .Inicializar clsConexao
      
      .SQL.Limpar
      .SQL.Mais " SELECT Descricao "
      .SQL.Mais " FROM Unidades "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND Codigo = " & .Vlr(ID)
      
      If Len(Trim(Situacao)) > 0 Then .SQL.Mais " AND Situacao  = " & .Txt(Situacao)
      
      If Not .Abrir(.SQL.Texto) Then
         objErro.Transferir = .TransferirErro
         objErro.ModRotina = "VerificaUnidade"
         GoTo DestruirObjetos
      End If
      
      If Not .EOF Then
         Dsc = .Valor("Descricao")
      Else
         objErro.Salvar Err, 1, "Unidade não localizado! Verifique."
         objErro.ModRotina = "VerificaUnidade"
         mFocus ID
         GoTo DestruirObjetos
      End If
      .Fechar
   End With
   
   VerificaUnidade = True
   
   GoTo DestruirObjetos
   
VerificaUnidade_E:
   objErro.Salvar Err
   objErro.ModRotina = "VerificaUnidade"

DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Function

Public Function VerificaFornecedor(objErro As Object, ID As ActiveText, Dsc As ActiveText, Optional Situacao As String = "") As Boolean
   On Error GoTo VerificaFornecedor_E
   
   Dim clsCursor As INF_Cursor.Cursor
   
   VerificaFornecedor = False
   
   If ID = 0 Then
      VerificaFornecedor = True
      GoTo DestruirObjetos
   End If
   
   Set clsCursor = CreateObject("INF_Cursor.Cursor")
   With clsCursor
      .Inicializar clsConexao
      
      .SQL.Limpar
      .SQL.Mais " SELECT RazaoSocial "
      .SQL.Mais " FROM Fornecedores "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND Codigo = " & .Vlr(ID)
      
      If Len(Trim(Situacao)) > 0 Then .SQL.Mais " AND Situacao  = " & .Txt(Situacao)
      
      If Not .Abrir(.SQL.Texto) Then
         objErro.Transferir = .TransferirErro
         objErro.ModRotina = "VerificaFornecedor"
         GoTo DestruirObjetos
      End If
      
      If Not .EOF Then
         Dsc = .Valor("RazaoSocial")
      Else
         objErro.Salvar Err, 1, "Fornecedor não localizado! Verifique."
         objErro.ModRotina = "VerificaFornecedor"
         mFocus ID
         GoTo DestruirObjetos
      End If
      .Fechar
   End With
   
   VerificaFornecedor = True
   
   GoTo DestruirObjetos
   
VerificaFornecedor_E:
   objErro.Salvar Err
   objErro.ModRotina = "VerificaFornecedor"

DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Function
