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
         objErro.TransferirErro = .TransferirErro
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
