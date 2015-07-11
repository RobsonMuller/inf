Attribute VB_Name = "mInfinity"
Option Explicit

Sub Main()
   Set f1 = New sFuncoes
   Set clsConexao = New INF_Conexao.Conexao
   
   InitCommonControls ' try Win9x version
   frmConnection.Show
End Sub

Public Property Get ModoDesenvolvimento() As Boolean
#If desenv = 1 Then
   ModoDesenvolvimento = True
#Else
   ModoDesenvolvimento = False
#End If
End Property

Public Sub mFocus(objCmp As Object)
   On Error Resume Next
   objCmp.SetFocus
   Exit Sub
End Sub

Public Sub mMsgInfo(Texto As String)
   MsgBox Texto, vbInformation, "Infinity - Informação"
End Sub

Public Sub mMsgAten(Texto As String)
   MsgBox Texto, vbCritical, "Infinity - Atenção"
End Sub

Public Function mMsgPerg(Texto As String) As Boolean
   If vbYes = MsgBox(Texto, vbYesNo, "Infinity - Pergunta") Then mMsgPerg = True
End Function

Public Function mCmpObrigatorio(objErro As Object, objCmp As Object, Descricao As String) As Boolean
   mCmpObrigatorio = False
   Select Case Left(objCmp.Name, 3)
   Case "txt"
      If Len(Trim(objCmp)) = 0 Then
         objErro.Salvar Err, 1, "O campo " & Descricao & " é de preenchimento obrigatório! Verifique."
         objCmp.SetFocus
         Exit Function
      End If
   Case "vlr"
      If objCmp = 0 Then
         objErro.Salvar Err, 1, "O campo " & Descricao & " é de preenchimento obrigatório! Verifique."
         objCmp.SetFocus
         Exit Function
      End If
   Case "dat"
      If Not IsDate(objCmp.Text) Then
         objErro.Salvar Err, 1, "O campo " & Descricao & " é de preenchimento obrigatório! Verifique."
         objCmp.SetFocus
         Exit Function
      End If
   End Select
   mCmpObrigatorio = True
End Function

Public Function HabilitarBotao(objErro As Object, objForm As Object, objButton As Object) As Boolean
   On Error GoTo HabilitarBotao_E
   
   Dim clsCursor As INF_Cursor.Cursor
   
   Set clsCursor = New INF_Cursor.Cursor
   With clsCursor
      .Inicializar clsConexao
      
      .SQL.Limpar
      .SQL.Mais " SELECT GlbPermissoes.IdButton "
      .SQL.Mais " FROM GlbPermissoes "
      .SQL.Mais " INNER JOIN GlbButton ON (GlbPermissoes.IdButton = GlbButton.Codigo)"
      .SQL.Mais " WHERE GlbPermissoes.Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND GlbPermissoes.IdUsuario = " & .Vlr(Prj.Sistema.IdUsuario)
      .SQL.Mais " AND GlbPermissoes.IdInterface = " & .Vlr(objForm.Tag)
      .SQL.Mais " AND GlbButton.Descricao = " & .Txt(Replace(objButton.Caption, "&", ""))
      
      If Not .Abrir(.SQL.Texto) Then
         objErro.TransferirErro = .TransferirErro
         objErro.ModRotina = "HabilitaBotao"
         GoTo DestruirObjetos
      End If
      
      If Not .EOF Then objButton.Enabled = True
      .Fechar
   End With
   
   HabilitarBotao = True
   
   GoTo DestruirObjetos

HabilitarBotao_E:
   objErro.Salvar Err
   objErro.ModRotina = "HabilitarBotao"
   
DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Function

Public Sub Exibir(objErr As Object, Rotina As String)
   Dim fso As Object
   Dim out As TextStream
   Dim FileName As String
   Dim strMsg As String
   
   'Informações para o arquivo
   '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-'
   Set fso = CreateObject("Scripting.FileSystemObject")
   
   FileName = Dir("C:\_Infinity\Teste.txt", vbDirectory)
   
   If Len(Trim(FileName)) = 0 Then
      Set out = fso.CreateTextFile("C:\Teste.txt", True, True)
   Else
      Set out = fso.OpenTextFile(FileName, ForAppending, True, -1)
   End If
   
   out.WriteLine ""
   out.WriteLine "Error - ( " & Now & " )"
   out.WriteLine "Number: " & objErr.Transferir.NroVB
   out.WriteLine "Description: " & objErr.Transferir.DscVB
   out.WriteLine "Routine: " & objErr.Transferir.ModRotPrj
   
   out.Close
   '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-'

   If objErr.Transferir.NroPrj > 0 Then
      MsgBox objErr.Transferir.DscPrj, vbInformation, "Infinity - Informativo"
   Else
      strMsg = "Error Number: " & objErr.Transferir.NroVB & vbNewLine
      strMsg = strMsg & "Description: " & objErr.Transferir.DscVB & vbNewLine
      strMsg = strMsg & "Mod. Project: " & Rotina & IIf(Len(Trim(objErr.Transferir.ModRotPrj)) > 0, "\", "") & objErr.Transferir.ModRotPrj
      
      MsgBox strMsg, vbCritical, "Infinity - Erro"
   End If
End Sub

Public Sub Ampulheta(Ativar As Boolean)
   MDIInfinity.MousePointer = 0
   If Ativar Then MDIInfinity.MousePointer = 11
End Sub
