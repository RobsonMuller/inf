VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sPermissoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private colPermissoes As Collection

Public Function CarregaMenu(objErr As Object, objTreeView As TreeView) As Boolean
   On Error GoTo CarregaMenu_E
   
   Dim nodMn As Node
   Dim strMn As String
   Dim intMn As Integer
   Dim objtmp As Object
   Dim intMnAux As Integer
   Dim intSubMn As Integer
   Dim strCaption As String
   Dim clsCursor As INF_Cursor.Cursor
   Dim blnComando As Boolean
   Dim intComandMn As Integer
   Dim clsContainer As sContainerPermissoes
   
   CarregaMenu = False
   
   f1.CollectionLimpar colPermissoes
   
   For Each objtmp In MDIInfinity
      strMn = Replace(objtmp.Name, "ID", "")
      strMn = Replace(strMn, "_", "")
      If Right(strMn, 1) = "C" Then blnComando = True
      strMn = Replace(strMn, "C", "")
      
      If Not strMn = 0 Then
         intMn = Left(strMn, 2)
         intSubMn = Mid(strMn, 3, 2)
         intComandMn = Right(strMn, 2)
         strCaption = Replace(objtmp.Caption, "&", "")
         
         If intMn <> intMnAux Then
            intMnAux = intMn
            Set nodMn = objTreeView.Nodes.Add(, , "ID_" & intMn, strCaption)
         End If
            
         If intComandMn > 0 Then
            Set nodMn = objTreeView.Nodes.Add("ID_" & intMn & intSubMn, tvwChild, "ID_" & intMn & intSubMn & intComandMn, strCaption)
         ElseIf intSubMn > 0 Then
            Set nodMn = objTreeView.Nodes.Add("ID_" & intMn, tvwChild, "ID_" & intMn & intSubMn, strCaption)
         End If
      
         nodMn.Selected = False
      End If
      
      If blnComando Then
         nodMn.Tag = strMn
         blnComando = False
         
         Set clsCursor = New INF_Cursor.Cursor
         With clsCursor
            .Inicializar clsConexao
            
            .SQL.Limpar
            .SQL.Mais " SELECT GlbInterface.Codigo AS IdInterface, GlbButton.Codigo AS IdButton, GlbButton.Descricao, "
            .SQL.Mais "    GlbInterface.Descricao AS DscInterface "
            .SQL.Mais " FROM GlbInterface "
            .SQL.Mais " INNER JOIN GlbButton ON (GlbInterface.IdButton = GlbButton.Codigo)"
            .SQL.Mais " WHERE GlbInterface.Codigo = " & .Vlr(strMn)
            
            If Not .Abrir(.SQL.Texto) Then
               objErr.Transferir = .TransferirErro
               objErr.ModRotina = "LoadMenu"
               GoTo DestruirObjetos
            End If
            
            Do Until .EOF
               Set clsContainer = New sContainerPermissoes
               clsContainer.IDInterface = .Valor("IdInterface")
               clsContainer.IDButton = .Valor("IdButton")
               
               'Se o botão for 99 = Relatórios, coloca o nome da interface na descrição do evento a ser ativado
               If .Valor("IdButton") = 99 Then
                  clsContainer.DescButton = .Valor("DscInterface")
               Else
                  clsContainer.DescButton = .Valor("Descricao")
               End If
               
               clsContainer.Selecionado = False
               colPermissoes.Add clsContainer, clsContainer.Key
               
               .ProximoRegistro
            Loop
            .Fechar
         End With
         Set clsCursor = Nothing
      End If
   Next
   
   CarregaMenu = True
   
   GoTo DestruirObjetos

CarregaMenu_E:
   objErr.Salvar Err
   objErr.ModRotina = "CarregaMenu"
   
DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Function

Public Function CarregaPermissoesUsuario(objErr As Object, CodUsuario As Long) As Boolean
   On Error GoTo CarregaPermissoesUsuario_E
   
   Dim clsCursor As INF_Cursor.Cursor
   Dim clsContainer As sContainerPermissoes
   
   CarregaPermissoesUsuario = False
   
   Set clsCursor = New INF_Cursor.Cursor
   With clsCursor
      .Inicializar clsConexao
      
      .SQL.Limpar
      .SQL.Mais " SELECT Empresa, IdUsuario, IdInterface, IdButton "
      .SQL.Mais " FROM GLBPermissoes "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND IdUsuario = " & .Vlr(CodUsuario)
      
      If Not .Abrir(.SQL.Texto) Then
         objErr.Transferir = .TransferirErro
         objErr.ModRotina = "CarregaPermissoesUsuario"
         GoTo DestruirObjetos
      End If
      
      Do Until .EOF
         For Each clsContainer In colPermissoes
            If CCur(clsContainer.IDInterface) = .Valor("IdInterface") And clsContainer.IDButton = .Valor("idbutton") Then
               clsContainer.Selecionado = True
            End If
         Next
         .ProximoRegistro
      Loop
      .Fechar
   End With
   
   CarregaPermissoesUsuario = True
   
   GoTo DestruirObjetos

CarregaPermissoesUsuario_E:
   objErr.Salvar Err
   objErr.ModRotina = "CarregaPermissoesUsuario"
   
DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Function

Public Property Get CollectionPermissoes() As Collection
   Set CollectionPermissoes = colPermissoes
End Property

Public Property Let CollectionPermissoes(ByVal vNewValue As Collection)
   Set colPermissoes = vNewValue
End Property

Private Sub Class_Initialize()
   Set colPermissoes = New Collection
End Sub

Private Sub Class_Terminate()
   Set colPermissoes = Nothing
End Sub
