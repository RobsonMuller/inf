VERSION 5.00
Begin VB.MDIForm MDIInfinity 
   BackColor       =   &H8000000C&
   Caption         =   "Infinity Cong ..."
   ClientHeight    =   5985
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   9975
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu ID_01_00_00 
      Caption         =   "&Iniciar"
      Begin VB.Menu ID_01_01_00 
         Caption         =   "&Usuario"
         Begin VB.Menu ID_01_01_01_C 
            Caption         =   "&Cadastro"
         End
         Begin VB.Menu ID_01_01_02_C 
            Caption         =   "&Permissões"
         End
         Begin VB.Menu ID_01_01_03_C 
            Caption         =   "&Alterar Senha"
         End
      End
      Begin VB.Menu ID_00_00_00 
         Caption         =   "-"
      End
      Begin VB.Menu ID_01_02_00_C 
         Caption         =   "&Conectar"
      End
      Begin VB.Menu ID_01_03_00 
         Caption         =   "&Encerrar"
      End
   End
   Begin VB.Menu ID_02_00_00 
      Caption         =   "&Lançamentos"
      Begin VB.Menu ID_02_01_00 
         Caption         =   "&Sub_Cadastros"
         Begin VB.Menu ID_02_01_01_C 
            Caption         =   "&Marcas"
         End
         Begin VB.Menu ID_02_01_02_C 
            Caption         =   "&Modelos"
         End
         Begin VB.Menu ID_02_01_03_C 
            Caption         =   "&Unidades"
         End
         Begin VB.Menu ID_02_01_04_C 
            Caption         =   "&Grupo"
         End
         Begin VB.Menu ID_02_01_05_C 
            Caption         =   "&Sub_Grupo"
         End
      End
      Begin VB.Menu ID_02_02_00_C 
         Caption         =   "&Produtos"
      End
      Begin VB.Menu ID_02_03_00 
         Caption         =   "&Estoque"
         Begin VB.Menu ID_02_03_01_C 
            Caption         =   "&Entrada"
         End
         Begin VB.Menu ID_02_03_02_C 
            Caption         =   "&Saida"
         End
         Begin VB.Menu ID_02_03_03_C 
            Caption         =   "&Geral"
         End
      End
      Begin VB.Menu ID_02_04_00 
         Caption         =   "&Pessoas"
         Begin VB.Menu ID_02_04_01_C 
            Caption         =   "&Fornecedores"
         End
         Begin VB.Menu ID_02_04_02_C 
            Caption         =   "&Clientes"
         End
         Begin VB.Menu ID_02_04_03_C 
            Caption         =   "&Empresas"
         End
      End
      Begin VB.Menu ID_02_05_00 
         Caption         =   "&Correio"
         Begin VB.Menu ID_02_05_01_C 
            Caption         =   "&Envio"
         End
         Begin VB.Menu ID_02_05_02_C 
            Caption         =   "&Recebimento"
         End
      End
   End
   Begin VB.Menu ID_03_00_00 
      Caption         =   "&Relatórios"
      Begin VB.Menu ID_03_01_00 
         Caption         =   "&Sub_Cadastros"
         Begin VB.Menu ID_03_01_01_C 
            Caption         =   "&Marcas"
         End
      End
   End
   Begin VB.Menu ID_04_00_00 
      Caption         =   "&Parâmetros"
      Begin VB.Menu ID_04_01_00_C 
         Caption         =   "&Configuração de E-mail"
      End
   End
End
Attribute VB_Name = "MDIInfinity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsErro As INF_Erro.Funcoes

Private Sub ID_1_01_01_Click()
   End Sub

Private Sub ID_01_01_01_C_Click()
   frmUsuarios.Show
End Sub

Private Sub ID_01_01_02_C_Click()
   frmPermissoes.Show
End Sub

Public Sub CarregaMenu()
   On Error GoTo CarregaMenu_E
   
   Dim objMn As Object
   Dim strCodMn As String
   Dim clsCursor As INF_Cursor.Cursor
   
   Set clsCursor = New INF_Cursor.Cursor
   With clsCursor
      .Inicializar clsConexao
      
      .SQL.Limpar
      .SQL.Mais " SELECT Empresa, IdUsuario, IdInterface "
      .SQL.Mais " FROM GlbPermissoes "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND IdUsuario = " & .Vlr(Prj.Sistema.IdUsuario)
      .SQL.Mais " GROUP BY Empresa, IdUsuario, IdInterface "
      
      If Not .Abrir(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         Exibir clsErro, "CarregaMenu"
         GoTo DestruirObjetos
      End If
      
      For Each objMn In MDIInfinity
         strCodMn = ""
         strCodMn = Replace(objMn.Name, "ID_", "")
         strCodMn = Replace(strCodMn, "_", "")
         strCodMn = Replace(strCodMn, "C", "")
         If strCodMn > 0 Then objMn.Enabled = False
      Next
      
      'Menu Especiais
      ObjMenu "010000" 'Iniciar
      ObjMenu "010200" 'Conectar
      ObjMenu "010300" 'Encerrar
               
      If Prj.Sistema.IdUsuario = 1 Then
         ObjMenu "010100" 'Sub Grupo (Usuarios)
         ObjMenu "010102" 'Menu de permissões
      End If
      
      Do Until .EOF
         strCodMn = Format(.Valor("IdInterface"), "000000")
         ObjMenu strCodMn
         
         strCodMn = Left(strCodMn, 4) & "00"
         ObjMenu strCodMn
         
         strCodMn = Left(strCodMn, 2) & "0000"
         ObjMenu strCodMn
         
         .ProximoRegistro
      Loop
      
      .Fechar
   End With
   
   GoTo DestruirObjetos

CarregaMenu_E:
   clsErro.Salvar Err
   Exibir clsErro, "CarregaMenu"

DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Sub

Private Sub ID_01_01_03_C_Click()
   frmAlterarSenha.Show vbModal
End Sub

Private Sub ID_01_02_00_C_Click()
   frmLogin.Show vbModal
End Sub

Private Sub ID_01_03_00_Click()
   Unload Me
End Sub

Private Sub ID_02_01_01_C_Click()
   Dim frmModal As New frmBasicCad
   
   frmModal.BaseForm enBasicCadMarca
   frmModal.Show
End Sub

Private Sub ID_02_01_02_C_Click()
   Dim frmModal As New frmBasicCad
   
   frmModal.BaseForm enBasicCadModelo
   frmModal.Show
End Sub

Private Sub ID_02_01_03_C_Click()
   Dim frmModal As New frmBasicCad
   
   frmModal.BaseForm enBasicCadUnidade
   frmModal.Show
End Sub

Private Sub ID_02_01_04_C_Click()
   Dim frmModal As New frmBasicCad
   
   frmModal.BaseForm enBasicCadGrupo
   frmModal.Show
End Sub

Private Sub ID_02_01_05_C_Click()
   Dim frmModal As New frmBasicCad
   
   frmModal.BaseForm enBasicCadSubGrupo
   frmModal.Show
End Sub

Private Sub ID_02_02_00_C_Click()
   frmProdutos.Show
End Sub

Private Sub ID_02_02_01_C_Click()
   
End Sub

Private Sub ID_02_03_01_C_Click()
   frmEntrada.Show
End Sub

Private Sub ID_02_04_01_C_Click()
   frmFornecedores.Show
End Sub

Private Sub ID_02_04_02_C_Click()
   'frmCliente.Show
End Sub

Private Sub ID_02_05_01_C_Click()
    'frmEnvioEmail.Show
End Sub

Private Sub ID_02_05_02_C_Click()
   'frmRecebeEmail.Show
End Sub

Private Sub ID_03_01_01_C_Click()
   'frmRelMarcas.Show
End Sub

Private Sub ID_04_01_00_C_Click()
   'frmConfigEmail.Show
End Sub

Private Sub MDIForm_Initialize()
   Set clsErro = New INF_Erro.Funcoes
End Sub

Private Sub ObjMenu(strValor As String)
   Dim objMn As Object
   Dim strCodMn As String
   Dim intCount As Integer
   
   For Each objMn In MDIInfinity
      strCodMn = ""
      strCodMn = Replace(objMn.Name, "ID_", "")
      strCodMn = Replace(strCodMn, "_", "")
      strCodMn = Replace(strCodMn, "C", "")
            
      If strCodMn = strValor Then objMn.Enabled = True
   Next
End Sub

Private Sub MDIForm_Load()
   Call CarregaMenu
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   Set clsErro = Nothing
   Set clsConexao = Nothing
   If GetSetting(NM_APP, "Connect", "SavePassword") = "N" Then DeleteSetting NM_APP, "Connect", "Password"
End Sub
