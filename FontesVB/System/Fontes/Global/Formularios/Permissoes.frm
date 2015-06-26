VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "InfinityControl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPermissoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Permissões"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9780
   Icon            =   "Permissoes.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   9780
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8565
      Top             =   2970
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Permissoes.frx":000C
            Key             =   "Desmarcado"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Permissoes.frx":045E
            Key             =   "Marcado"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   345
      Left            =   8385
      TabIndex        =   9
      Top             =   1215
      Width           =   1320
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "&Limpar"
      Height          =   345
      Left            =   8385
      TabIndex        =   8
      Top             =   840
      Width           =   1320
   End
   Begin VB.CommandButton cmdSalvar 
      Cancel          =   -1  'True
      Caption         =   "&Salvar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   8385
      TabIndex        =   7
      Top             =   465
      Width           =   1320
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "&Consultar"
      Height          =   345
      Left            =   8385
      TabIndex        =   6
      Top             =   90
      Width           =   1320
   End
   Begin VB.Frame fraParametros 
      Caption         =   "Parâmetros"
      Height          =   4650
      Left            =   60
      TabIndex        =   11
      Top             =   645
      Width           =   8250
      Begin MSComctlLib.ListView lstComandos 
         Height          =   4290
         Left            =   4155
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   225
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   7567
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descrição"
            Object.Width           =   6068
         EndProperty
      End
      Begin MSComctlLib.TreeView treMenu 
         Height          =   4260
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   7514
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         Appearance      =   1
      End
   End
   Begin VB.Frame fraIdentificacao 
      Caption         =   "Identificação"
      Height          =   645
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   8235
      Begin rdActiveText.ActiveText txtNome 
         Height          =   315
         Left            =   2505
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   556
         BackColor       =   -2147483644
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
         Locked          =   -1  'True
      End
      Begin VB.CommandButton cmdPesq 
         Caption         =   "..."
         Height          =   255
         Left            =   2160
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   285
         Width           =   270
      End
      Begin rdActiveText.ActiveText vlrCod 
         Height          =   315
         Left            =   990
         TabIndex        =   1
         Top             =   240
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   10
         Text            =   "0"
         TextMask        =   3
         RawText         =   3
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin VB.Label lblCodigo 
         Caption         =   "Código:"
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   270
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmPermissoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsErro As INF_Erro.Funcoes
Private colPermissoes As Collection

Private Sub cmdConsultar_Click()
   On Error GoTo cmdConsultar_Click_E
   
   Dim clsCursor As INF_Cursor.Cursor
   Dim clsPermissoes As sPermissoes
   
  ' If Not mCmpObrigatorio(Me.vlrCod) Then Exit Sub
   
   f1.CollectionLimpar colPermissoes
   
   Set clsCursor = CreateObject("INF_Cursor.Cursor")
   With clsCursor
      .Inicializar clsConexao
      
      .SQL.Limpar
      .SQL.Mais " SELECT Empresa, Codigo, Usuario "
      .SQL.Mais " FROM Usuarios "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND Codigo = " & .Vlr(Me.vlrCod)
   
      If Not .Abrir(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         Exibir clsErro, "cmdConsultar_Click"
         GoTo DestruirObjetos
      End If
      
      If .EOF Then
         mMsgInfo "Usuário não localizado! Verifique."
         mFocus Me.vlrCod
         GoTo DestruirObjetos
      End If
      .Fechar
   End With
   
   Set clsPermissoes = New sPermissoes
   With clsPermissoes
      .CollectionPermissoes = colPermissoes
      
      If Not .CarregaMenu(clsErro, Me.treMenu) Then
         Exibir clsErro, "cmdConsultar_Click"
         GoTo DestruirObjetos
      End If
      
      If Not .CarregaPermissoesUsuario(clsErro, Me.vlrCod) Then
         Exibir clsErro, "cmdConsultar_Click"
         GoTo DestruirObjetos
      End If
      
      Set colPermissoes = .CollectionPermissoes
   End With
      
   Me.cmdConsultar.Enabled = False
   Me.cmdSalvar.Enabled = True
   
   Me.fraIdentificacao.Enabled = False
   Me.fraParametros.Enabled = True
   
   GoTo DestruirObjetos

cmdConsultar_Click_E:
   clsErro.Salvar Err
   Exibir clsErro, "cmdConsultar_Click"

DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub cmdLimpar_Click()
   Me.treMenu.Nodes.Clear
   Me.lstComandos.ListItems.Clear
   
   Me.vlrCod = 0
   Me.txtNome = ""
   Me.fraIdentificacao.Enabled = True
   Me.fraParametros.Enabled = False
   Me.cmdConsultar.Enabled = True
   Me.cmdSalvar.Enabled = False
   mFocus Me.vlrCod
End Sub

Private Sub cmdPesq_Click()
   Dim frmModal As frmCG
   
   Set frmModal = New frmCG
   With frmModal
      .Codigo = Me.vlrCod
      .TpDefinicao = enUsuarios
      .Show vbModal
      
      If Not .Cancelado Then Me.vlrCod = .Codigo
   End With
   
DestruirObjetos:
   If Not (frmModal Is Nothing) Then Unload frmModal
   Set frmModal = Nothing
End Sub

Private Sub cmdSalvar_Click()
   On Error GoTo cmdSalvar_Click_E
   
   Dim clsContainer As sContainerPermissoes
   
   With clsConexao
      .Begin
      
      .SQL.Limpar
      .SQL.Mais " DELETE FROM GlbPermissoes "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND IdUsuario = " & .Vlr(Me.vlrCod)
      
      If Not .Executar(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         .RollBack
         Exibir clsErro, "cmdSalvar_Click"
         Exit Sub
      End If
      
      For Each clsContainer In colPermissoes
         If clsContainer.Selecionado Then
            .SQL.Limpar
            .SQL.Mais " INSERT INTO GlbPermissoes ("
            .SQL.Mais "    Empresa, IdUsuario, IdInterface, IdButton "
            .SQL.Mais " ) VALUES ( "
            .SQL.Mais .Txt(Prj.Sistema.IdEmpresa, True)
            .SQL.Mais .Vlr(Me.vlrCod, True)
            .SQL.Mais .Vlr(clsContainer.IDInterface, True)
            .SQL.Mais .Vlr(clsContainer.IDButton)
            .SQL.Mais " )"
            
            If Not .Executar(.SQL.Texto) Then
               clsErro.Transferir = .TransferirErro
               .RollBack
               Exibir clsErro, "cmdSalvar_Click"
               Exit Sub
            End If
         End If
      Next
      .Commit
   End With
   
   'Atualiza o menu após a alteração
   Call MDIInfinity.CarregaMenu
   
   mMsgInfo "Registro salvo com sucesso!"
   cmdLimpar_Click
   
   Exit Sub

cmdSalvar_Click_E:
   clsErro.Salvar Err
   clsConexao.RollBack
   Exibir clsErro, "cmdSalvar_Click"
End Sub

Private Sub Form_Load()
   Set colPermissoes = New Collection
   Set clsErro = CreateObject("INF_Erro.Funcoes")
   
   f1.FormCentralizar Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set colPermissoes = Nothing
   Set clsErro = Nothing
End Sub

Private Sub lstComandos_DblClick()
   Dim itemX As ListItem
   Dim clsContainer As sContainerPermissoes
   
   Set itemX = Me.lstComandos.SelectedItem
   If itemX Is Nothing Then Exit Sub
   If Not itemX.Selected Then Exit Sub
   
   Set clsContainer = colPermissoes(itemX.Key)
   
   If Not clsContainer.Selecionado Then
      itemX.SmallIcon = "Marcado"
      clsContainer.Selecionado = True
   Else
      itemX.SmallIcon = "Desmarcado"
      clsContainer.Selecionado = False
   End If
End Sub

Private Sub treMenu_Click()
   On Error GoTo treMenu_Click_E
   
   Dim nodMn As Node
   Dim itemX As ListItem
   Dim clsContainer As sContainerPermissoes
   
   Set nodMn = Me.treMenu.SelectedItem
   If nodMn Is Nothing Then Exit Sub
   If Not nodMn.Selected Then Exit Sub
   
   Me.lstComandos.ListItems.Clear
   
   If Len(Trim(nodMn.Tag)) > 0 Then
      For Each clsContainer In colPermissoes
         If nodMn.Tag = Format(clsContainer.IDInterface, "000000") Then
            Set itemX = Me.lstComandos.ListItems.Add(, clsContainer.Key, clsContainer.DescButton)
            itemX.Selected = False
            If clsContainer.Selecionado Then
               itemX.SmallIcon = "Marcado"
            Else
               itemX.SmallIcon = "Desmarcado"
            End If
         End If
      Next
   End If
   
   Exit Sub

treMenu_Click_E:
   clsErro.Salvar Err
   Exibir clsErro, "treMenu_Click"
End Sub

Private Sub vlrCod_LostFocus()
   If Not VerificaUsuario(clsErro, Me.vlrCod, Me.txtNome) Then
      Exibir clsErro, "vlrCod_LostFocus"
      mFocus Me.vlrCod
   End If
End Sub
