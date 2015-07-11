VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "InfinityControl.ocx"
Begin VB.Form frmCG 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consulta Geral"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7020
   Icon            =   "CG.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3540
      Top             =   4725
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPesquisar 
      Caption         =   "&Pesquisar"
      Height          =   345
      Left            =   4230
      TabIndex        =   2
      Top             =   4935
      Width           =   1320
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   5610
      TabIndex        =   3
      Top             =   4935
      Width           =   1320
   End
   Begin VB.Frame fraParametros 
      Caption         =   "Parâmetros"
      Height          =   1095
      Left            =   90
      TabIndex        =   4
      Top             =   3795
      Width           =   6825
      Begin rdActiveText.ActiveText txtDesc 
         Height          =   315
         Left            =   960
         TabIndex        =   7
         Top             =   630
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   40
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText vlrCod 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   1380
         _ExtentX        =   2434
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
      Begin VB.Label lblDesc 
         Caption         =   "Descrição:"
         Height          =   225
         Left            =   120
         TabIndex        =   6
         Top             =   705
         Width           =   1290
      End
      Begin VB.Label lblCod 
         Caption         =   "Código:"
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   270
         Width           =   1290
      End
   End
   Begin MSComctlLib.ListView lstConsulta 
      Height          =   3660
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   6456
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "vlr"
         Text            =   "Código"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "txt"
         Text            =   "Descrição"
         Object.Width           =   6068
      EndProperty
   End
End
Attribute VB_Name = "frmCG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private clsErro As INF_Erro.Funcoes
Private colCG As Collection

Private Const LST_COL_CODIGO As Integer = 0
Private Const LST_COL_DESCRICAO As Integer = 1

Private Const LST_ICO_GRAVADO As String = "Gravado"

Private curCodigo As Currency
Private strDescricao As String
Private intDefinicao As Integer
Private blnCancelado As Boolean
Private blnAtivo As Boolean

Public Enum DefineConsulta
   enUsuarios = 1
   enClientes = 2
   enMarcas = 3
   enUnidades = 4
   enFornecedores = 5
   enModelos = 6
   enGrupos = 7
   enSubGrupos = 8
   
   'Acima do 100, são tabelas que não possuem o campo empresa "Tabelas Globais"
   enMunicipios = 100
End Enum

Private Sub cmdCancelar_Click()
   blnCancelado = True
   Me.Hide
End Sub

Private Sub cmdPesquisar_Click()
   On Error GoTo cmdPesquisar_Click_E
   
   Dim strAux As String
   Dim itemX As ListItem
   Dim strCod As String
   Dim strDesc As String
   Dim strTable As String
   Dim clsCursor As INF_Cursor.Cursor
   Dim clsContainer As sContainerCG
   
   f1.CollectionLimpar colCG
   Me.lstConsulta.ListItems.Clear
   
   Set clsCursor = New INF_Cursor.Cursor
   With clsCursor
      .Inicializar clsConexao
      
      .SQL.Limpar
      .SQL.Mais " SELECT "
      
      If intDefinicao < 100 Then
      
         Select Case intDefinicao
         Case enUsuarios
            strCod = "Codigo"
            strDesc = "Nome"
            strTable = "Usuarios"
            
         Case enClientes
            strCod = "Codigo"
            strDesc = "Nome"
            strTable = "Clientes"
         
         Case enMarcas
            strCod = "Codigo"
            strDesc = "Descricao"
            strTable = "Marcas"
         
         Case enUnidades
            strCod = "Codigo"
            strDesc = "Descricao"
            strTable = "Unidades"
   
         Case enFornecedores
            strCod = "Codigo"
            strDesc = "RazaoSocial"
            strTable = "Fornecedores"
         
         Case enModelos
            strCod = "Codigo"
            strDesc = "Descricao"
            strTable = "Modelos"
         
         Case enGrupos
            strCod = "Codigo"
            strDesc = "Descricao"
            strTable = "Grupos"
         
         Case enSubGrupos
            strCod = "Codigo"
            strDesc = "Descricao"
            strTable = "SubGrupos"
         
         End Select
      
         .SQL.Mais strCod & " AS Cod, " & strDesc & " AS Dsc "
         .SQL.Mais " FROM " & strTable
         .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
         
         If Len(Trim(Me.vlrCod)) > 0 And Me.vlrCod <> 0 Then .SQL.Mais " AND " & strCod & " = " & .Vlr(Me.vlrCod)
         If Len(Trim(Me.txtDesc)) > 0 Then .SQL.Mais " AND " & strDesc & " Like " & .Txt(Me.txtDesc & "%")
         If blnAtivo Then .SQL.Mais " AND Situacao = " & .Txt("1")
      Else
      
         strAux = ""
         
         Select Case intDefinicao
         Case enMunicipios
            strCod = "Codigo"
            strDesc = "Municipio"
            strTable = "Municipios"
         
         End Select
         
         .SQL.Mais strCod & " AS Cod, " & strDesc & " AS Dsc "
         .SQL.Mais " FROM " & strTable
         
         If Len(Trim(Me.vlrCod)) > 0 And Me.vlrCod <> 0 Then strAux = " WHERE " & strCod & " = " & .Vlr(Me.vlrCod)
         If Len(Trim(Me.txtDesc)) > 0 Then
            If Len(Trim(strAux)) > 0 Then
               If Len(Trim(Me.txtDesc)) > 0 Then strAux = strAux & " AND " & strDesc & " Like " & .Txt("%" & Me.txtDesc & "%")
            Else
               If Len(Trim(Me.txtDesc)) > 0 Then strAux = strAux & " WHERE " & strDesc & " Like " & .Txt("%" & Me.txtDesc & "%")
            End If
         End If
         
         If blnAtivo Then
            If Len(Trim(strAux)) > 0 Then
               strAux = strAux & " AND Ativo = " & .Txt("S")
            Else
               strAux = strAux & " WHERE Ativo = " & .Txt("S")
            End If
         End If
         
         If Len(Trim(strAux)) > 0 Then .SQL.Mais strAux
      End If
      
      If Not .Abrir(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         Exibir clsErro, "cmdPesquisar_Click"
         GoTo DestruirObjetos
      End If
      
      Do Until .EOF
         Set clsContainer = New sContainerCG
         clsContainer.Codigo = .Valor("Cod")
         clsContainer.Descricao = .Valor("Dsc")
         colCG.Add clsContainer, clsContainer.Key
         
         Set itemX = Me.lstConsulta.ListItems.Add(, clsContainer.Key, clsContainer.Codigo)
         itemX.SubItems(LST_COL_DESCRICAO) = clsContainer.Descricao
         itemX.Selected = True
      
         .ProximoRegistro
      Loop
      .Fechar
   End With
   
   GoTo DestruirObjetos

cmdPesquisar_Click_E:
   clsErro.Salvar Err
   Exibir clsErro, "cmdPesquisar_Click"

DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
   Set itemX = Nothing
   Set clsContainer = Nothing
End Sub

Private Sub Form_Initialize()
   Set colCG = New Collection
   Set clsErro = New INF_Erro.Funcoes
End Sub

Private Sub Form_Terminate()
   f1.CollectionLimpar colCG
   Set colCG = Nothing
   Set clsErro = Nothing
End Sub

Private Sub lstConsulta_DblClick()
   On Error GoTo lstConsultar_DblClick_E
   
   Dim itemX As ListItem
   
   Set itemX = Me.lstConsulta.SelectedItem
   If itemX Is Nothing Then Exit Sub
   If Not itemX.Selected Then Exit Sub
   
   Codigo = itemX.Text
   Descricao = itemX.SubItems(LST_COL_DESCRICAO)
   
   blnCancelado = False
   Me.Hide
   
   Exit Sub
   
lstConsultar_DblClick_E:
   clsErro.Salvar Err
   Exibir clsErro, "lstConsulta_DblClick"
End Sub

Public Property Get Codigo() As Currency
   Codigo = curCodigo
End Property

Public Property Let Codigo(ByVal vNewValue As Currency)
   curCodigo = vNewValue
End Property

Public Property Get Descricao() As String
   Descricao = strDescricao
End Property

Public Property Let Descricao(ByVal vNewValue As String)
   strDescricao = vNewValue
End Property

Public Property Let TpDefinicao(ByVal vNewValue As Integer)
   intDefinicao = vNewValue
End Property

Public Property Get Cancelado() As Boolean
   Cancelado = blnCancelado
End Property

Public Property Let Ativo(ByVal vNewValue As Boolean)
   blnAtivo = vNewValue
End Property
