VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "InfinityControl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   9285
   Tag             =   "20402"
   Begin VB.CommandButton cmdMapa 
      Height          =   810
      Left            =   7905
      Picture         =   "Cliente.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   53
      TabStop         =   0   'False
      ToolTipText     =   "Visualizar localização."
      Top             =   2340
      Width           =   1320
   End
   Begin VB.Frame fraContatos 
      Caption         =   "Contatos"
      Height          =   1785
      Left            =   60
      TabIndex        =   46
      Top             =   4380
      Width           =   7770
      Begin VB.CommandButton cmdAdicionar 
         Caption         =   "<<"
         Height          =   315
         Left            =   7275
         TabIndex        =   49
         ToolTipText     =   "Adicionar Fornecedor."
         Top             =   240
         Width           =   345
      End
      Begin VB.CommandButton cmdRemover 
         Caption         =   ">>"
         Height          =   315
         Left            =   7275
         TabIndex        =   48
         ToolTipText     =   "Remover Fornecedor."
         Top             =   930
         Width           =   345
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "="
         Height          =   315
         Left            =   7275
         TabIndex        =   47
         ToolTipText     =   "Adicionar Fornecedor."
         Top             =   585
         Width           =   345
      End
      Begin MSComctlLib.ListView lstContatos 
         Height          =   1440
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   7065
         _ExtentX        =   12462
         _ExtentY        =   2540
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList"
         SmallIcons      =   "ImageList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Contato"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Telefone"
            Object.Width           =   2647
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "E-mail"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Novo"
      Enabled         =   0   'False
      Height          =   345
      Left            =   7905
      TabIndex        =   33
      Top             =   90
      Width           =   1320
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "&Consultar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   7905
      TabIndex        =   32
      Top             =   465
      Width           =   1320
   End
   Begin VB.CommandButton cmdSalvar 
      Cancel          =   -1  'True
      Caption         =   "&Salvar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   7905
      TabIndex        =   31
      Tag             =   "20102"
      Top             =   840
      Width           =   1320
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Enabled         =   0   'False
      Height          =   345
      Left            =   7905
      TabIndex        =   30
      Tag             =   "6"
      Top             =   1215
      Width           =   1320
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "&Limpar"
      Height          =   345
      Left            =   7905
      TabIndex        =   29
      Top             =   1590
      Width           =   1320
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   345
      Left            =   7905
      TabIndex        =   28
      Top             =   1965
      Width           =   1320
   End
   Begin VB.Frame Frame3 
      Caption         =   "Parâmetros"
      Height          =   3720
      Left            =   60
      TabIndex        =   1
      Top             =   660
      Width           =   7755
      Begin VB.ComboBox cmbTipoPessoa 
         Height          =   315
         Left            =   3735
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   1005
         Width           =   1650
      End
      Begin VB.ComboBox cmbSituacao 
         Height          =   315
         Left            =   6585
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   1005
         Width           =   1020
      End
      Begin VB.CommandButton cmdConsultarCEP 
         Height          =   315
         Left            =   2205
         Picture         =   "Cliente.frx":5ADA
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Consultar CEP"
         Top             =   1005
         Width           =   330
      End
      Begin VB.ComboBox cmbUf 
         Height          =   315
         Left            =   6585
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2145
         Width           =   1020
      End
      Begin VB.CommandButton cmdPesqCidade 
         Caption         =   "..."
         Height          =   255
         Left            =   2445
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1830
         Width           =   270
      End
      Begin rdActiveText.ActiveText vlrCEP 
         Height          =   315
         Left            =   1005
         TabIndex        =   2
         Top             =   1005
         Width           =   1125
         _ExtentX        =   1984
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
         MaxLength       =   9
         TextMask        =   6
         RawText         =   6
         Mask            =   "#####-###"
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtBairro 
         Height          =   315
         Left            =   1005
         TabIndex        =   3
         Top             =   2145
         Width           =   4335
         _ExtentX        =   7646
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtDescCidade 
         Height          =   315
         Left            =   2775
         TabIndex        =   4
         Top             =   1770
         Width           =   4830
         _ExtentX        =   8520
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
      Begin rdActiveText.ActiveText vlrCodCidade 
         Height          =   315
         Left            =   1005
         TabIndex        =   5
         Top             =   1770
         Width           =   1125
         _ExtentX        =   1984
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
         Text            =   "0"
         TextMask        =   3
         RawText         =   3
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText vlrNro 
         Height          =   315
         Left            =   6585
         TabIndex        =   6
         Top             =   1395
         Width           =   1020
         _ExtentX        =   1799
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
         Text            =   "0"
         TextMask        =   3
         RawText         =   3
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtEndereco 
         Height          =   315
         Left            =   1005
         TabIndex        =   10
         Top             =   1395
         Width           =   4350
         _ExtentX        =   7673
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText datCadastro 
         Height          =   315
         Left            =   6585
         TabIndex        =   20
         Top             =   240
         Width           =   1020
         _ExtentX        =   1799
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
         MaxLength       =   10
         TextMask        =   1
         RawText         =   1
         Mask            =   "##/##/####"
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
         Locked          =   -1  'True
      End
      Begin rdActiveText.ActiveText vlrCPFCNPJ 
         Height          =   315
         Left            =   1005
         TabIndex        =   21
         Top             =   615
         Width           =   2700
         _ExtentX        =   4763
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
         TextMask        =   9
         RawText         =   9
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtNome 
         Height          =   315
         Left            =   1005
         TabIndex        =   22
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText vlrIE 
         Height          =   315
         Left            =   4905
         TabIndex        =   23
         Top             =   615
         Width           =   2700
         _ExtentX        =   4763
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
         TextMask        =   9
         RawText         =   9
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtEmail 
         Height          =   315
         Left            =   1005
         TabIndex        =   36
         Top             =   2895
         Width           =   6600
         _ExtentX        =   11642
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtTelefone 
         Height          =   315
         Left            =   1005
         TabIndex        =   37
         Top             =   2505
         Width           =   1785
         _ExtentX        =   3149
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtFax 
         Height          =   315
         Left            =   3585
         TabIndex        =   38
         Top             =   2505
         Width           =   1755
         _ExtentX        =   3096
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtCelular 
         Height          =   315
         Left            =   5850
         TabIndex        =   39
         Top             =   2520
         Width           =   1755
         _ExtentX        =   3096
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
         MaxLength       =   14
         TextMask        =   5
         RawText         =   5
         Mask            =   "(###)####-####"
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtSite 
         Height          =   315
         Left            =   1005
         TabIndex        =   40
         Top             =   3270
         Width           =   6600
         _ExtentX        =   11642
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin VB.Label lblTipoPessoa 
         Caption         =   "Tp. Pessoa:"
         Height          =   255
         Left            =   2745
         TabIndex        =   55
         Top             =   1035
         Width           =   1230
      End
      Begin VB.Label lblTelefone 
         Caption         =   "Telefone:"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   2535
         Width           =   1125
      End
      Begin VB.Label lblEmail 
         Caption         =   "E-mail:"
         Height          =   210
         Left            =   120
         TabIndex        =   44
         Top             =   2925
         Width           =   1095
      End
      Begin VB.Label lblSite 
         Caption         =   "Site:"
         Height          =   240
         Left            =   120
         TabIndex        =   43
         Top             =   3300
         Width           =   1110
      End
      Begin VB.Label lblFax 
         Caption         =   "Fax:"
         Height          =   195
         Left            =   3030
         TabIndex        =   42
         Top             =   2535
         Width           =   885
      End
      Begin VB.Label lblCelular 
         Caption         =   "Cel.:"
         Height          =   195
         Left            =   5445
         TabIndex        =   41
         Top             =   2550
         Width           =   885
      End
      Begin VB.Label lblSituação 
         Caption         =   "Situação:"
         Height          =   240
         Left            =   5445
         TabIndex        =   35
         Top             =   1035
         Width           =   1050
      End
      Begin VB.Label lblNome 
         Caption         =   "Nome:"
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Top             =   270
         Width           =   1110
      End
      Begin VB.Label lblCNPJ 
         Caption         =   "CNPJ:"
         Height          =   225
         Left            =   120
         TabIndex        =   26
         Top             =   645
         Width           =   1125
      End
      Begin VB.Label lblIE 
         Caption         =   "Ins. Estadual:"
         Height          =   225
         Left            =   3810
         TabIndex        =   25
         Top             =   645
         Width           =   1080
      End
      Begin VB.Label lblDataCadastro 
         Caption         =   "Dt. Cadastro:"
         Height          =   180
         Left            =   5445
         TabIndex        =   24
         Top             =   270
         Width           =   1395
      End
      Begin VB.Label lblCEP 
         Caption         =   "CEP"
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   1050
         Width           =   1200
      End
      Begin VB.Label lblEndereco 
         Caption         =   "Endereço:"
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Top             =   1425
         Width           =   1155
      End
      Begin VB.Label lblBairro 
         Caption         =   "Bairro:"
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   2175
         Width           =   1140
      End
      Begin VB.Label lblNumero 
         Caption         =   "Número:"
         Height          =   270
         Left            =   5445
         TabIndex        =   13
         Top             =   1440
         Width           =   1140
      End
      Begin VB.Label lblCidade 
         Caption         =   "Cidade:"
         Height          =   300
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   1170
      End
      Begin VB.Label lblUF 
         Caption         =   "Estado:"
         Height          =   225
         Left            =   5445
         TabIndex        =   11
         Top             =   2175
         Width           =   990
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Identificação"
      Height          =   660
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   7755
      Begin VB.CommandButton cmdPesqForn 
         Caption         =   "..."
         Height          =   255
         Left            =   2070
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   315
         Width           =   270
      End
      Begin rdActiveText.ActiveText vlrCod 
         Height          =   315
         Left            =   855
         TabIndex        =   18
         Top             =   240
         Width           =   1140
         _ExtentX        =   2011
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
         Height          =   270
         Left            =   120
         TabIndex        =   19
         Top             =   270
         Width           =   1065
      End
   End
   Begin MSComctlLib.ListView lstLegenda 
      Height          =   1065
      Left            =   7905
      TabIndex        =   51
      Top             =   5100
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   1879
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList"
      SmallIcons      =   "ImageList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Descricao"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Label lblLegenda 
      Caption         =   "Legenda"
      Height          =   225
      Left            =   7905
      TabIndex        =   52
      Top             =   4875
      Width           =   1215
   End
End
Attribute VB_Name = "frmCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private colContatos As Collection
Private clsErro As INF_Erro.Funcoes

Private Sub cmdConsultar_Click()
   On Error GoTo cmdConsultar_Click_E

   Dim clsCursor As INF_Cursor.Cursor

   If Not mCmpObrigatorio(clsErro, Me.vlrCod, "Código") Then GoTo Erro_Msg
   
   Set clsCursor = CreateObject("INF_Cursor.Cursor")
   With clsCursor
      .Inicializar clsConexao
      
      .SQL.Limpar
      .SQL.Mais " SELECT Empresa, Codigo, Nome, DataCad, CPFCNPJ, RGIE, CEP, TpPessoa, "
      .SQL.Mais "    Situacao, Endereco, Numero, CodCidade, Bairro, Estado, Telefone, "
      .SQL.Mais "    Fax, Cel, Email, Site "
      .SQL.Mais " FROM Clientes "
      .SQL.Mais " WHERE Empresa = " & .Txt(Prj.Sistema.IdEmpresa)
      .SQL.Mais " AND Codigo = " & .Vlr(Me.vlrCod)
      
      If Not .Abrir(.SQL.Texto) Then
         clsErro.Transferir = .TransferirErro
         Exibir clsErro, "cmdConsultar_Click"
         GoTo DestruirObjetos
      End If
      
      If Not .EOF Then
         
      End If
   End With
      
   With clsCursor
      .SQL.Limpar
      .SQL.Mais " SELECT "
   End With
   GoTo DestruirObjetos
   
cmdConsultar_Click_E:
   clsErro.Salvar Err
   Exibir clsErro, "cmdConsultar_Click"
   GoTo DestruirObjetos
   
Erro_Msg:
   Exibir clsErro, "cmdConsultar_Click"
   
DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
End Sub

Private Sub cmdConsultarCEP_Click()
   On Error GoTo cmdConsultarCEP_Click_E
   
   Dim strEndereco As String
   Dim url As String
   Dim clsCursor As INF_Cursor.Cursor
   Dim objXml As MSXML2.DOMDocument
   Dim objXmlNodes As MSXML2.IXMLDOMElement
   Dim ObjXmlElement As MSXML2.IXMLDOMElement
   
   If Not f1.VerificaConexaoInternet(clsErro) Then
      Exibir clsErro, "cmdConsultarCEP_Click"
      GoTo DestruirObjetos
   End If
   
   Set objXml = New MSXML2.DOMDocument
   
   url = "http://republicavirtual.com.br/web_cep.php?cep=" & Replace(Me.vlrCEP, "-", "") & "&formato=XML"
   objXml.async = False
   
   Call objXml.Load(url)

   If objXml.parseError.reason <> "" Then
      MsgBox objXml.parseError.reason
      Exit Sub
   End If

   Set objXmlNodes = objXml.documentElement

   For Each ObjXmlElement In objXmlNodes.childNodes
      Select Case ObjXmlElement.nodeName
      
      Case "resultado"
         If ObjXmlElement.nodeTypedValue = 0 Then
            mMsgInfo "CEP inválido! Verifique."
            mFocus Me.vlrCEP
            Exit Sub
         End If
      Case "cidade"
         
         Set clsCursor = CreateObject("INF_Cursor.Cursor")
         With clsCursor
            .Inicializar clsConexao
            
            .SQL.Limpar
            .SQL.Mais " SELECT Codigo, Municipio, UF "
            .SQL.Mais " FROM Municipios "
            .SQL.Mais " WHERE UPPER(Municipio) = " & .Txt(UCase(ObjXmlElement.nodeTypedValue))
            
            If Not .Abrir(.SQL.Texto) Then
               clsErro.Transferir = .TransferirErro
               Exibir clsErro, "cmdConsultarCEP_Click"
               GoTo DestruirObjetos
            End If
            
            If Not .EOF Then
               Me.vlrCodCidade = .Valor("Codigo")
               Me.txtDescCidade = .Valor("Municipio")
               Me.cmbUf.ListIndex = f1.CmbValor(Me.cmbUf, .Valor("UF"), enDescricao, 2)
            End If
            .Fechar
         End With
         
      Case "bairro"
         Me.txtBairro = ObjXmlElement.nodeTypedValue
         
      Case "tipo_logradouro"
         Me.txtEndereco = ObjXmlElement.nodeTypedValue
         
      Case "logradouro"
         Me.txtEndereco = Me.txtEndereco & " " & ObjXmlElement.nodeTypedValue
      End Select
   Next
   
   mFocus Me.vlrNro
   
   GoTo DestruirObjetos
   
cmdConsultarCEP_Click_E:
   clsErro.Salvar Err
   Exibir clsErro, "cmdConsultarCEP_Click"

DestruirObjetos:
   If Not (clsCursor Is Nothing) Then clsCursor.Fechar
   Set clsCursor = Nothing
   Set objXml = Nothing
   Set objXmlNodes = Nothing
   Set ObjXmlElement = Nothing
End Sub

Private Sub cmdMapa_Click()
   Dim frmModal As frmMapa
   
   If Len(Trim(Me.txtEndereco)) > 0 Then
      Set frmModal = New frmMapa
      frmModal.Endereco = Me.txtEndereco
      frmModal.LocalizarEndereco
      frmModal.Show vbModal
   End If
End Sub

Private Sub Form_Load()
   Set clsErro = New INF_Erro.Funcoes
   
   f1.FormCentralizar Me
End Sub
