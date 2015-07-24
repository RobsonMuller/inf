VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public SQL As New CmdSQL
Private strEmpresa As String
Private strNmReport As String
Private objFormReportCrystal As Object
Private clsErro As INF_Erro.Funcoes
Private clsConexao As INF_Conexao.Conexao

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Property Get TransferirErro() As Object
   Set TransferirErro = clsErro.Transferir
End Property

Public Function Inicializar(objConex As Object, Emp As String, NmRPT As String, objFormCrystal As Object) As Boolean
   On Error GoTo Inicializar_E
   
   Inicializar = False
   
   Set clsErro = CreateObject("INF_Erro.Funcoes")
   Set clsConexao = CreateObject("INF_Conexao.Conexao")
   If Not clsConexao.Abrir(objConex.ConnectionString) Then
      clsErro.Transferir = clsConexao.TransferirErro
      clsErro.ModRotina = "Inicializar"
      Exit Function
   End If
   Set objFormReportCrystal = objFormCrystal
      
   strEmpresa = Emp
   strNmReport = NmRPT
   
   Inicializar = True
   
   Exit Function

Inicializar_E:
   clsErro.Salvar Err
   clsErro.ModRotina = "Inicializar"
End Function

Public Function Imprimir() As Boolean
   On Error GoTo Imprimir_E
   
   'DLLs
   ' OLE DB - PDSOLEDB.DLL
   ' ODBC - PDSODBC.DLL
   
   Dim strSQL As String
   Dim ADOCon As ADODB.Connection
   Dim ADORst As ADODB.Recordset
   Dim ADOCommand As ADODB.Command
   
   Dim strPathReport As String
   Dim CryRpt As CRAXDRT.Report
   Dim CryApp As CRAXDRT.Application
   
   Dim strServer As String
   Dim strBase As String
   Dim strID As String
   Dim strPass As String
   
   Imprimir = False
   
   'Carrega Temporaria
   '--------------------------------------------------------------------------'
   Set ADOCon = New ADODB.Connection
   ADOCon.ConnectionString = clsConexao.ConnectionString
   ADOCon.Open
   
   ADOCon.BeginTrans
   strSQL = "IF (OBJECT_ID('dbo.VW_Rel') IS NOT NULL) DROP VIEW dbo.VW_Rel"
   
   Set ADOCommand = New ADODB.Command
   ADOCommand.ActiveConnection = ADOCon
   ADOCommand.Prepared = True
   ADOCommand.CommandText = strSQL
   ADOCommand.Execute
   
   strSQL = "CREATE VIEW VW_Rel AS " & SQL.Texto
   Set ADOCommand = New ADODB.Command
   ADOCommand.ActiveConnection = ADOCon
   ADOCommand.Prepared = True
   ADOCommand.CommandText = strSQL
   ADOCommand.Execute
   ADOCon.CommitTrans
      
   'Verifica caminho de relatórios
   strSQL = ""
   strSQL = strSQL & " SELECT PathReport "
   strSQL = strSQL & " FROM Empresas "
   strSQL = strSQL & " WHERE Codigo = '" & strEmpresa & "'"
   
   Set ADORst = New ADODB.Recordset
   ADORst.CursorLocation = adUseClient
   ADORst.CursorType = adOpenDynamic
   ADORst.ActiveConnection = ADOCon
   ADORst.Open strSQL
   
   If ADORst.EOF Then GoTo DestruirObjetos
   strPathReport = ADORst.Fields("PathReport")
   ADORst.Close
   
   Set CryApp = New CRAXDRT.Application
   Set CryRpt = CryApp.OpenReport("C:\Users\Robson\Desktop\_Desenv\Infinity\FontesVB\Cng\Relatorios\" & strNmReport, 1) '
   
   strID = GetSetting("Infinity", "Connect", "User")
   strPass = GetSetting("Infinity", "Connect", "Password")
   strBase = GetSetting("Infinity", "Connect", "Base")
   strServer = NmComp & "\" & GetSetting("Infinity", "Connect", "Server")
   
   CryRpt.Database.Tables(1).SetLogOnInfo strServer, strBase, strID, strPass
   CryRpt.DiscardSavedData
   
   objFormReportCrystal.CRViewer.ReportSource = CryRpt
   objFormReportCrystal.CRViewer.Zoom 100
   objFormReportCrystal.CRViewer.ViewReport
   objFormReportCrystal.Show
  
   Imprimir = True
   
   GoTo DestruirObjetos
   
Imprimir_E:
   clsErro.Salvar Err
   clsErro.ModRotina = "Imprimir"

DestruirObjetos:
   Set ADORst = Nothing
End Function

Private Sub Class_Terminate()
   Set objFormReportCrystal = Nothing
   Set clsErro = Nothing
   Set clsConexao = Nothing
   Set SQL = Nothing
End Sub

Public Property Get Vlr(Valor As Variant, Optional ComVirgula As Boolean = False) As String
   Vlr = Replace(CStr(Valor), ",", ".") & IIf(ComVirgula, ",", "")
End Property

Public Property Get Txt(Texto As String, Optional ComVirgula As Boolean = False) As String
   Txt = "'" & Texto & "'" & IIf(ComVirgula, ", ", "")
End Property

Public Property Get Dt(Data As Variant, Optional ComVirgula As Boolean = False, Optional ComHora As Boolean = False) As String
   If Not IsDate(Data) Then
      Dt = "CONVERT(datetime, NULL)"
   ElseIf IsDate(Data) Then
      If Year(Data) < 1900 Then
         Dt = "CONVERT(datetime, NULL)"
      Else
         If ComHora Then
            Dt = "{ts '" & Format$(Data, "yyyy-mm-dd hh:mm:ss") & "'}"
         Else
            Dt = "{d '" & Format(Data, "yyyy-mm-dd") & "'}"
         End If
      End If
   End If
   
   If ComVirgula Then Dt = Dt & ", "
End Property

'COMPABILIDADE COM O VISUAL BASIC 5
'VB6 JÁ POSSUI O REPLACE
Private Function Replace(Source As String, Find As String, ReplaceStr As String, _
    Optional ByVal Start As Long = 1, Optional Count As Long = -1, _
    Optional Compare As VbCompareMethod = vbBinaryCompare) As String

   Dim findLen As Long
   Dim replaceLen As Long
   Dim index As Long
   Dim counter As Long
   
   findLen = Len(Find)
   replaceLen = Len(ReplaceStr)
   If findLen = 0 Then Err.Raise 5
   
   If Start < 1 Then Start = 1
   index = Start
   Replace = Source
   Do
      index = InStr(index, Replace, Find, Compare)
      If index = 0 Then Exit Do
      If findLen = replaceLen Then
         Mid$(Replace, index, findLen) = ReplaceStr
      Else
         Replace = Left$(Replace, index - 1) & ReplaceStr & Mid$(Replace, index + findLen)
      End If
      index = index + replaceLen
      counter = counter + 1
   Loop Until counter = Count
   If Start > 1 Then Replace = Mid$(Replace, Start)
End Function

Public Function NmComp() As String
   Dim strNome As String
   Dim lretorno As Long

   strNome = String(255, Chr$(0))
   lretorno = GetComputerName(strNome, 255)
   strNome = Left$(strNome, InStr(1, strNome, Chr$(0)) - 1)
   NmComp = strNome
   Exit Function
End Function
