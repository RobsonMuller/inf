VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Funcoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private clsDeclarations As Declarations

Private Sub Class_Initialize()
   Set clsDeclarations = New Declarations
End Sub

Private Sub Class_Terminate()
   Set clsDeclarations = Nothing
End Sub

Public Sub Salvar(objErrVB As ErrObject, Optional NumPRJ As Long, Optional MsgPRJ As String, Optional CmdSQLPrj As String)
   'Project
   clsDeclarations.NroPrj = NumPRJ
   clsDeclarations.DscPrj = MsgPRJ
   clsDeclarations.CmdSQLPrj = CmdSQLPrj
   
   'Visual Basic
   clsDeclarations.NroVB = objErrVB.Number
   clsDeclarations.DscVB = objErrVB.Description
   
   clsDeclarations.ErrExiste = True
End Sub

Public Property Get ModRotina() As String
   ModRotina = clsDeclarations.ModRotPrj
End Property

Public Property Let ModRotina(ByVal vNewValue As String)
   If Len(clsDeclarations.ModRotPrj) > 0 Then clsDeclarations.ModRotPrj = clsDeclarations.ModRotPrj & "\"
   clsDeclarations.ModRotPrj = clsDeclarations.ModRotPrj & vNewValue
End Property

Public Property Get ErrExiste() As Boolean
   ErrExiste = clsDeclarations.ErrExiste
End Property

Public Property Let ErrExiste(ByVal vNewValue As Boolean)
   clsDeclarations.ErrExiste = vNewValue
   
   If Not clsDeclarations.ErrExiste Then
      'Project
      clsDeclarations.NroPrj = 0
      clsDeclarations.DscPrj = ""
      clsDeclarations.CmdSQLPrj = ""
      
      'Visual Basic
      clsDeclarations.NroVB = 0
      clsDeclarations.DscVB = ""
      Err.Clear
   End If
End Property

Public Property Get Transferir() As Declarations
   Set Transferir = clsDeclarations
End Property

Public Property Let Transferir(ErrDeclarations As Declarations)
   'Project
   clsDeclarations.NroPrj = ErrDeclarations.NroPrj
   clsDeclarations.DscPrj = ErrDeclarations.DscPrj
   clsDeclarations.ModRotPrj = ErrDeclarations.ModRotPrj
   clsDeclarations.CmdSQLPrj = ErrDeclarations.CmdSQLPrj
   
   'Visual Basic
   clsDeclarations.NroVB = ErrDeclarations.NroVB
   clsDeclarations.DscVB = ErrDeclarations.DscVB
   clsDeclarations.ErrExiste = ErrDeclarations.ErrExiste
   
   clsDeclarations.ErrExiste = False
End Property
