VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CmdSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private strCmdSQL As String

Public Sub Limpar()
   strCmdSQL = ""
End Sub

Public Property Get Tamanho() As Long
   Tamanho = Len(strCmdSQL)
End Property

Public Sub Mais(Texto As String, Optional ComVirgula As Boolean = False)
   strCmdSQL = strCmdSQL & Texto & IIf(ComVirgula, ",", "")
End Sub

Public Property Get Texto() As String
   Texto = strCmdSQL
End Property

