Public Class Erro
    Private clsDeclarations As New Declaration

    Public Sub Salvar(objErrVB As ErrObject, Optional NumPRJ As Long = 0, Optional MsgPRJ As String = "", Optional CmdSQLPrj As String = "")
        clsDeclarations.NroPrj = NumPRJ
        clsDeclarations.DscPrj = MsgPRJ
        clsDeclarations.CmdSQLPrj = CmdSQLPrj
        clsDeclarations.NroVB = objErrVB.Number
        clsDeclarations.DscVB = objErrVB.Description
        clsDeclarations.ErrExiste = True
    End Sub

    Public Property ModRotina() As String
        Get
            Return clsDeclarations.ModRotPrj
        End Get
        Set(vNewValue As String)
            If Len(clsDeclarations.ModRotPrj) > 0 Then clsDeclarations.ModRotPrj = clsDeclarations.ModRotPrj & "\"
            clsDeclarations.ModRotPrj = clsDeclarations.ModRotPrj & vNewValue
        End Set
    End Property


    Public Property ErrExiste() As Boolean
        Get
            Return clsDeclarations.ErrExiste
        End Get
        Set(vNewValue As Boolean)
            clsDeclarations.ErrExiste = vNewValue

            If Not clsDeclarations.ErrExiste Then
                clsDeclarations.NroPrj = 0
                clsDeclarations.DscPrj = ""
                clsDeclarations.CmdSQLPrj = ""
                clsDeclarations.NroVB = 0
                clsDeclarations.DscVB = ""
                Err.Clear()
            End If
        End Set
    End Property

    Public Property Transferir() As Declaration
        Get
            Return clsDeclarations
        End Get
        Set(vNewValue As Declaration)
            clsDeclarations.NroPrj = vNewValue.NroPrj
            clsDeclarations.DscPrj = vNewValue.DscPrj
            clsDeclarations.ModRotPrj = vNewValue.ModRotPrj
            clsDeclarations.CmdSQLPrj = vNewValue.CmdSQLPrj
            clsDeclarations.NroVB = vNewValue.NroVB
            clsDeclarations.DscVB = vNewValue.DscVB
            clsDeclarations.ErrExiste = vNewValue.ErrExiste
            clsDeclarations.ErrExiste = False
        End Set
    End Property
End Class
