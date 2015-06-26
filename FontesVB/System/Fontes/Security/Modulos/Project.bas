Attribute VB_Name = "mProject"
Option Explicit

Public Sub mCmbSimNao(objCmb As Object, Optional Default As Boolean = True)
   If Not TypeName(objCmb) = "ComboBox" Then Exit Sub
   
   objCmb.Clear
   objCmb.AddItem ("Não")
   objCmb.ItemData(objCmb.NewIndex) = 0
   
   objCmb.AddItem ("Sim")
   objCmb.ItemData(objCmb.NewIndex) = 1
   
   If Default Then
      objCmb.ListIndex = f1.CmbValor(objCmb, 1)
   Else
      objCmb.ListIndex = f1.CmbValor(objCmb, 1)
   End If
End Sub

Public Sub mCmbValorPerc(objCmb As ComboBox)
   If Not TypeName(objCmb) = "ComboBox" Then Exit Sub
   
   objCmb.Clear
   objCmb.AddItem ("%")
   objCmb.ItemData(objCmb.NewIndex) = 0
   
   objCmb.AddItem ("R$")
   objCmb.ItemData(objCmb.NewIndex) = 1
   objCmb.ListIndex = f1.CmbValor(objCmb, 0)
End Sub
