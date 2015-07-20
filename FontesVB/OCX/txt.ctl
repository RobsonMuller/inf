VERSION 5.00
Begin VB.UserControl txt 
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1845
   ScaleHeight     =   330
   ScaleWidth      =   1845
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1845
   End
End
Attribute VB_Name = "txt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Change()
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private m_ForeColor As OLE_COLOR

Private Sub Text1_Change()
   RaiseEvent Change
End Sub

Public Property Get Enabled() As Boolean
   Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
   UserControl.Enabled() = vNewValue
   Me.ForeColor = IIf(vNewValue, m_ForeColor, vbGrayText)
   PropertyChanged "Enabled"
End Property

Public Property Get ForeColor() As OLE_COLOR
   ForeColor = Text1.ForeColor
End Property

Public Property Let ForeColor(ByVal vNewValue As Variant)
   Text1.ForeColor() = vNewValue
   PropertyChanged "ForeColor"
End Property

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseUp(Button, Shift, ScaleX(X, vbTwips, vbContainerPosition), ScaleY(Y, vbTwips, vbContainerPosition))
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseUp(Button, Shift, _
      ScaleX(X, vbTwips, vbContainerPosition), _
      ScaleY(Y, vbTwips, vbContainerPosition))
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseUp(Button, Shift, ScaleX(X, vbTwips, vbContainerPosition), ScaleY(Y, vbTwips, vbContainerPosition))
End Sub

Private Sub UserControl_Resize()
   Dim sngHeight As Single
   Dim lngTexto As Long
   
   sngHeight = IIf(ScaleHeight < 315, 315, ScaleHeight)
   lngTexto = (UserControl.Width - 120)
   lngTexto = IIf(lngTexto > 180, lngTexto, 180)
   Text1.Move 0, 0, lngTexto, sngHeight
   UserControl.Width = lngTexto
   Text1.Text = AlinharTexto(Trim(Text1.Text))
End Sub

Private Function AlinharTexto(Texto As String) As String
   Dim lngWidth As Long
   
   AlinharTexto = Texto
   lngWidth = (Text1.Width - (Screen.TwipsPerPixelX * 8))
   If UserControl.TextWidth(AlinharTexto) < lngWidth Then
      While UserControl.TextWidth(AlinharTexto) < lngWidth
         AlinharTexto = " " & AlinharTexto
      Wend
   End If
End Function

Public Property Get hDC() As Long
   hDC = UserControl.hDC
End Property

Public Property Get hwnd() As Long
   hwnd = UserControl.hwnd
End Property
