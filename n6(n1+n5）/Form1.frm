VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1680
      Top             =   1320
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   720
      Top             =   1560
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal a As Byte) As Long

Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const lup = &H2
Private Const ldown = &H4
Private Const rup = &H8
Private Const rdown = &H10

Dim t As Integer

Private Sub Timer1_Timer()
If GetAsyncKeyState(vbKeyF2) Then
t = t + 1
If t Mod 2 = 1 Then
MsgBox "ON"
Timer2.Enabled = True
Else
Timer2.Enabled = False
MsgBox "OFF"
End If
End If

End Sub

Private Sub Timer2_Timer()
mouse_event ldown, 0, 0, 0, 0: mouse_event lup, 0, 0, 0, 0
End Sub
