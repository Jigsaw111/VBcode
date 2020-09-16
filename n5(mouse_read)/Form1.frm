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
Private Declare Sub Sleep Lib "kernel32" (ByVal a As Long)

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal a As Byte) As Long

Private Declare Sub keybd_event Lib "user32" (ByVal a As Byte, ByVal b As Byte, ByVal c As Long, ByVal d As Long)
Private Const up = &H2

Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
 


Private Sub mouse(x As Long, y As Long)
    SetCursorPos x, y
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

Private Sub key(a As Integer)
keybd_event a, 0, 0, 0
keybd_event a, 0, up, 0
End Sub

Private Sub exam(s As String)
Dim i As Integer, x As Long, y As Long
For i = 1 To Len(s)
If Mid(s, i, 1) = "," Then x = Mid(s, 1, i - 1): y = Mid(s, i + 1, Len(s)): mouse x, y
Next i
If s = "Tab" Then key vbKeyTab
If s = "CapsLock" Then key vbKeyCapital
If s = "Shift" Then key vbKeyShift
If s = "Bksp" Then key vbKeyBack
If s = "Ctrl" Then key vbKeyControl
If s = "Alt" Then key vbKeyMenu
If s = "Enter" Then key 13
If s = "Esc" Then key vbKeyEscape
If s = "+" Then key vbKeyAdd
If s = "`" Then key 192
If s = "-" Then key 189
If s = "=" Then key 187
If s = "Home" Then key vbKeyHome
If s = "PgUp" Then key vbKeyPageUp
If s = "PgDn" Then key vbKeyPageDown
If s = "/" Then key 111
If s = "[" Then key 219
If s = "]" Then key 221
If s = "\" Then key 220
If s = "Del" Then key vbKeyDelete
If s = "End" Then key vbKeyEnd
If s = "*" Then key 106
If s = ";" Then key 186
If s = "'" Then key 222
If s = "Insert" Then key vbKeyInsert
If s = "Pause" Then key vbKeyPause
If s = "-" Then key 109
If s = "," Then key 188
If s = "." Then key 110
If s = "/" Then key 191
If s = "Up" Then key vbKeyUp
If s = "Down" Then key vbKeyDown
If s = "Left" Then key vbKeyLeft
If s = "Right" Then key vbKeyRight
If s = "PrtScn" Then key vbKeyPrint
If s = "ScrLock" Then key vbKeyScrollLock
If s = "Space" Then key vbKeySpace
If s = "NumLock" Then key vbKeyNumlock
If s = "Help" Then key vbKeyHelp
For i = 65 To 90 ' a to z
If s = Chr(i + 32) Then key i
Next i
For i = 48 To 57 ' 0 to 9
If s = Chr(i) Then key i
Next i
End Sub


Private Sub Timer1_Timer()
If GetAsyncKeyState(vbKeyF2) Then
Dim s As String
Open "1.txt" For Input As #1
Do Until EOF(1)
Input #1, s
exam s
Loop
Close #1
End If
End Sub
