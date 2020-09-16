VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   12
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1680
      Top             =   1320
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1680
      Top             =   1320
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   16200
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   28800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As Integer

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal asc As Long) As Long

Private Sub Form_Load()
Text1.Text = "系统发生故障 正在尝试修复"
End Sub

Private Sub Text1_Change()
Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub Timer1_Timer()
Dim s As String, a As Integer
If GetAsyncKeyState(vbKeyF1) Then End
For i = 1 To 240
a = Fix(Rnd * 80)
If a > 1 Then s = s & " " Else s = s & a
Randomize
Next i
Text1.Text = Text1.Text & vbCrLf & s
t = t + 1
If t = 66 Then t = 0: Text1.Text = Right(Text1.Text, 0.5 * Len(Text1.Text))
End Sub

Private Sub Timer2_Timer()
t = t + 1
If t = 3 Then Form1.WindowState = 2: Timer2.Enabled = False: Timer1.Enabled = True
End Sub
