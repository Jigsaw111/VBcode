VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10755
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   6090
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10755
   ScaleWidth      =   6090
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   10095
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   6135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Top             =   10200
      Width           =   1215
   End
   Begin VB.Menu Menu1 
      Caption         =   "文件"
      NegotiatePosition=   1  'Left
      WindowList      =   -1  'True
      Begin VB.Menu 没 
         Caption         =   "打开"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim zj(3000) As String, z As Integer, s As String, zz As Integer

Private Sub Command1_Click()
Open "xs.txt" For Input As #1
Do Until EOF(1)
Input #1, s
For i = 1 To Len(s) - 2
If Mid(s, i, 1) = "第" Then
For j = i + 2 To i + 8
If Mid(s, j, 1) = "章" Then z = z + 1: j = i + 8
Next j
End If
Next i
If z > 0 Then
If zj(z) = "" Then zj(z) = zj(z) & s Else zj(z) = zj(z) & vbCrLf & s
End If
Loop
Close #1
Text1.Text = zj(zz)
End Sub

Private Sub 没_Click()

End Sub
