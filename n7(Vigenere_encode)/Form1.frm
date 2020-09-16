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
   StartUpPosition =   3  '窗口缺省
   Begin VB.OptionButton Option2 
      Caption         =   "加密"
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   1320
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "解密"
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ins(100) As String, inl As Integer
Dim keys(30) As String, keyl As Integer
Dim ms As Integer, out As String
Dim b(1 To 2) As Integer, e(1 To 2) As Integer

Private Sub init()
ins(0) = Text1.Text
inl = Len(ins(0))
keys(0) = Text2.Text
keyl = Len(keys(0))
For i = 1 To inl
ins(i) = Mid(ins(0), i, 1)
Next i
For i = 1 To keyl
keys(i) = Mid(keys(0), i, 1)
If Asc(keys(i)) < b(2) Then keys(i) = Chr(Asc(keys(i)) + 32)
Next i
End Sub

Private Function trans(x As String, y As String, m As Integer)
Dim dxx As Integer
If Asc(x) < b(2) Then dxx = 1 Else dxx = 2
If m = 1 Then

 If Asc(x) + Asc(y) - b(2) > e(dxx) Then
 trans = Chr(Asc(x) + Asc(y) - b(2) - 26)
 Else
 trans = Chr(Asc(x) + Asc(y) - b(2))
 End If
 Else
 
 If Asc(x) - Asc(y) + b(2) < b(dxx) Then
 trans = Chr(Asc(x) - Asc(y) + b(2) + 26)
 Else
 trans = Chr(Asc(x) - Asc(y) + b(2))
 End If
 
 End If
End Function

Private Sub Command1_Click()
out = ""
init
For i = 1 To inl
If ins(i) = " " Then
out = out & ins(i)
Else
j = j + 1
If j = keyl + 1 Then j = 1
out = out & trans(ins(i), keys(j), ms)
End If
Next i
Text3.Text = out
End Sub

Private Sub Form_Load()
b(1) = 65: b(2) = 97: e(1) = 90: e(2) = 122: ms = 1
End Sub

Private Sub Option1_Click()
ms = 2
End Sub

Private Sub Option2_Click()
ms = 1
End Sub

