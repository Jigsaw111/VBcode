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
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   615
      Left            =   1320
      TabIndex        =   1
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "App"
      Top             =   720
      Width           =   3855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   2280
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal ascii As Long) As Long

Dim app As String

Private Sub Command1_Click()
If Not Text1.Text = "" Then
If Right(Text1.Text, 4) = ".exe" Then app = Text1.Text Else app = Text1.Text & ".exe"
Form1.Visible = False: MsgBox "Press 'F1' to stop", , "Killer": Timer1.Enabled = True
Else
MsgBox "Please write down application", , "Killer"
End If
End Sub

Private Sub Timer1_Timer()
If GetAsyncKeyState(vbKeyF1) Then Timer1.Enabled = False: Form1.Visible = True
Shell "taskkill /f /im " & app, vbHide
End Sub
