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
   Begin VB.TextBox Text3 
      Height          =   1575
      Left            =   3360
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim Email As Object
On Error Resume Next
Const NameSpace = "http://schemas.microsoft.com/cdo/configuration/"
Set Email = CreateObject("cdo.message")
Email.From = "13015621507@163.com"  '发件人邮箱
Email.to = "1290541225@qq.com"  '收件人邮箱
Email.Subject = "收租提醒" '主题
Email.Textbody = "您收租日期到了！" '邮件内容
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
MsgBox "邮件发送成功!"
End Sub

