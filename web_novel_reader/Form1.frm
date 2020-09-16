VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9255
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13065
   LinkTopic       =   "Form1"
   ScaleHeight     =   9255
   ScaleWidth      =   13065
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   240
      Width           =   8535
   End
   Begin VB.TextBox Text1 
      Height          =   6735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   960
      Width           =   11175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   -120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error Resume Next
    Set oDoc = CreateObject("htmlfile")
    Set ms = CreateObject("MSScriptControl.ScriptControl")
    ms.Language = "JScript"
    With CreateObject("Microsoft.XMLHTTP")
        .Open "GET", Text2.Text, False
        .setRequestHeader "If-Modified-Since", "Thu, 01 Jan 1970 00:00:00 GMT" '取时时最新数据'等于"0"也可以达到最新
        .send
        oDoc.body.innerHTML = .responseText
        Text1 = oDoc.body.innerText
    End With
End Sub
'纣临 https://www.biquke.com/bq/52/52044/4392319.html
