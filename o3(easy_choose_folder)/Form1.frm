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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function ShowFolderDialog() As String
'/��򵥵���ʾ�ļ���ѡ��Ի��򷽷�
Dim spShell, spFolder, spFolderItem, spPath As String
Const WINDOW_HANDLE = 0
Const NO_OPTIONS = 0
Set spShell = CreateObject("Shell.Application")
Set spFolder = spShell.BrowseForFolder(WINDOW_HANDLE, "ѡ��Ŀ¼:", NO_OPTIONS, "C:\Scripts")
If spFolder Is Nothing Then
    ShowFolderDialog = ""
Else
    Set spFolderItem = spFolder.Self
    spPath = spFolderItem.path
    spPath = Replace(spPath, "\", "\")
    ShowFolderDialog = spPath
End If
End Function
Private Sub Command1_Click()
Dim path As String
path = ShowFolderDialog()
MsgBox IIf(Len(path) = 0, "������ȡ����ť��", "��ѡ����ļ����ǣ�" & path)
End Sub
