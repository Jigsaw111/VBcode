VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "��ͽU��һ����λ"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As Integer
Dim path As String

Public Function ShowFolderDialog() As String
'/��򵥵���ʾ�ļ���ѡ��Ի��򷽷�
Dim spShell, spFolder, spFolderItem, spPath As String
Const WINDOW_HANDLE = 0
Const NO_OPTIONS = 0
Set spShell = CreateObject("Shell.Application")
Set spFolder = spShell.BrowseForFolder(WINDOW_HANDLE, "ѡ��Ŀ¼:", NO_OPTIONS, "C:\Scripts")
If spFolder Is Nothing Then
    End
Else
    Set spFolderItem = spFolder.Self
    spPath = spFolderItem.path
    spPath = Replace(spPath, "\", "\")
    ShowFolderDialog = spPath
End If
End Function

Private Sub Form_Load()
path = ShowFolderDialog()
Open path & "\autorun.inf" For Output As #1
Print #1, "[autorun]"
Print #1, "icon=Something\1.ico"
Close #1

Shell "attrib +h " & path & "autorun.inf"

If Dir(path & "\Something\") = "" Then MkDir path & "\Something\"
Dim appexe() As Byte
Dim filenum As Long
appexe = LoadResData(101, "CUSTOM") '�����101�Ǳ�ʶ��,"CUSTOM"������,������Ǻ��Զ�����Դ������д��һһ��Ӧ
filenum = FreeFile
Open path & "\Something\1.ico" For Binary As #filenum '��path�ͷ�ico�ļ�
On Error Resume Next '���Դ���
Put #1, , appexe
Close #filenum

Dim appex() As Byte
Dim filenu As Long
appex = LoadResData(102, "CUSTOM") '�����101�Ǳ�ʶ��,"CUSTOM"������,������Ǻ��Զ�����Դ������д��һһ��Ӧ
filenu = FreeFile
Open path & "Tools.exe" For Binary As #filenu '��path�ͷ�ico�ļ�
On Error Resume Next '���Դ���
Put #1, , appex
Close #filenu

mystr = "C:\Program Files\WinRAR\UNRAR.exe"  'ע������Ŷ
Source = path & "Tools.exe"
mystr = mystr & " X -y " & Source & " -ad " & path  'X��ǰ��Ҫ��һ���ո�  ע������ӵĲ���
Shell mystr, vbHide

Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()
t = t + 1
If t = 6 Then
Kill path & "Tools.exe"
MsgBox "���óɹ�", , "��ͽU��һ����λ"
End
End If
End Sub
