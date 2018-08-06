VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设置"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2820
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   2820
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "检查更新"
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "打开代理"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "清除历史纪录"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   480
      TabIndex        =   1
      Text            =   ".Gooer主页"
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "当前版本：Ieye Alpha V0.8.1"
      Height          =   735
      Left            =   1200
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "主页"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Sub Command1_Click()
Shell "cmd.exe /c" & "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 4351"
End Sub


Private Sub Command2_Click()
Kill App.Path & "\NVer.dll"
a = "https://raw.githubusercontent.com/CYFPro/About-IEye/master/NVer.dll"
R = URLDownloadToFile(0, a, App.Path & "\NVer.dll", 0, 0)
Open App.Path & "\NVer.dll" For Input As #1
Dim NVer, Ver
Input #1, NVer
Close #1

Open App.Path & "\Ver.dll" For Input As #1
Input #1, Ver
Close #1

If NVer = Ver Then
MsgBox "版本Ieye Alpha V0.8.1没有更新", 48, "更新"
Else
MsgBox "版本" & NVer & "有更新!点击确定下载!", 48, "IEye"
a = "https://raw.githubusercontent.com/CYFPro/About-IEye/master/IEye%E6%96%B0%E7%89%88%E6%9C%AC%E7%AE%80%E4%BB%8B.txt"
R = URLDownloadToFile(0, a, App.Path & "\IEye新版本简介.txt", 0, 0)

Shell "C:\Windows\notepad.exe " & App.Path & "\IEye新版本简介.txt"
a = "https://raw.githubusercontent.com/CYFPro/About-IEye/master/Ieyer.exe"
R = URLDownloadToFile(0, a, App.Path & "\Ieyer.exe", 0, 0)
Shell "cmd.exe /c " & App.Path & "\Ieyer.exe"
End If
End Sub

Private Sub Command3_Click()
Form3.Show

End Sub

Private Sub Form_Load()

Open App.Path & "\Setting.dll" For Input As #1
Input #1, ZY
Text1.Text = ZY

Close #1

End Sub

Private Sub Form_Unload(Cancel As Integer)

Open App.Path & "\Setting.dll" For Output As #1
Close #1
Kill App.Path & "\Setting.dll"
Open App.Path & "\Setting.dll" For Output As #1
Dim a
a = Text1.Text
Write #1, a
Close #1
MsgBox "已保存!", 32, "IEye浏览器"
End Sub

