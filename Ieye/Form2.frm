VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设置"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2835
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   2835
   StartUpPosition =   3  '窗口缺省
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
Private Sub Command1_Click()
Shell "cmd.exe /c" & "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 4351"
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
