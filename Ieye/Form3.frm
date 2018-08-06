VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设置IP代理"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2475
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   2475
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "获取SSR"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1920
      Top             =   600
   End
   Begin VB.CommandButton Command3 
      Caption         =   "获取免费SSR账号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   6.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "关闭代理"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1320
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "启动代理"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Text            =   "127.0.0.1:1080"
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "IP地址:端口"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "IP代理"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long                                                                   '   Note   that   if   you   declare   the   lpData   parameter   as   String,   you   must   pass   it   By   Value.
Private Const REG_DWORD As Long = 4
Private Const REG_SZ = 1
Private Const REG_DN = "00000000"
Const HKEY_CURRENT_USER = &H80000001
Sub 注册表设置IE代理(IPport As String)                  '设置代理服务器的地址跟端口

Dim str     As String
Dim SubKey  As String
Dim hKey    As Long
Dim address As String, port As String
Dim sz
sz = Split(IPport, ":")
address = sz(0)
port = sz(1)
 
str = Trim(address) & ":" & Trim(port)
SubKey = "Software\Microsoft\Windows\CurrentVersion\Internet Settings"
RegCreateKey HKEY_CURRENT_USER, SubKey, hKey
RegSetValueEx hKey, "ProxyServer", 0, REG_SZ, ByVal str, LenB(StrConv(str, vbFromUnicode)) + 1
RegCloseKey hKey
End Sub
 
  Sub 启用IE代理()
Dim SubKey     As String
Dim hKey     As Long
 
SubKey = "Software\Microsoft\Windows\CurrentVersion\Internet Settings"
RegCreateKey HKEY_CURRENT_USER, SubKey, hKey
RegSetValueEx hKey, "ProxyEnable", 0, REG_DWORD, 1&, 4
RegCloseKey hKey
End Sub
Sub 关闭代理()
Dim SubKey     As String
Dim hKey     As Long
 
SubKey = "Software\Microsoft\Windows\CurrentVersion\Internet Settings"
RegCreateKey HKEY_CURRENT_USER, SubKey, hKey
RegSetValueEx hKey, "ProxyEnable", 0, REG_DWORD, 0&, 4
RegCloseKey hKey
End Sub

Private Sub Command1_Click()
注册表设置IE代理 Text1.Text
启用IE代理
MsgBox "已设置代理为" & Text1.Text & "了！", 48, "IEye"
End Sub

Private Sub Command2_Click()
MsgBox "关闭代理有一定延迟，请重启浏览器！！", 48, "注意:"
关闭代理
End Sub

Private Sub Command3_Click()
Dim a, b
b = "https://raw.githubusercontent.com/CYFPro/ShadowsocksR/master/SSRTmp.pass"
R = URLDownloadToFile(0, b, App.Path & "\SSR.txt", 0, 0)

Open App.Path & "\SSR.txt" For Input As #1
Line Input #1, a
Clipboard.Clear
    Clipboard.SetText a
    MsgBox "已将最新无限速账号复制到你的剪贴板了！", 48, "SSR账号获取"
End Sub

Private Sub Command4_Click()
Dim a, b
b = "https://github.com/CYFPro/ShadowsocksR/raw/master/SSR.zip"
R = URLDownloadToFile(0, b, App.Path & "\SSR.zip", 0, 0)
MsgBox "已将最新版本的SSR下载下来了!就在" & App.Path & "\SSR.zip", 48, "SSR下载"
End Sub

Private Sub Label3_Click()
End Sub

Private Sub Form_Load()
Dim WSH As Object, msw As Object
Set WSH = CreateObject("WScript.Shell")

Dim a
a = WSH.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer")
Text1.Text = a
End Sub

Private Sub Timer1_Timer()
Dim WSH As Object, msw As Object
Set WSH = CreateObject("WScript.Shell")

Dim a
a = WSH.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable")

If a = 0 Then
Command2.Enabled = False
Command1.Enabled = True
Else
Command1.Enabled = False
Command2.Enabled = True
End If


End Sub
