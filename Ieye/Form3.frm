VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����IP����"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2475
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   2475
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command4 
      Caption         =   "��ȡSSR"
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��ȡ���SSR�˺�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   6.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�رմ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1320
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Text            =   "1080"
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
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
      Text            =   "127.0.0.1"
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "�˿�:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "IP��ַ:"
      BeginProperty Font 
         Name            =   "����"
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
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "IP����"
      BeginProperty Font 
         Name            =   "����"
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
Sub ע�������IE����(IPport As String)                  '���ô���������ĵ�ַ���˿�

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
 
  Sub ����IE����()
Dim SubKey     As String
Dim hKey     As Long
 
SubKey = "Software\Microsoft\Windows\CurrentVersion\Internet Settings"
RegCreateKey HKEY_CURRENT_USER, SubKey, hKey
RegSetValueEx hKey, "ProxyEnable", 0, REG_DWORD, 1&, 4
RegCloseKey hKey
End Sub
Sub �رմ���()
Shell "cmd.exe /c" & App.Path & "\NO.reg"
End Sub

Private Sub Command1_Click()
ע�������IE���� Text1.Text & ":" & Text2.Text
����IE����
End Sub

Private Sub Command2_Click()
MsgBox "��ͬ�⵼���REG,���ǰ�ȫ�ģ�ֻ��Ϊ�˹رմ���", 32, "ע��:"
�رմ���
End Sub

Private Sub Command3_Click()
Dim A, b
b = "https://raw.githubusercontent.com/CYFPro/ShadowsocksR/master/SSRTmp.pass"
R = URLDownloadToFile(0, b, App.Path & "\SSR.txt", 0, 0)

Open App.Path & "\SSR.txt" For Input As #1
Line Input #1, A
Clipboard.Clear
    Clipboard.SetText A
    MsgBox "�ѽ������������˺Ÿ��Ƶ���ļ������ˣ�", 32, "SSR�˺Ż�ȡ"
End Sub

Private Sub Command4_Click()
Dim A, b
b = "https://github.com/CYFPro/ShadowsocksR/raw/master/SSR.zip"
R = URLDownloadToFile(0, b, App.Path & "\SSR.zip", 0, 0)
MsgBox "�ѽ����°汾��SSR����������!����" & App.Path & "\SSR.zip", 32, "SSR����"
End Sub
