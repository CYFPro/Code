VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{1339B53E-3453-11D2-93B9-000000000000}#1.0#0"; "mozctl.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7140
   ClientLeft      =   120
   ClientTop       =   2760
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   11700
   Begin VB.CommandButton Command5 
      Caption         =   "搜"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   12
      Top             =   0
      Width           =   615
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "Form1.frx":0000
      Left            =   6480
      List            =   "Form1.frx":001C
      TabIndex        =   11
      Text            =   "谷歌（中国）"
      Top             =   0
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   10
      Text            =   "Google"
      Top             =   0
      Width           =   2895
   End
   Begin VB.CommandButton Command4 
      Caption         =   "主"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   615
   End
   Begin MOZILLACONTROLLibCtl.MozillaBrowser MozillaBrowser1 
      Height          =   1215
      Left            =   0
      OleObjectBlob   =   "Form1.frx":009C
      TabIndex        =   7
      Top             =   960
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "Form1.frx":00C0
      Left            =   5040
      List            =   "Form1.frx":00CA
      TabIndex        =   6
      Text            =   "IE"
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "新"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "设"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   0
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   6480
      Top             =   480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "刷"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   735
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
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Text            =   "https://Github.com"
      Top             =   480
      Width           =   4215
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   10335
      ExtentX         =   18230
      ExtentY         =   9551
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label2 
      Caption         =   "搜:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   9
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   6360
      Width           =   11775
   End
   Begin VB.Menu Iwent 
      Caption         =   "我想要..."
      Begin VB.Menu Home 
         Caption         =   "打开主页"
      End
      Begin VB.Menu NewWindow 
         Caption         =   "新的窗口"
      End
      Begin VB.Menu F5 
         Caption         =   "刷新"
         Shortcut        =   {F5}
      End
      Begin VB.Menu IE 
         Caption         =   "用IE打开"
      End
      Begin VB.Menu Oil 
         Caption         =   "打开油航"
      End
      Begin VB.Menu Close 
         Caption         =   "关闭"
      End
   End
   Begin VB.Menu Setting 
      Caption         =   "设置"
      Begin VB.Menu IP 
         Caption         =   "IP代理"
      End
      Begin VB.Menu clean 
         Caption         =   "清除缓存、Cookie、历史纪录"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub clean_Click()
Shell "cmd.exe /c" & "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 4351"
End Sub

Private Sub Close_Click()
Unload Me
End Sub

Private Sub Command1_Click()
MozillaBrowser1.Visible = False
WebBrowser1.Visible = False
Label1.Caption = "载入中..."
If Combo1.Text = "FF" Then
MsgBox "HTTP协议推荐使用火狐内核,HTTPS火狐无法打开!!", 32, "IEye"
MozillaBrowser1.Visible = True
MozillaBrowser1.Navigate (Text1.Text)
Else

WebBrowser1.Navigate (Text1.Text)
WebBrowser1.Visible = True
End If
End Sub

Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Command3_Click()
Shell "cmd.exe /c" & App.Path & "\IEye.exe"
End Sub

Private Sub Command4_Click()
Open App.Path & "\Setting.dll" For Input As #1
Input #1, ZY
Text1.Text = ZY
WebBrowser1.Navigate (ZY)


End Sub

Private Sub Command5_Click()
If Combo2.Text = "GOOGLE（美国）" Then
WebBrowser1.Navigate "https://www.google.com/search?client=aff-cs-360chromium&ie=UTF-8&q=" & Text2.Text
Else
End If
If Combo2.Text = "百度" Then
WebBrowser1.Navigate "https://www.baidu.com/s?ie=utf-8&f=8&rsv_bp=1&tn=baidu&wd=" & Text2.Text
Else
End If
If Combo2.Text = "谷歌（中国）" Then
WebBrowser1.Navigate "https://www.google.cn/search?client=aff-cs-360chromium&ie=UTF-8&q=" & Text2.Text
Else
End If
If Combo2.Text = "Bing（必应）" Then
WebBrowser1.Navigate "https://www.bing.com/search?q=" & Text2.Text
Else
End If
If Combo2.Text = "Yahoo！（雅虎）" Then
WebBrowser1.Navigate "https://search.yahoo.com/search?p=" & Text2.Text
Else
End If
If Combo2.Text = "Wikipedia（维基英文）" Then
WebBrowser1.Navigate "https://en.wikipedia.org/wiki/" & Text2.Text
Else
End If
If Combo2.Text = "维基百科（维基中文）" Then
WebBrowser1.Navigate "https://zh.wikipedia.org/wiki/" & Text2.Text
Else
End If
If Combo2.Text = "DuckduckGo" Then
WebBrowser1.Navigate "https://duckduckgo.com/?q=" & Text2.Text
Else
End If

End Sub

Private Sub F5_Click()
MozillaBrowser1.Visible = False
WebBrowser1.Visible = False
Label1.Caption = "载入中..."
If Combo1.Text = "FF" Then
MsgBox "HTTP协议推荐使用火狐内核,HTTPS火狐无法打开!!", 32, "IEye"
MozillaBrowser1.Visible = True
MozillaBrowser1.Navigate (Text1.Text)
Else

WebBrowser1.Navigate (Text1.Text)
WebBrowser1.Visible = True
End If
End Sub

Private Sub Form_Load()
Label1.Caption = "欢迎使用IEye-IE增强浏览器!载入主页中..."
Open App.Path & "\Setting.dll" For Input As #1
Input #1, ZY
Text1.Text = ZY
WebBrowser1.Navigate (ZY)
Close #1
End Sub

Private Sub Form_Resize()
MozillaBrowser1.Visible = True


WebBrowser1.Height = Form1.Height - 2100
WebBrowser1.Width = Form1.Width - 330
MozillaBrowser1.Height = Form1.Height - 2100
MozillaBrowser1.Width = Form1.Width - 330
Label1.Width = Form1.Width - 330
Label1.Top = Form1.Height - 1115
WebBrowser1.Left = 20
MozillaBrowser1.Left = 20
MozillaBrowser1.Visible = False
WebBrowser1.Visible = False
Label1.Caption = "载入中..."
If Combo1.Text = "FF" Then

MozillaBrowser1.Visible = True
MozillaBrowser1.Navigate (Text1.Text)
Else

WebBrowser1.Navigate (Text1.Text)
WebBrowser1.Visible = True
End If
End Sub

Private Sub Home_Click()
Open App.Path & "\Setting.dll" For Input As #1
Input #1, ZY
Text1.Text = ZY
WebBrowser1.Navigate (ZY)


End Sub

Private Sub IE_Click()
Shell "cmd.exe /c" & "C:\Program Files\Internet Explorer\iexplore.exe " & Text1.Text
End Sub

Private Sub IP_Click()
Form3.Show
End Sub

Private Sub NewWindow_Click()
Shell "cmd.exe /c" & App.Path & "\IEye.exe"
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)

Label1.Caption = " IE内核所访问的是: " & URL
Text1.Text = URL
End Sub
Private Sub MozillaBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)

Label1.Caption = "火狐内核所访问的是: " & URL
Text1.Text = URL
End Sub

