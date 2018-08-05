VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{1339B53E-3453-11D2-93B9-000000000000}#1.0#0"; "mozctl.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6915
   ClientLeft      =   120
   ClientTop       =   2460
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   11550
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "Form1.frx":0000
      Left            =   4200
      List            =   "Form1.frx":000A
      TabIndex        =   7
      Text            =   "IE"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "新"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   5
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "设置/Settings"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Top             =   0
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   5760
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GO!/刷新"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Text            =   "https://Github.com"
      Top             =   0
      Width           =   4215
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   480
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
   Begin MOZILLACONTROLLibCtl.MozillaBrowser MozillaBrowser1 
      Height          =   1575
      Left            =   0
      OleObjectBlob   =   "Form1.frx":0016
      TabIndex        =   6
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   6000
      Width           =   10695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
MozillaBrowser1.Visible = False
WebBrowser1.Visible = False
Label1.Caption = "载入中..."
If Combo1.Text = "FF" Then
MsgBox "HTTP协议推荐使用火狐内核,HTTPS火狐无法打开!!", 32, "IEye"
MozillaBrowser1.Visible = True
MozillaBrowser1.Navigate (Text1.Text)
Else
MsgBox "HTTPS/FTP协议推荐使用IE原生内核,HTTPS火狐无法打开!!", 32, "IEye"
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

Private Sub Form_Load()
Label1.Caption = "欢迎使用IEye-IE增强浏览器!载入主页中..."
Open App.Path & "\Setting.dll" For Input As #1
Input #1, ZY
Text1.Text = ZY
WebBrowser1.Navigate (ZY)
Close #1
End Sub

Private Sub Form_Resize()
WebBrowser1.Height = Form1.Height - 1325
WebBrowser1.Width = Form1.Width - 330
MozillaBrowser1.Height = Form1.Height - 1325
MozillaBrowser1.Width = Form1.Width - 330
Label1.Width = Form1.Width - 330
Label1.Top = Form1.Height - 780
WebBrowser1.Left = 20
MozillaBrowser1.Left = 20
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)

Label1.Caption = " IE内核所访问的是: " & URL
Text1.Text = URL
End Sub
Private Sub MozillaBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)

Label1.Caption = "火狐内核所访问的是: " & URL
Text1.Text = URL
End Sub
