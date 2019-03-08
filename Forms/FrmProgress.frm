VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FrmProgress 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   960
      TabIndex        =   4
      Top             =   4440
      Width           =   5895
      Begin VB.Label LblContent 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "努力加载中......"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   5775
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   6000
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   5055
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3375
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   4815
      ExtentX         =   8493
      ExtentY         =   5953
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
      Location        =   ""
   End
End
Attribute VB_Name = "FrmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : FrmProgress
' Author    : YPN
' Date      : 2018-04-08 20:51
' Purpose   : 进度窗体
'---------------------------------------------------------------------------------------
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2


Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Call formTransparent(200)
    With WebBrowser1
        .AddressBar = False
        .MenuBar = False
        .StatusBar = False
        .TheaterMode = False
        .Navigate ("C:\WINDOWS\system32\loading.gif")        '播放gif动画
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : formTransparent
' Author    : YPN
' Date      : 2015/02/15
' Purpose   : 窗体透明
'---------------------------------------------------------------------------------------
'
Private Sub formTransparent(i_alpha As Integer)
    If i_alpha >= 0 And i_alpha <= 255 Then
        Dim rtn As Long
        
        rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
        rtn = rtn Or WS_EX_LAYERED
        SetWindowLong hWnd, GWL_EXSTYLE, rtn
        SetLayeredWindowAttributes hWnd, 0, i_alpha, LWA_ALPHA
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    WebBrowser1.Top = Screen.Height / 4
    WebBrowser1.Left = Screen.Width / 3
    
    Frame1.Top = WebBrowser1.Top - 90
    Frame1.Left = WebBrowser1.Left
    
    Frame2.Top = WebBrowser1.Top
    Frame2.Left = WebBrowser1.Left + WebBrowser1.Width - 460
    
    Frame3.Top = WebBrowser1.Top + WebBrowser1.Height - 280
    Frame3.Left = WebBrowser1.Left
    
    Frame4.Top = WebBrowser1.Top
    Frame4.Left = WebBrowser1.Left - 150
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

