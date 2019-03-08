VERSION 5.00
Begin VB.Form FrmMessage 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer2 
      Left            =   3960
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Left            =   3480
      Top             =   480
   End
   Begin VB.Label lbl_DateTime 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "2018-01-23 14:59:00"
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lbl_MsgSource 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ypn.common.vb"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label lbl_MsgContent 
      BackColor       =   &H00FFFFFF&
      Caption         =   "测试内容"
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4935
   End
   Begin VB.Label lbl_MsgTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "测试标题"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "FrmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : FrmMessage
' Author    : YPN
' Date      : 2018-04-08 21:01
' Purpose   : 屏幕右下角弹出窗
'---------------------------------------------------------------------------------------
Option Explicit
'任务栏高度
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48
Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type
'透明
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
'延迟
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'最前
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_BOTTOM = 1
Private Const HWND_BROADCAST = &HFFFF&
Private Const HWND_DESKTOP = 0
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
'可见区域
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private f_MyRect     As Long
Private f_MyRgn      As Long
Private f_X1         As Integer, f_Y1 As Integer
Private f_X2         As Integer, f_Y2 As Integer
Private f_OpenSpeed  As Integer
Private f_CloseSpeed As Integer
Public F_WaitTime   As Integer  '关闭前等待时间(秒)，为0则不会自动关闭


Private Sub Form_Load()
    '------------------------------------------------------------------
    f_OpenSpeed = 10     '出现时速度
    f_CloseSpeed = 10    '关闭时淡出的速度
    Timer1.Interval = 10 '出现时显示平滑度
    lbl_DateTime = Now()
    '------------------------------------------------------------------
    
    '计算任务栏高
    Dim v_Res           As Long
    Dim v_RectVal       As RECT
    Dim v_TaskbarHeight As Integer
    
    v_Res = SystemParametersInfo(SPI_GETWORKAREA, 0, v_RectVal, 0)
    v_TaskbarHeight = Screen.Height - v_RectVal.Bottom * Screen.TwipsPerPixelY
    '确定位置
    'Me.Move Screen.Width * 0.75, Screen.Height * 0.75 - v_TaskbarHeight, Screen.Width \ 4, Screen.Height \ 4    '相对位置
    Me.Move Screen.Width - Me.Width, Screen.Height - Me.Height - v_TaskbarHeight, Me.Width, Me.Height            '使自适应
    '永在最前
    SetWindowPos Me.hWnd, HWND_TOPMOST, Me.Left \ Screen.TwipsPerPixelX, Me.Top \ Screen.TwipsPerPixelY, Me.Width, Me.Height, 1
    '为遮蔽窗体计算坐标
    f_X1 = 0
    f_Y1 = Me.Width \ Screen.TwipsPerPixelX
    f_X2 = Me.Width \ Screen.TwipsPerPixelX
    f_Y2 = Me.Height \ Screen.TwipsPerPixelY - 1
    '遮蔽部分窗体为不可见
    f_MyRect = CreateRectRgn(f_X1, f_Y1, f_X2, f_Y2)
    f_MyRgn = SetWindowRgn(Me.hWnd, f_MyRect, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call closeMe(1) '以什么样的方式关闭自己，有 1-淡出 和 2-收缩 可选
    Call DeleteObject(f_MyRect)
End Sub

Private Sub Timer1_Timer()
    f_Y2 = f_Y2 - f_OpenSpeed
    If f_Y2 <= 0 Then
        f_MyRect = CreateRectRgn(0, 0, Me.Width \ Screen.TwipsPerPixelX, f_Y2)
        f_MyRgn = SetWindowRgn(Me.hWnd, f_MyRect, True)
        
        Timer1.Enabled = False
        
        '----------------------
        If F_WaitTime <> 0 Then
            Timer2.Interval = 1000
            Timer2.Enabled = True
        End If
    End If
    f_MyRect = CreateRectRgn(f_X1, f_Y1, f_X2, f_Y2)
    f_MyRgn = SetWindowRgn(Me.hWnd, f_MyRect, True)
End Sub

Private Sub Timer2_Timer()
    Static v_NL As Integer
    v_NL = v_NL + 1
    If v_NL >= F_WaitTime Then
        v_NL = 0
        Unload Me
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : closeMe
' Author    : YPN
' Date      : 2018/01/23 14:25
' Purpose   :
' Param     : i_N：0 - 不使用卸载效果
'                  1 - 使用透明淡出效果
'                  2 - 使用收缩效果
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Private Sub closeMe(Optional i_N As Integer = 1)
    Select Case i_N
    Case 0
        Exit Sub
        
    Case 1
        Dim rtn As Long
        
        rtn = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
        rtn = rtn Or WS_EX_LAYERED
        SetWindowLong Me.hWnd, GWL_EXSTYLE, rtn
        
        For i = 255 To 10 Step -10
            SetLayeredWindowAttributes Me.hWnd, 0, i, LWA_ALPHA
            DoEvents
            Sleep f_CloseSpeed
        Next i
        
    Case 2
        While f_Y2 < (Me.Height / Screen.TwipsPerPixelY)
            f_Y2 = f_Y2 + f_OpenSpeed
            f_MyRect = CreateRectRgn(f_X1, f_Y1, f_X2, f_Y2)
            f_MyRgn = SetWindowRgn(Me.hWnd, f_MyRect, True)
            Sleep f_OpenSpeed
        Wend
    End Select
End Sub
