VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13845
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   13845
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "Form2"
      Height          =   375
      Left            =   240
      TabIndex        =   33
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   2160
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   5280
      Width           =   2775
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   2160
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   4800
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "去除两边空格与换行符"
      Height          =   495
      Left            =   11880
      TabIndex        =   28
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox Text11 
      Height          =   1335
      Left            =   7560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   27
      Text            =   "Form1.frx":0000
      Top             =   3360
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "弹出屏幕右下角消息窗"
      Height          =   375
      Left            =   7560
      TabIndex        =   26
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate"
      Height          =   375
      Left            =   7560
      TabIndex        =   24
      Top             =   840
      Width           =   1095
   End
   Begin VB.ComboBox cmb1 
      Height          =   300
      Index           =   0
      Left            =   8760
      Style           =   2  'Dropdown List
      TabIndex        =   23
      ToolTipText     =   "Version"
      Top             =   840
      Width           =   1095
   End
   Begin VB.ComboBox cmb1 
      Height          =   300
      Index           =   1
      Left            =   9960
      Style           =   2  'Dropdown List
      TabIndex        =   22
      ToolTipText     =   "Error correction level"
      Top             =   840
      Width           =   1095
   End
   Begin VB.ComboBox cmb1 
      Height          =   300
      Index           =   2
      Left            =   8760
      Style           =   2  'Dropdown List
      TabIndex        =   21
      ToolTipText     =   "Mask type"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text10 
      Height          =   765
      Left            =   11160
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   20
      Text            =   "Form1.frx":0007
      Top             =   840
      Width           =   2655
   End
   Begin VB.ComboBox cmb1 
      Height          =   300
      Index           =   3
      Left            =   9960
      Style           =   2  'Dropdown List
      TabIndex        =   19
      ToolTipText     =   "Text encoding"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   2160
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   4320
      Width           =   4575
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   3360
      Width           =   2775
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   510
      Width           =   2775
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "2018-04-20的月末："
      Height          =   180
      Left            =   240
      TabIndex        =   32
      Top             =   5370
      Width           =   1620
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "2018-04-20的月初："
      Height          =   180
      Left            =   240
      TabIndex        =   30
      Top             =   4890
      Width           =   1620
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   8880
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "二维码生成测试："
      Height          =   180
      Left            =   7560
      TabIndex        =   25
      Top             =   480
      Width           =   1440
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "GUID："
      Height          =   180
      Left            =   240
      TabIndex        =   18
      Top             =   4410
      Width           =   540
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "从文件名获取后缀名："
      Height          =   180
      Left            =   240
      TabIndex        =   16
      Top             =   3930
      Width           =   1800
   End
   Begin VB.Label Label71 
      AutoSize        =   -1  'True
      Caption         =   "从全路径获取文件名："
      Height          =   180
      Left            =   4920
      TabIndex        =   14
      Top             =   3480
      Width           =   1800
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "从全路径获取文件名："
      Height          =   180
      Left            =   240
      TabIndex        =   13
      Top             =   3450
      Width           =   1800
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "第一个汉字首字母："
      Height          =   180
      Left            =   240
      TabIndex        =   11
      Top             =   2970
      Width           =   1620
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "所有汉字首字母："
      Height          =   180
      Left            =   240
      TabIndex        =   9
      Top             =   2490
      Width           =   1440
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "MD5加密物理盘序列号："
      Height          =   180
      Left            =   240
      TabIndex        =   7
      Top             =   2010
      Width           =   1890
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "物理盘型号："
      Height          =   180
      Left            =   240
      TabIndex        =   5
      Top             =   1530
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "物理盘序列号："
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   1050
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "逻辑盘序列号："
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1260
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function WideCharToMultiByte Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpWideCharStr As Any, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByRef lpDefaultChar As Any, ByRef lpUsedDefaultChar As Any) As Long
Private Const CP_UTF8 As Long = 65001




Private Sub Command1_Click()

    'test
    Dim b2() As Byte
    Dim s As String
    Dim i As Long, m As Long
    '///
    For i = 0 To cmb1.UBound
        If cmb1(i).ListIndex < 0 Then Exit Sub
    Next i
    '///
    Select Case cmb1(3).ListIndex
    Case 1
        s = Text10.Text
        m = Len(s)
        i = m * 3 + 64
        ReDim b2(i)
        m = WideCharToMultiByte(CP_UTF8, 0, ByVal StrPtr(s), m, b2(0), i, ByVal 0, ByVal 0)
    Case Else
        s = StrConv(Text10.Text, vbFromUnicode)
        b2 = s
        m = LenB(s)
    End Select
    '///
'    Set Image1.Picture = v_QRC.Encode(Text10.Text, m, cmb1(0).ListIndex, cmb1(1).ListIndex + 1, cmb1(2).ListIndex - 1)
    '//YPN.QRCode("",21,0,2,-1)
    Set Image1.Picture = YPN.QRCode(Text10.Text, cmb1(0).ListIndex, cmb1(1).ListIndex + 1, cmb1(2).ListIndex - 1, cmb1(3).Text)
    Set Image2.Picture = YPN.BarCode128(Text10.Text, 6, True)

End Sub

Private Sub Command2_Click()

    Call YPN.ShowMessage("WRP-PDP", Me.Icon, "产前数据准备系统消息", "合同评审", "您有新的合同需要评审！" & vbCrLf & "请及时进入【合同内容评审】进行评审。" & vbCrLf & vbCrLf & "待评审新增：2 份" & vbCrLf & "待评审共计：5 份", 5)

End Sub

Private Sub Command3_Click()

    Text11.Text = YPN.TrimText(Text11.Text)

End Sub

Private Sub Command4_Click()
    Form2.Show
End Sub

Private Sub Form_Load()

    Me.Text1.Text = YPN.GetHardDriveSerialNumber("D")
    Me.Text2.Text = YPN.GetHardDiskSerialNumber
    Me.Text3.Text = YPN.GetHardDiskModel
    Me.Text4.Text = YPN.MD5(Me.Text2.Text, 32)
    Me.Text5.Text = YPN.GetInitialAll(Me.Label5.Caption)
    Me.Text6.Text = YPN.GetInitialFirst(Me.Label6.Caption)
    Me.Label71.Caption = App.Path + "\Form1.frm"
    Me.Text7.Text = YPN.GetFileNameInPath(Me.Label71.Caption, True)
    Me.Text8.Text = YPN.GetSuffixInFileName(Me.Text7.Text)
    Me.Text9.Text = YPN.GetGUID()
    Me.Text12.Text = YPN.GetMonthBegin(Left(Label11.Caption, 10))
    Me.Text13.Text = YPN.GetMonthEnd(Left(Label12.Caption, 10))
            
            
            
    Dim i As Long
    cmb1(0).AddItem "Automatic"
    For i = 1 To 40
        cmb1(0).AddItem CStr(i)
    Next i
    cmb1(0).ListIndex = 0
    cmb1(1).AddItem "L - 7%"
    cmb1(1).AddItem "M - 15%"
    cmb1(1).AddItem "Q - 25%"
    cmb1(1).AddItem "H - 30%"
    cmb1(1).ListIndex = 1
    cmb1(2).AddItem "Automatic"
    For i = 0 To 7
        cmb1(2).AddItem CStr(i)
    Next i
    cmb1(2).ListIndex = 0
    cmb1(3).AddItem "ANSI"
    cmb1(3).AddItem "UTF-8"
    cmb1(3).ListIndex = 1
    
    Call Command1_Click
End Sub

Private Sub Image1_Click()

    If Not Image1.Picture Is Nothing Then
        Clipboard.Clear
        Clipboard.SetData Image1.Picture
        MsgBox "Successfully copied to clipboard"
    End If
    
End Sub

Private Sub Text10_GotFocus()
    On Error Resume Next
    Text10.SelStart = 0
    Text10.SelLength = Len(Text1.Text)
End Sub

Private Sub ypnButton_Shape1_Click()
    Form2.Show
End Sub
