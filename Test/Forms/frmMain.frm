VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Main"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13845
   LinkTopic       =   "frmMain"
   ScaleHeight     =   8205
   ScaleWidth      =   13845
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command7 
      Caption         =   "�����ʼ�"
      Height          =   375
      Left            =   4680
      TabIndex        =   37
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "FTP����"
      Height          =   375
      Left            =   3240
      TabIndex        =   36
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "JSON"
      Height          =   375
      Left            =   1680
      TabIndex        =   35
      Top             =   120
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1815
      Left            =   7560
      TabIndex        =   34
      Top             =   4800
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3201
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SSTab"
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
      Caption         =   "ȥ�����߿ո��뻻�з�"
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
      Text            =   "frmMain.frx":001C
      Top             =   3360
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "������Ļ���½���Ϣ��"
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
      Text            =   "frmMain.frx":0023
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
      Caption         =   "2018-04-20����ĩ��"
      Height          =   180
      Left            =   240
      TabIndex        =   32
      Top             =   5370
      Width           =   1620
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "2018-04-20���³���"
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
      Caption         =   "��ά�����ɲ��ԣ�"
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
      Caption         =   "GUID��"
      Height          =   180
      Left            =   240
      TabIndex        =   18
      Top             =   4410
      Width           =   540
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "���ļ�����ȡ��׺����"
      Height          =   180
      Left            =   240
      TabIndex        =   16
      Top             =   3930
      Width           =   1800
   End
   Begin VB.Label Label71 
      AutoSize        =   -1  'True
      Caption         =   "��ȫ·����ȡ�ļ�����"
      Height          =   180
      Left            =   4920
      TabIndex        =   14
      Top             =   3480
      Width           =   1800
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "��ȫ·����ȡ�ļ�����"
      Height          =   180
      Left            =   240
      TabIndex        =   13
      Top             =   3450
      Width           =   1800
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "��һ����������ĸ��"
      Height          =   180
      Left            =   240
      TabIndex        =   11
      Top             =   2970
      Width           =   1620
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "���к�������ĸ��"
      Height          =   180
      Left            =   240
      TabIndex        =   9
      Top             =   2490
      Width           =   1440
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "MD5�������������кţ�"
      Height          =   180
      Left            =   240
      TabIndex        =   7
      Top             =   2010
      Width           =   1890
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "�������ͺţ�"
      Height          =   180
      Left            =   240
      TabIndex        =   5
      Top             =   1530
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "���������кţ�"
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   1050
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�߼������кţ�"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1260
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Declare Function WideCharToMultiByte Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpWideCharStr As Any, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByRef lpDefaultChar As Any, ByRef lpUsedDefaultChar As Any) As Long
Private Const CP_UTF8 As Long = 65001


Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
'���õĴ���ҳ��
Const cpUTF8 = 65001
Const cpGB2312 = 936
Const cpGB18030 = 54936
Const cpUTF7 = 65000


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
    
    Call YPN.ShowMessage("WRP-PDP", Me.Icon, "��ǰ����׼��ϵͳ��Ϣ", "��ͬ����", "�����µĺ�ͬ��Ҫ����" & vbCrLf & "�뼰ʱ���롾��ͬ�������󡿽�������" & vbCrLf & vbCrLf & "������������2 ��" & vbCrLf & "�����󹲼ƣ�5 ��", 5)
    
End Sub

Private Sub Command3_Click()
    
    Text11.Text = YPN.TrimText(Text11.Text)
    
End Sub

Private Sub Command4_Click()
    frmSSTab.Show
End Sub

Private Sub Command5_Click()
    frmJSON.Show
End Sub

Private Sub Command6_Click()
    
    Dim e As Object
    Dim v_fileName As String
    
    v_fileName = "1.2���PCE�ۺ����ü����׼�����Ŀ/��ʽ��ͬ/1.2���PCE�ۺ����ü����׼�����Ŀ_2��_2̨_2018-04-09.xls"
    
'    Set e = CreateObject("MSScriptControl.ScriptControl")
'    e.Language = "javascript"
'    Dim d As String
'    d = e.Eval("encodeURI('΢������')") '����javascript�ű��ĺ���
'    MsgBox d
'    MsgBox e.Eval("decodeURI('" & d & "')")
    
'    MsgBox v_fileName
'    v_fileName = e.Eval("encodeURI('" & v_fileName & "')")
'    MsgBox v_fileName
    
    MsgBox MultiByteToUTF16(UTF16ToMultiByte(v_fileName, cpUTF8), cpUTF8)
    MsgBox UTF16ToMultiByte(v_fileName, cpUTF8)
    
    
    Call ModFTPUtils.FTPFileDownload("10.1.50.45", "xx", "xx", LoadAsUTF8(v_fileName), "D:\WRP\����ɭװ������ҵ��ǰ����׼��ϵͳ\xsgl\XSGL\Files\1.2���PCE�ۺ����ü����׼�����Ŀ_2��_2̨_2018-04-09.xls", False)
    
End Sub

'����Ҫ����  Microsoft ActiveX Data Objects 2.8����������ͨ�÷����������ģ����
Public Sub SaveAsUTF8(ByVal Text As String, ByVal FileName As String)
  Dim oStream As ADODB.Stream

  Set oStream = New ADODB.Stream
  oStream.Open
  oStream.Charset = "UTF-8"
  oStream.Type = adTypeText
  oStream.WriteText Text
  oStream.SaveToFile FileName, adSaveCreateOverWrite
  oStream.Close
End Sub

Public Function LoadAsUTF8(ByVal FileName As String) As String
  Dim oStream As ADODB.Stream

  Set oStream = New ADODB.Stream
  oStream.Open
  oStream.Charset = "UTF-8"
  oStream.LoadFromFile FileName

  LoadAsUTF8 = oStream.ReadText()

  oStream.Close
End Function

Function MultiByteToUTF16(UTF8() As Byte, CodePage As Long) As String
    Dim bufSize As Long
    bufSize = MultiByteToWideChar(CodePage, 0&, UTF8(0), UBound(UTF8) + 1, 0, 0)
    MultiByteToUTF16 = Space(bufSize)
    MultiByteToWideChar CodePage, 0&, UTF8(0), UBound(UTF8) + 1, StrPtr(MultiByteToUTF16), bufSize
End Function

Function UTF16ToMultiByte(UTF16 As String, CodePage As Long) As Byte()
    Dim bufSize As Long
    Dim arr() As Byte
    bufSize = WideCharToMultiByte(CodePage, 0&, StrPtr(UTF16), Len(UTF16), 0, 0, 0, 0)
    ReDim arr(bufSize - 1)
    WideCharToMultiByte CodePage, 0&, StrPtr(UTF16), Len(UTF16), arr(0), bufSize, 0, 0
    UTF16ToMultiByte = arr
End Function


Private Sub Command7_Click()

    Dim v_body As String
    
    v_body = "<table border=0>" + "<tr><td>��Ŀ���ƣ�</td><td>" + App.Path + "</td></tr><tr><td>��ҵ���</td><td>" + App.EXEName + "</td></tr></table>"
    
    Call ModPublic.MSendEmail("smtp.qiye.163.com", "system@wzyb.com.cn", "WZYBwzyb9114", "yd@wzyb.com.cn", "YPN����VB�����ʼ�", v_body)
    Call ModPublic.MSendHTMLEmail("smtp.qiye.163.com", "system@wzyb.com.cn", "WZYBwzyb9114", "yd@wzyb.com.cn", "YPN����VB�����ʼ�", v_body, "D:\YPNCloud\YPN.Git\ypn.common.vb\Test\TestDLL.vbg")
    
End Sub

Private Sub Form_Load()
    
    Me.Text1.Text = YPN.GetHardDriveSerialNumber("D")
    Me.Text2.Text = YPN.GetHardDiskSerialNumber
    Me.Text3.Text = YPN.GetHardDiskModel
    Me.Text4.Text = YPN.MD5(Me.Text2.Text, 32)
    Me.Text5.Text = YPN.GetInitialAll(Me.Label5.Caption)
    Me.Text6.Text = YPN.GetInitialFirst(Me.Label6.Caption)
    Me.Label71.Caption = App.Path + "\frmMain.frm"
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
    
    Call YPN.SSTabInit(Me.SSTab1, 2)
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
