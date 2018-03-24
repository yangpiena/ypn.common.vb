VERSION 5.00
Begin VB.Form FrmSSTab 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "FrmSSTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : FrmSSTab
' Author    : YPN
' Date      : 2018-03-24 22:51
' Purpose   : 用于定义SSTab的初始化
'             因为需要定义WithEvents类型的类变量，故只能放到Form中，不能在Module中定义
'---------------------------------------------------------------------------------------

Option Explicit
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const LF_FACESIZE = 32
Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To LF_FACESIZE) As Byte
End Type

Public Enum EMDTSTYLE
    DT_ACCEPT_DBCS = (&H20)
    DT_AGENT = (&H3)
    DT_BOTTOM = &H8
    DT_CALCRECT = &H400
    DT_CENTER = &H1
    DT_CHARSTREAM = 4
    DT_DISPFILE = 6
    DT_DISTLIST = (&H1)
    DT_EDITABLE = (&H2)
    DT_EDITCONTROL = &H2000
    DT_EDITCONTROL_CON = &H2000&
    DT_END_ELLIPSIS = &H8000
    DT_END_ELLIPSIS_CON = &H8000&
    DT_EXPANDTABS = &H40
    DT_EXTERNALLEADING = &H200
    DT_FOLDER = (&H1000000)
    DT_FOLDER_LINK = (&H2000000)
    DT_FOLDER_SPECIAL = (&H4000000)
    DT_FORUM = (&H2)
    DT_GLOBAL = (&H20000)
    DT_HIDEPREFIX = &H100000
    DT_INTERNAL = &H1000
    DT_LEFT = &H0
    DT_LOCAL = (&H30000)
    DT_MAILUSER = (&H0)
    DT_METAFILE = 5
    DT_MODIFIABLE = (&H10000)
    DT_MODIFYSTRING = &H10000
    DT_MULTILINE = (&H1)
    DT_NOCLIP = &H100
    DT_NOFULLWIDTHCHARBREAK = &H80000
    DT_NOPREFIX = &H800
    DT_NOT_SPECIFIC = (&H50000)
    DT_ORGANIZATION = (&H4)
    DT_PASSWORD_EDIT = (&H10)
    DT_PATH_ELLIPSIS = &H4000
    DT_PATH_ELLIPSIS_CON = &H4000&
    DT_PLOTTER = 0
    DT_PREFIXONLY = &H200000
    DT_PRIVATE_DISTLIST = (&H5)
    DT_RASCAMERA = 3
    DT_RASDISPLAY = 1
    DT_RASPRINTER = 2
    DT_REMOTE_MAILUSER = (&H6)
    DT_REQUIRED = (&H4)
    DT_RIGHT = &H2
    DT_RTLREADING = &H20000
    DT_RTLREADING_CON = &H20000
    DT_SET_IMMEDIATE = (&H8)
    DT_SET_SELECTION = (&H40)
    DT_SINGLELINE = &H20
    DT_TABSTOP = &H80
    DT_TOP = &H0
    DT_VCENTER = &H4
    DT_WAN = (&H40000)
    DT_WORD_ELLIPSIS = &H40000
    DT_WORD_ELLIPSIS_CON = &H40000
    DT_WORDBREAK = &H10
End Enum

Private Const OPAQUE        As Long = 2
Private Const TRANSPARENT   As Long = 1
Private Const PS_SOLID      As Long = 0

Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type tagINITCOMMONCONTROLSEX
    lngSize As Long
    lngICC As Long
End Type
Private Const ICC_USEREX_CLASSES = &H200
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagINITCOMMONCONTROLSEX) As Boolean

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal nSavedDC As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetBoundsRect Lib "gdi32" (ByVal hdc As Long, lprcBounds As RECT, ByVal flags As Long) As Long
Private Declare Function SetBoundsRect Lib "gdi32" (ByVal hdc As Long, lprcBounds As RECT, ByVal flags As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long

Private Const STATE_NORMAL      As Long = &H0
Private Const STATE_SELECTED    As Long = &H1
Private Const STATE_HOVER       As Long = &H2
Private Const STATE_PUSHED      As Long = &H4
Private Const STATE_FOCUSED     As Long = &H8
Private Const STATE_DISABLED    As Long = &H10

Private m_XPResBitmap     As IPictureDisp
Private m_hXPResDC        As Long
Private m_nXPResSaveDC    As Long

Private m_QQResBitmap     As IPictureDisp
Private m_hQQResDC        As Long
Private m_nQQResSaveDC    As Long

Private m_OfficeResBitmap     As IPictureDisp
Private m_hOfficeResDC        As Long
Private m_nOfficeResSaveDC    As Long

Private m_hFont             As Long
Private m_hBoldFont         As Long

Private WithEvents m_YPNSSTab As ClsYPNSSTab
Private m_SSTab               As SSTab
Private m_Style               As Integer


'---------------------------------------------------------------------------------------
' Procedure : FSSTabInit
' Author    : YPN
' Date      : 2018-03-24 22:59
' Purpose   : 初始化SSTab（重绘SSTab）
' Param     : i_SSTab             SSTab类型
'             i_Style （可选参数）样式类型：0 XP样式；1 QQ样式；2 Office样式
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Sub FSSTabInit(ByVal i_SSTab As Object, Optional ByVal i_Style As Integer = 0)
    If Not (TypeOf i_SSTab Is SSTab) Then Err.Raise 5
    
    Set m_SSTab = i_SSTab
    m_Style = i_Style
    
    Call loadResDC(m_SSTab)
    
    Set m_YPNSSTab = New ClsYPNSSTab
    Call m_YPNSSTab.Attach(m_SSTab)
    
    If (0 = m_Style) Then
        m_YPNSSTab.UpDownWidth = 16
        m_YPNSSTab.UpDownHeight = 16
        
        m_YPNSSTab.TabWidth = 100
        m_YPNSSTab.TabHeight = 20
    ElseIf (1 = m_Style) Then
        m_YPNSSTab.UpDownWidth = 16
        m_YPNSSTab.UpDownHeight = 16
        
        m_YPNSSTab.TabWidth = 100
        m_YPNSSTab.TabHeight = 24
    ElseIf (2 = m_Style) Then
        m_YPNSSTab.UpDownWidth = 16
        m_YPNSSTab.UpDownHeight = 16
        
        m_YPNSSTab.TabWidth = 100
        m_YPNSSTab.TabHeight = 27
    End If
    
    m_YPNSSTab.UpdateAll
End Sub

Private Function loadResDC(i_SSTab As SSTab) As Boolean
    loadResDC = False
    
    Dim hTmpDC As Long
    Dim i As Integer
    Dim lf As LOGFONT
    Dim fnByte() As Byte
    
    fnByte = StrConv(i_SSTab.Font.Name & vbNullString, vbFromUnicode)
    For i = 0 To UBound(fnByte)
        lf.lfFaceName(i + 1) = fnByte(i)
    Next i
    
    lf.lfHeight = -(i_SSTab.Font.Size + 2)
    lf.lfItalic = i_SSTab.Font.Italic
    lf.lfWeight = IIf(i_SSTab.Font.Bold, FW_BOLD, FW_NORMAL)
    lf.lfUnderline = i_SSTab.Font.Underline
    lf.lfStrikeOut = i_SSTab.Font.Strikethrough
    lf.lfCharSet = i_SSTab.Font.Charset
    m_hFont = CreateFontIndirect(lf)
    
    lf.lfWeight = FW_BOLD
    m_hBoldFont = CreateFontIndirect(lf)
    
    Set m_XPResBitmap = LoadResPicture("IDB_XPTHEME", vbResBitmap)
    Set m_QQResBitmap = LoadResPicture("IDB_QQTHEME", vbResBitmap)
    Set m_OfficeResBitmap = LoadResPicture("IDB_OFFICETHEME", vbResBitmap)
    If (m_XPResBitmap Is Nothing Or m_QQResBitmap Is Nothing Or m_OfficeResBitmap Is Nothing) Then
        Debug.Assert False
        Exit Function
    End If
    
    hTmpDC = GetDC(0)
    
    m_hXPResDC = CreateCompatibleDC(hTmpDC)
    m_nXPResSaveDC = SaveDC(m_hXPResDC)
    Call SelectObject(m_hXPResDC, m_XPResBitmap.Handle)
    
    m_hQQResDC = CreateCompatibleDC(hTmpDC)
    m_nQQResSaveDC = SaveDC(m_hQQResDC)
    Call SelectObject(m_hQQResDC, m_QQResBitmap.Handle)
    
    m_hOfficeResDC = CreateCompatibleDC(hTmpDC)
    m_nOfficeResSaveDC = SaveDC(m_hOfficeResDC)
    Call SelectObject(m_hOfficeResDC, m_OfficeResBitmap.Handle)
    
    Call ReleaseDC(0, hTmpDC)
End Function

Private Function destroyResDC() As Boolean
    destroyResDC = False
    
    If (m_hXPResDC <> 0) Then
        Call RestoreDC(m_hXPResDC, m_nXPResSaveDC)
        Call DeleteDC(m_hXPResDC)
        m_hXPResDC = 0
        m_nXPResSaveDC = 0
        destroyResDC = True
    End If
    
    If (m_hQQResDC <> 0) Then
        Call RestoreDC(m_hQQResDC, m_nQQResSaveDC)
        Call DeleteDC(m_hQQResDC)
        m_hQQResDC = 0
        m_nQQResSaveDC = 0
        destroyResDC = True
    End If
    
    If (m_hOfficeResDC <> 0) Then
        Call RestoreDC(m_hOfficeResDC, m_nOfficeResSaveDC)
        Call DeleteDC(m_hOfficeResDC)
        m_hOfficeResDC = 0
        m_nOfficeResSaveDC = 0
        destroyResDC = True
    End If
    
    If (m_hFont <> 0) Then
        Call DeleteObject(m_hFont)
        m_hFont = 0
    End If
    
    If (m_hBoldFont <> 0) Then
        Call DeleteObject(m_hBoldFont)
        m_hBoldFont = 0
    End If
    
    Set m_XPResBitmap = Nothing
End Function

Private Sub m_YPNSSTab_DrawBackGround(ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long)
    
    Dim hBrush As Long, hOldBrush  As Long
    
    Dim rcBackground As RECT
    Call SetRect(rcBackground, 0, 0, nWidth, nHeight)
    
    If m_Style = 0 Then
        ' XP样式
        Dim hPen As Long, hOldPen As Long
        Dim nOldBkMode As Long
        
        hPen = CreatePen(PS_SOLID, 1, &H9C9B91)
        hBrush = CreateSolidBrush(&HFEFCFC)
        hOldPen = SelectObject(hdc, hPen)
        hOldBrush = SelectObject(hdc, hBrush)
        nOldBkMode = SetBkMode(hdc, OPAQUE)
        Call FillRect(hdc, rcBackground, hBrush)
        Call Rectangle(hdc, 0, m_SSTab.TabHeight - 1, nWidth, nHeight)
        
        Call SelectObject(hdc, hOldPen)
        Call SelectObject(hdc, hOldBrush)
        
        Call DeleteObject(hPen)
        Call DeleteObject(hBrush)
        Call SetBkMode(hdc, nOldBkMode)
    ElseIf m_Style = 1 Then
        ' QQ样式
        hBrush = CreateSolidBrush(&HF9FCFB)
        Call FillRect(hdc, rcBackground, hBrush)
        Call gridBlt(hdc, 0, 0, nWidth, m_SSTab.TabHeight, m_hQQResDC, 0, 0, 1, 24, 0, 0, 0, 1)
        Call DeleteObject(hBrush)
    ElseIf m_Style = 2 Then
        ' Office样式
        hBrush = CreateSolidBrush(&HF6EAE1)
        Call FillRect(hdc, rcBackground, hBrush)
        Call gridBlt(hdc, 0, 0, nWidth, m_SSTab.TabHeight, m_hOfficeResDC, 0, 0, 1, 27, 0, 0, 0, 1)
        Call DeleteObject(hBrush)
    End If
End Sub

Private Sub m_YPNSSTab_DrawTab(ByVal nTab As Long, ByVal nState As Long, ByVal hdc As Long, ByVal nleft As Long, ByVal nTop As Long, ByVal nWidth As Long, ByVal nHeight As Long)
    Dim rcTab As RECT, rcFocus As RECT
    Dim nOldBkMode As Long
    Dim hOldFont As Long
    Dim dwOldTextColor As Long
    Dim bSelected As Boolean, bFocused As Boolean, bHover As Boolean, bDisabled As Boolean
    Call SetRect(rcTab, nleft, nTop, nleft + nWidth, nTop + nHeight)
    nOldBkMode = SetBkMode(hdc, TRANSPARENT)
    
    bSelected = (nState And STATE_SELECTED)
    bFocused = (nState And STATE_FOCUSED)
    bHover = (nState And STATE_HOVER)
    bDisabled = (nState And STATE_DISABLED)
    
    If m_Style = 0 Then
        ' 增加间隙
        rcTab.Right = rcTab.Right - 2
        ' XP样式
        If (bSelected) Then
            Call gridBlt(hdc, rcTab.Left, rcTab.Top, rcTab.Right - rcTab.Left, rcTab.Bottom - rcTab.Top, m_hXPResDC, 14, 0, 7, 20, 3, 3, 3, 3, RGB(255, 0, 255))
        ElseIf (bHover And Not bDisabled) Then
            Call gridBlt(hdc, rcTab.Left, rcTab.Top, rcTab.Right - rcTab.Left, rcTab.Bottom - rcTab.Top, m_hXPResDC, 7, 0, 7, 20, 3, 3, 3, 3, RGB(255, 0, 255))
        Else
            Call gridBlt(hdc, rcTab.Left, rcTab.Top, rcTab.Right - rcTab.Left, rcTab.Bottom - rcTab.Top, m_hXPResDC, 0, 0, 7, 20, 3, 3, 3, 3, RGB(255, 0, 255))
        End If
        
        If (bFocused) Then
            Call SetRect(rcFocus, rcTab.Left + 3, rcTab.Top + 4, rcTab.Right - 3, rcTab.Bottom - 2)
            Call DrawFocusRect(hdc, rcFocus)
        End If
        
        ' 让文字下沉一点
        rcTab.Top = rcTab.Top + 2
    ElseIf m_Style = 1 Then
        ' QQ样式
        If (bSelected) Then
            Call gridBlt(hdc, rcTab.Left, rcTab.Top, rcTab.Right - rcTab.Left, rcTab.Bottom - rcTab.Top, m_hQQResDC, 7, 0, 3, 24, 1, 3, 1, 0)
        ElseIf (nTab < m_SSTab.Tab) Then
            Call gridBlt(hdc, rcTab.Left, rcTab.Top, rcTab.Right - rcTab.Left, rcTab.Bottom - rcTab.Top, m_hQQResDC, IIf(bHover And Not bDisabled, 3, 0) + 1, 0, 2, 24, 1, 0, 0, 0)
        Else
            Call gridBlt(hdc, rcTab.Left, rcTab.Top, rcTab.Right - rcTab.Left, rcTab.Bottom - rcTab.Top, m_hQQResDC, IIf(bHover And Not bDisabled, 3, 0) + 2, 0, 2, 24, 0, 0, 1, 0)
        End If
        
        ' 让文字下沉一点
        rcTab.Top = rcTab.Top + 2
    ElseIf m_Style = 2 Then
        ' Office样式
        If (bSelected) Then
            Call gridBlt(hdc, rcTab.Left, rcTab.Top, rcTab.Right - rcTab.Left, rcTab.Bottom - rcTab.Top, m_hOfficeResDC, IIf(bHover, 26, 13), 0, 13, 27, 6, 5, 6, 2)
        ElseIf (bHover And Not bDisabled) Then
            Call gridBlt(hdc, rcTab.Left, rcTab.Top, rcTab.Right - rcTab.Left, rcTab.Bottom - rcTab.Top, m_hOfficeResDC, 0, 0, 13, 27, 6, 5, 6, 2)
        Else
            Call gridBlt(hdc, rcTab.Left, rcTab.Top, rcTab.Right - rcTab.Left, rcTab.Bottom - rcTab.Top, m_hOfficeResDC, 0, 0, 1, 27, 0, 0, 0, 1)
        End If
        
        ' 让文字下沉一点
        rcTab.Top = rcTab.Top + 2
    End If
    
    hOldFont = SelectObject(hdc, IIf(bSelected, m_hBoldFont, m_hFont))
    dwOldTextColor = GetTextColor(hdc)
    If (m_SSTab.TabEnabled(nTab) = False) Then Call SetTextColor(hdc, RGB(150, 150, 150))
    
    Call DrawText(hdc, m_SSTab.TabCaption(nTab), -1, rcTab, DT_CENTER Or DT_VCENTER Or DT_END_ELLIPSIS_CON Or DT_SINGLELINE)
    '   Call SetBkMode(hdc, nOldBkMode)
    Call SelectObject(hdc, hOldFont)
    Call SetTextColor(hdc, dwOldTextColor)
End Sub

Private Sub m_YPNSSTab_DrawUpDown(ByVal bUpButton As Boolean, ByVal nState As Long, ByVal hdc As Long, ByVal nleft As Long, ByVal nTop As Long, ByVal nWidth As Long, ByVal nHeight As Long)
    Dim rcUpDown As RECT
    Dim bSelected As Boolean, bFocused As Boolean, bHover As Boolean, bDisabled As Boolean
    
    Call SetRect(rcUpDown, nleft, nTop, nleft + nWidth, nTop + nHeight)
    
    bSelected = (nState And STATE_SELECTED)
    bFocused = (nState And STATE_FOCUSED)
    bHover = (nState And STATE_HOVER)
    bDisabled = (nState And STATE_DISABLED)
    
    If (bUpButton) Then
        If (bDisabled) Then
            Call gridBlt(hdc, rcUpDown.Left, rcUpDown.Top, rcUpDown.Right - rcUpDown.Left, rcUpDown.Bottom - rcUpDown.Top, m_hXPResDC, 53, 0, 16, 16, 1, 1, 1, 1)
        Else
            Call gridBlt(hdc, rcUpDown.Left, rcUpDown.Top, rcUpDown.Right - rcUpDown.Left, rcUpDown.Bottom - rcUpDown.Top, m_hXPResDC, IIf(bHover, 37, 21), 0, 16, 16, 1, 1, 1, 1)
        End If
    Else
        If (bDisabled) Then
            Call gridBlt(hdc, rcUpDown.Left, rcUpDown.Top, rcUpDown.Right - rcUpDown.Left, rcUpDown.Bottom - rcUpDown.Top, m_hXPResDC, 101, 0, 16, 16, 1, 1, 1, 1)
        Else
            Call gridBlt(hdc, rcUpDown.Left, rcUpDown.Top, rcUpDown.Right - rcUpDown.Left, rcUpDown.Bottom - rcUpDown.Top, m_hXPResDC, IIf(bHover, 85, 69), 0, 16, 16, 1, 1, 1, 1)
        End If
    End If
End Sub

'九宫格绘图 (中间伸展，边框不变)
Private Function gridBlt(ByVal hDestDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, _
    ByVal hSrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, _
    Optional ByVal gX1 As Long = 0, Optional ByVal gY1 As Long = 0, Optional ByVal gX2 As Long = 0, Optional ByVal gY2 As Long = 0, _
    Optional ByVal MaskColor As Variant) As Long
    If dstWidth = 0 Or dstHeight = 0 Or srcWidth = 0 Or srcHeight = 0 Then Exit Function
    
    Dim hTmpDC As Long
    Dim hMemDC As Long
    Dim hMemBitmap As Long, hOldMemBitmap As Long
    hTmpDC = GetDC(0)
    hMemDC = CreateCompatibleDC(hTmpDC)
    hMemBitmap = CreateCompatibleBitmap(hTmpDC, dstWidth, dstHeight)
    hOldMemBitmap = SelectObject(hMemDC, hMemBitmap)
    
    If gX1 <= 0 And gX2 <= 0 And gY1 <= 0 And gY2 <= 0 Then
        StretchBlt hMemDC, 0, 0, dstWidth, dstHeight, hSrcDC, SrcX + gX1, SrcY + gY1, srcWidth - gX2, srcHeight - gY2, vbSrcCopy
    Else
        If gX1 > 0 And gY1 > 0 Then '左上角
            BitBlt hMemDC, 0, 0, gX1, gY1, hSrcDC, SrcX, SrcY, vbSrcCopy
        End If
        If gX2 > 0 And gY1 > 0 Then '右上角
            BitBlt hMemDC, dstWidth - gX2, 0, gX2, gY1, hSrcDC, SrcX + srcWidth - gX2, SrcY, vbSrcCopy
        End If
        If gX1 > 0 And gY2 > 0 Then '左下角
            BitBlt hMemDC, 0, dstHeight - gY2, gX1, gY2, hSrcDC, SrcX, SrcY + srcHeight - gY2, vbSrcCopy
        End If
        If gX2 > 0 And gY2 > 0 Then '右下角
            BitBlt hMemDC, dstWidth - gX2, dstHeight - gY2, gX2, gY2, hSrcDC, SrcX + srcWidth - gX2, SrcY + srcHeight - gY2, vbSrcCopy
        End If
        If gX1 > 0 Then '左边框
            StretchBlt hMemDC, 0, gY1, gX1, dstHeight - gY1 - gY2, hSrcDC, SrcX, SrcY + gY1, gX1, srcHeight - gY1 - gY2, vbSrcCopy
        End If
        If gX2 > 0 Then '右边框
            StretchBlt hMemDC, dstWidth - gX2, gY1, gX2, dstHeight - gY1 - gY2, hSrcDC, SrcX + srcWidth - gX2, SrcY + gY1, gX2, srcHeight - gY1 - gY2, vbSrcCopy
        End If
        If gY1 > 0 Then '上边框
            StretchBlt hMemDC, gX1, 0, dstWidth - gX1 - gX2, gY1, hSrcDC, SrcX + gX1, SrcY, srcWidth - gX1 - gX2, gY1, vbSrcCopy
        End If
        If gY2 > 0 Then '下边框
            StretchBlt hMemDC, gX1, dstHeight - gY2, dstWidth - gX1 - gX2, gY2, hSrcDC, SrcX + gX1, SrcY + srcHeight - gY2, srcWidth - gX1 - gX2, gY2, vbSrcCopy
        End If
        '中间的伸展部分
        StretchBlt hMemDC, gX1, gY1, dstWidth - gX1 - gX2, dstHeight - gY1 - gY2, hSrcDC, SrcX + gX1, SrcY + gY1, srcWidth - gX1 - gX2, srcHeight - gY1 - gY2, vbSrcCopy
    End If
    If IsMissing(MaskColor) Then
        gridBlt = BitBlt(hDestDC, dstX, dstY, dstWidth, dstHeight, hMemDC, 0, 0, vbSrcCopy)
    Else
        gridBlt = TransparentBlt(hDestDC, dstX, dstY, dstWidth, dstHeight, hMemDC, 0, 0, dstWidth, dstHeight, CLng(Val(MaskColor)))
    End If
    
    Call SelectObject(hMemDC, hOldMemBitmap)
    Call DeleteDC(hMemDC)
    Call DeleteObject(hMemBitmap)
    Call ReleaseDC(0, hTmpDC)
End Function
