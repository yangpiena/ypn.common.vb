VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsYPNSSTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : ClsCustomSSTab
' Author    : YPN
' Date      : 2018-03-24 00:37
' Purpose   : 重绘SSTab
' 改编自    ：
'             文  件名：    CCustomSSTab.cls
'             功    能：    重绘SSTab的子类化类模块
'             创建时间：    2015-10-05 09:59:09
'             最后修改：    2015-10-05 09:59:18
'             声    明：    自由使用、传播， 但请保留原作者信息
'             联 系 QQ：    981666572
'             电子邮箱：    owhao@126.com
'             工作原理：
'                           子类化 WM_PAINT消息对控件绘画自定义的外观
'                           子类化 WM_LBUTTONDOWN 等消息，实现响应自定义外观的标签页点击事件
'---------------------------------------------------------------------------------------

Option Explicit
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left                        As Long
    Top                         As Long
    Right                       As Long
    Bottom                      As Long
End Type

Private Type PAINTSTRUCT
    hdc                         As Long
    fErase                      As Long
    rcPaint                     As RECT
    fRestore                    As Long
    fIncUpdate                  As Long
    rgbReserved(32)             As Byte
End Type

Private Type BITMAP '14 bytes
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
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

Private Type TRACKMOUSEEVENTTYPE
    cbSize                          As Long
    dwFlags                         As Long
    hwndTrack                       As Long
    dwHoverTime                     As Long
End Type

Private Type tSubData
    hWnd                            As Long
    nAddrSub                        As Long
    nAddrOrig                       As Long
    nMsgCntA                        As Long
    nMsgCntB                        As Long
    aMsgTblA()                      As Long
    aMsgTblB()                      As Long
End Type

Private Enum eMsgWhen
    [MSG_AFTER] = &H1
    [MSG_BEFORE] = &H2
    [MSG_BEFORE_AND_AFTER] = MSG_AFTER Or MSG_BEFORE
End Enum

Private Const ALL_MESSAGES          As Long = &HFFFF
Private Const CODE_LEN              As Long = &HC5
Private Const GWL_WNDPROC           As Long = &HFFFC
Private Const PATCH_04              As Long = &H58
Private Const PATCH_05              As Long = &H5D
Private Const PATCH_08              As Long = &H84
Private Const PATCH_09              As Long = &H89

Private Const WM_PAINT              As Long = &HF
Private Const WM_SIZE               As Long = &H5
Private Const WM_GETTEXT            As Long = &HD
Private Const WM_SETFOCUS           As Long = &H7
Private Const WM_KILLFOCUS          As Long = &H8
Private Const WM_ERASEBKGND         As Long = &H14
Private Const WM_NCDESTROY          As Long = &H82
Private Const WM_MOUSEHOVER         As Long = &H2A1
Private Const WM_MOUSELEAVE         As Long = &H2A3
Private Const WM_MOUSEMOVE          As Long = &H200
Private Const WM_LBUTTONDBLCLK      As Long = &H203
Private Const WM_LBUTTONDOWN        As Long = &H201
Private Const WM_LBUTTONUP          As Long = &H202
Private Const WM_MBUTTONDBLCLK      As Long = &H209
Private Const WM_MBUTTONDOWN        As Long = &H207
Private Const WM_MBUTTONUP          As Long = &H208
Private Const WM_RBUTTONDBLCLK      As Long = &H206
Private Const WM_RBUTTONDOWN        As Long = &H204
Private Const WM_RBUTTONUP          As Long = &H205

Private Const BLACK_BRUSH           As Long = 4
Private Const NULL_BRUSH            As Long = 5

Private Const COLOR_WINDOW          As Long = 5
Private Const COLOR_BTNFACE         As Long = 15

Private Const OPAQUE                As Long = 2
Private Const TRANSPARENT           As Long = 1

Private Const DCB_ACCUMULATE        As Long = &H2
Private Const DCB_DISABLE           As Long = &H8
Private Const DCB_ENABLE            As Long = &H4
Private Const DCB_RESET             As Long = &H1
Private Const DCB_SET               As Long = (DCB_RESET Or DCB_ACCUMULATE)

Private Const PS_SOLID              As Long = 0
Private Const PS_DASH               As Long = 1         '  -------
Private Const PS_DOT                As Long = 2         '  .......
Private Const PS_DASHDOT            As Long = 3         '  _._._._
Private Const PS_DASHDOTDOT         As Long = 4         '  _.._.._
Private Const PS_NULL               As Long = 5
Private Const PS_INSIDEFRAME        As Long = 6
Private Const PS_USERSTYLE          As Long = 7
Private Const PS_ALTERNATE          As Long = 8

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Private Declare Function IsRectEmpty Lib "user32" (lpRect As RECT) As Long
Private Declare Function InvalidateRectNull Lib "user32" Alias "InvalidateRect" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long

Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function GetBoundsRect Lib "gdi32" (ByVal hdc As Long, lprcPaint As RECT, ByVal flags As Long) As Long
Private Declare Function SetBoundsRect Lib "gdi32" (ByVal hdc As Long, lprcBounds As RECT, ByVal flags As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal nSavedDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

Private Declare Function TrackMouseEvent2 Lib "comctl32.dll" Alias "_TrackMouseEvent" (ByRef lpEventTrack As TRACKMOUSEEVENTTYPE) As Long ' Win95 w/ IE 3.0
Private Declare Function TrackMouseEvent Lib "user32.dll" (ByRef lpEventTrack As TRACKMOUSEEVENTTYPE) As Long ' Win98 or later

Private sc_aSubData()           As tSubData
Private sc_aBuf(1 To CODE_LEN)  As Byte
Private sc_pCWP                 As Long
Private sc_pEbMode              As Long
Private sc_pSWL                 As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''
Private m_bIsTrackHandler32     As Boolean
Private m_bIsTracked            As Boolean

Private m_hMemDC                As Long
Private m_hMemBitmap            As Long
Private m_nMemDCState           As Long

Private m_TabWidth              As Long
Private m_TabHeight             As Long

Private m_UpDownWidth           As Long
Private m_UpDownHeight          As Long

Private Const TAB_NULL          As Long = -1
Private Const TAB_UP            As Long = -2
Private Const TAB_DOWN          As Long = -3

Private m_HoverTab              As Long
Private m_TabOffsetX            As Long

Private WithEvents m_SSTab      As SSTab
Attribute m_SSTab.VB_VarHelpID = -1
Private m_hSSTabWnd             As Long

Public Event DrawBackGround(ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long)

' nState的取值如下
Private Const STATE_NORMAL      As Long = &H0
Private Const STATE_SELECTED    As Long = &H1
Private Const STATE_HOVER       As Long = &H2
Private Const STATE_PUSHED      As Long = &H4
Private Const STATE_FOCUSED     As Long = &H8
Private Const STATE_DISABLED    As Long = &H10

Public Event DrawTab(ByVal nTab As Long, ByVal nState As Long, ByVal hdc As Long, ByVal nleft As Long, ByVal nTop As Long, ByVal nWidth As Long, ByVal nHeight As Long)
Public Event DrawUpDown(ByVal bUpButton As Boolean, ByVal nState As Long, ByVal hdc As Long, ByVal nleft As Long, ByVal nTop As Long, ByVal nWidth As Long, ByVal nHeight As Long)

'这个过程一定要放在类模块的最顶部，否则出错
Public Sub ZSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lhWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
    Dim Button As Long
    Dim Point As POINTAPI
    
    Select Case uMsg
    Case WM_NCDESTROY
        Detach
        Exit Sub
        
    Case WM_PAINT
        lReturn = onPaint(hWnd, uMsg, wParam, lParam, bHandled)
        Exit Sub
        
    Case WM_ERASEBKGND
        bHandled = True
        lReturn = 1
        Exit Sub
        
    Case WM_SIZE
        Call onSize(wParam, loWORD(lParam), hiWORD(lParam))
        
    Case WM_MOUSEMOVE
        bHandled = True
        lReturn = 0
        If m_bIsTracked = False Then
            m_bIsTracked = True
            Call trackMouseTracking(m_hSSTabWnd)
        End If
        Button = wParam
        Point.x = loWORD(lParam)
        Point.y = hiWORD(lParam)
        Call onMouseMove(Button, Point)
        Exit Sub
        
    Case WM_MOUSEHOVER:
        m_bIsTracked = True
        Call trackMouseTracking(m_hSSTabWnd)
        Call onMouseHover
        Exit Sub
        
    Case WM_MOUSELEAVE
        m_bIsTracked = False
        Call onMouseLeave
        Exit Sub
        
    Case WM_LBUTTONDBLCLK, WM_MBUTTONDBLCLK, WM_RBUTTONDBLCLK
        bHandled = True
        lReturn = 0
        Point.x = loWORD(lParam)
        Point.y = hiWORD(lParam)
        If uMsg = WM_LBUTTONDBLCLK Then Button = 1
        If uMsg = WM_RBUTTONDBLCLK Then Button = 2
        If uMsg = WM_MBUTTONDBLCLK Then Button = 3
        Call onMouseDBClick(Button, Point)
        Exit Sub
        
    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN
        bHandled = True
        lReturn = 0
        Point.x = loWORD(lParam)
        Point.y = hiWORD(lParam)
        If uMsg = WM_LBUTTONDOWN Then Button = 1
        If uMsg = WM_RBUTTONDOWN Then Button = 2
        If uMsg = WM_MBUTTONDOWN Then Button = 3
        m_SSTab.SetFocus
        Call onMouseDown(Button, Point)
        Exit Sub
        
    Case WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
        bHandled = True
        lReturn = 0
        Point.x = loWORD(lParam)
        Point.y = hiWORD(lParam)
        If uMsg = WM_LBUTTONUP Then Button = 1
        If uMsg = WM_RBUTTONUP Then Button = 2
        If uMsg = WM_MBUTTONUP Then Button = 3
        Call onMouseUp(Button, Point)
        Exit Sub
        
    Case WM_SETFOCUS, WM_KILLFOCUS
        Call updateTab(m_SSTab.Tab)
        Exit Sub
        
    End Select
End Sub

Private Sub subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    With sc_aSubData(zIdx(lng_hWnd))
        If (When And eMsgWhen.MSG_BEFORE) Then
            Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If (When And eMsgWhen.MSG_AFTER) Then
            Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

Private Sub subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    
    With sc_aSubData(zIdx(lng_hWnd))
        If (When And eMsgWhen.MSG_BEFORE) Then
            Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If (When And eMsgWhen.MSG_AFTER) Then
            Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

Private Function subclass_InIDE() As Boolean
    Debug.Assert zSetTrue(subclass_InIDE)
End Function

Private Function subclass_Start(ByVal lng_hWnd As Long) As Long
    
    Dim i                        As Long
    Dim j                        As Long
    Dim nSubIdx                  As Long
    Dim sSubCode                 As String
    
    Const GMEM_FIXED             As Long = 0
    Const PAGE_EXECUTE_READWRITE As Long = &H40&
    Const PATCH_01               As Long = 18
    Const PATCH_02               As Long = 68
    Const PATCH_03               As Long = 78
    Const PATCH_06               As Long = 116
    Const PATCH_07               As Long = 121
    Const PATCH_0A               As Long = 186
    Const FUNC_CWP               As String = "CallWindowProcA"
    Const FUNC_EBM               As String = "EbMode"
    Const FUNC_SWL               As String = "SetWindowLongA"
    Const MOD_USER               As String = "user32"
    Const MOD_VBA5               As String = "vba5"
    Const MOD_VBA6               As String = "vba6"
    
    If (sc_aBuf(1) = 0) Then
        
        sSubCode = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D0000005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D000000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E32F09C978078B450CF2AF75248D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF501CC3"
        i = 1
        Do While j < CODE_LEN
            j = j + 1
            sc_aBuf(j) = CByte("&H" & Mid$(sSubCode, i, 2))
            i = i + 2
        Loop
        
        'Get API function addresses
        If (subclass_InIDE) Then
            sc_aBuf(16) = &H90
            sc_aBuf(17) = &H90
            sc_pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)
            If (sc_pEbMode = 0) Then
                sc_pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)
            End If
        End If
        
        Call zPatchVal(VarPtr(sc_aBuf(1)), PATCH_0A, ObjPtr(Me))
        
        sc_pCWP = zAddrFunc(MOD_USER, FUNC_CWP)
        sc_pSWL = zAddrFunc(MOD_USER, FUNC_SWL)
        ReDim sc_aSubData(0 To 0) As tSubData
        
    Else
        nSubIdx = zIdx(lng_hWnd, True)
        If (nSubIdx = -1) Then
            nSubIdx = UBound(sc_aSubData()) + 1
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData
        End If
        
        subclass_Start = nSubIdx
    End If
    
    With sc_aSubData(nSubIdx)
        
        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)
        Call VirtualProtect(ByVal .nAddrSub, CODE_LEN, PAGE_EXECUTE_READWRITE, i)
        Call RtlMoveMemory(ByVal .nAddrSub, sc_aBuf(1), CODE_LEN)
        
        .hWnd = lng_hWnd
        .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)
        
        Call zPatchRel(.nAddrSub, PATCH_01, sc_pEbMode)
        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)
        Call zPatchRel(.nAddrSub, PATCH_03, sc_pSWL)
        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)
        Call zPatchRel(.nAddrSub, PATCH_07, sc_pCWP)
    End With
End Function

Private Sub subclass_StopAll()
    
    Dim i As Long
    
    i = UBound(sc_aSubData())
    Do While i >= 0
        With sc_aSubData(i)
            If (.hWnd <> 0) Then
                Call subclass_Stop(.hWnd)
            End If
        End With
        
        i = i - 1
    Loop
End Sub

Private Sub subclass_Stop(ByVal lng_hWnd As Long)
    
    With sc_aSubData(zIdx(lng_hWnd))
        Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)
        Call zPatchVal(.nAddrSub, PATCH_05, 0)
        Call zPatchVal(.nAddrSub, PATCH_09, 0)
        Call GlobalFree(.nAddrSub)
        .hWnd = 0
        .nMsgCntB = 0
        .nMsgCntA = 0
        Erase .aMsgTblB
        Erase .aMsgTblA
    End With
End Sub


Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    
    Dim nEntry  As Long
    Dim nOff1   As Long
    Dim nOff2   As Long
    
    If (uMsg = ALL_MESSAGES) Then
        nMsgCnt = ALL_MESSAGES
    Else
        Do While nEntry < nMsgCnt
            nEntry = nEntry + 1
            
            If (aMsgTbl(nEntry) = 0) Then
                aMsgTbl(nEntry) = uMsg
                Exit Sub
            ElseIf (aMsgTbl(nEntry) = uMsg) Then
                Exit Sub
            End If
        Loop
        
        nMsgCnt = nMsgCnt + 1
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long
        aMsgTbl(nMsgCnt) = uMsg
    End If
    
    If (When = eMsgWhen.MSG_BEFORE) Then
        nOff1 = PATCH_04
        nOff2 = PATCH_05
    Else
        nOff1 = PATCH_08
        nOff2 = PATCH_09
    End If
    
    If (uMsg <> ALL_MESSAGES) Then
        Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))
    End If
    Call zPatchVal(nAddr, nOff2, nMsgCnt)
End Sub

Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc
End Function

Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    
    Dim nEntry As Long
    
    If (uMsg = ALL_MESSAGES) Then
        nMsgCnt = 0
        If When = eMsgWhen.MSG_BEFORE Then
            nEntry = PATCH_05
        Else
            nEntry = PATCH_09
        End If
        Call zPatchVal(nAddr, nEntry, 0)
    Else
        Do While nEntry < nMsgCnt
            nEntry = nEntry + 1
            If (aMsgTbl(nEntry) = uMsg) Then
                aMsgTbl(nEntry) = 0
                Exit Do
            End If
        Loop
    End If
End Sub

Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
    
    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0
        With sc_aSubData(zIdx)
            If (.hWnd = lng_hWnd) Then
                If (Not bAdd) Then
                    Exit Function
                End If
            ElseIf (.hWnd = 0) Then
                If (bAdd) Then
                    Exit Function
                End If
            End If
        End With
        zIdx = zIdx - 1
    Loop
    
    If (Not bAdd) Then
        Debug.Assert False
    End If
    
End Function

Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
    zSetTrue = True
    bValue = True
End Function

Private Sub Class_Initialize()
    m_bIsTrackHandler32 = IsFunctionSupported("TrackMouseEvent", "User32")
    
    m_TabWidth = m_TabHeight = 0
    m_UpDownWidth = m_UpDownHeight = 0
    
    m_TabOffsetX = 0
End Sub

Private Sub Class_Terminate()
    Call Detach
End Sub

Private Function IsFunctionSupported(sFunction As String, sModule As String) As Boolean
    Dim hModule As Long
    hModule = GetModuleHandleA(sModule)
    If (hModule = 0) Then
        hModule = LoadLibrary(sModule)
    End If
    If (hModule) Then
        If (GetProcAddress(hModule, sFunction)) Then
            IsFunctionSupported = True
        End If
        FreeLibrary hModule
    End If
End Function

Private Sub trackMouseTracking(hWnd As Long)
    Dim lpEventTrack As TRACKMOUSEEVENTTYPE
    With lpEventTrack
        .cbSize = Len(lpEventTrack)
        .dwFlags = &H1 Or &H2
        .dwHoverTime = 100
        .hwndTrack = hWnd
    End With
    If (m_bIsTrackHandler32) Then
        TrackMouseEvent lpEventTrack
    Else
        TrackMouseEvent2 lpEventTrack
    End If
End Sub

Private Function loWORD(DWord As Long) As Integer
    loWORD = (DWord And &H7FFF) Or (((DWord And &H8000&) <> 0) And &H8000)
End Function

Private Function hiWORD(DWord As Long) As Integer
    hiWORD = (DWord And &H7FFF0000) \ &H10000 Or (((DWord And &H80000000) <> 0) And &H8000)
End Function

Private Function pointInRect(ByRef rc As RECT, ByRef pt As POINTAPI) As Boolean
    pointInRect = (pt.x >= rc.Left And pt.x <= rc.Right And pt.y >= rc.Top And pt.y <= rc.Bottom)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' 以下这些重绘无窗口控件的方法还不是很完善。。。
Private Sub enumWindowLessControlsToDraw(ByVal hdc As Long, ByRef rcPaint As RECT)
    Dim ctl As Control
    For Each ctl In m_SSTab.Container
        If (ctl.Container Is m_SSTab) Then
            If TypeName(ctl) = "Label" Then
                Call drawLabel(ctl, hdc, rcPaint)
            ElseIf TypeName(ctl) = "Image" Then
                Call drawImage(ctl, hdc, rcPaint)
            ElseIf TypeName(ctl) = "Shape" Then
                Call drawShape(ctl, hdc, rcPaint)
            ElseIf TypeName(ctl) = "Line" Then
                Call drawLine(ctl, hdc, rcPaint)
            End If
        End If
    Next
End Sub

Private Sub drawLabel(ByRef ctl As Label, ByVal hdc As Long, ByRef rcPaint As RECT)
    If (ctl.Visible = False) Then Exit Sub
    Dim rcControl As RECT
    Dim rcIntersect As RECT
    rcControl.Left = ctl.Left / Screen.TwipsPerPixelX
    rcControl.Top = ctl.Top / Screen.TwipsPerPixelY
    rcControl.Right = rcControl.Left + ctl.Width / Screen.TwipsPerPixelX
    rcControl.Bottom = rcControl.Top + ctl.Height / Screen.TwipsPerPixelY
    
    Call IntersectRect(rcIntersect, rcPaint, rcControl)
    If IsRectEmpty(rcIntersect) Then Exit Sub
    
    Dim hFont As Long, hOldFont As Long
    Dim i As Integer
    Dim nOldBkMode As Long
    Dim lf As LOGFONT
    Dim fnByte() As Byte
    
    fnByte = StrConv(ctl.Font.Name & vbNullString, vbFromUnicode)
    For i = 0 To UBound(fnByte)
        lf.lfFaceName(i + 1) = fnByte(i)
    Next i
    
    lf.lfHeight = ctl.Font.Size + 4
    lf.lfItalic = ctl.Font.Italic
    lf.lfWeight = IIf(ctl.Font.Bold, FW_BOLD, FW_NORMAL)
    lf.lfUnderline = ctl.Font.Underline
    lf.lfStrikeOut = ctl.Font.Strikethrough
    lf.lfCharSet = ctl.Font.Charset
    hFont = CreateFontIndirect(lf)
    
    nOldBkMode = SetBkMode(hdc, IIf(ctl.BackStyle = 1, OPAQUE, TRANSPARENT))
    hOldFont = SelectObject(hdc, hFont)
    Call TextOut(hdc, ctl.Left / Screen.TwipsPerPixelX, ctl.Top / Screen.TwipsPerPixelY, ctl.Caption, Len(ctl.Caption))
    Call SelectObject(hdc, hOldFont)
    Call SetBkMode(hdc, nOldBkMode)
    Call DeleteObject(hFont)
End Sub

Private Sub drawImage(ByRef ctl As Image, ByVal hdc As Long, ByRef rcPaint As RECT)
    If (ctl.Visible = False) Then Exit Sub
    Dim rcControl As RECT
    Dim rcIntersect As RECT
    rcControl.Left = ctl.Left / Screen.TwipsPerPixelX
    rcControl.Top = ctl.Top / Screen.TwipsPerPixelY
    rcControl.Right = rcControl.Left + ctl.Width / Screen.TwipsPerPixelX
    rcControl.Bottom = rcControl.Top + ctl.Height / Screen.TwipsPerPixelY
    
    Call IntersectRect(rcIntersect, rcPaint, rcControl)
    If IsRectEmpty(rcIntersect) Then Exit Sub
    
    If (ctl.Picture Is Nothing) Then Exit Sub
    
    Dim hTmpDC As Long
    Dim hMemDC As Long
    Dim hOldBitmap As Long
    Dim bmp As BITMAP
    hTmpDC = GetDC(0)
    hMemDC = CreateCompatibleDC(hTmpDC)
    hOldBitmap = SelectObject(hMemDC, ctl.Picture.Handle)
    Call GetObject(ctl.Picture.Handle, Len(bmp), bmp)
    If (ctl.Stretch = True) Then
        Call StretchBlt(hdc, rcControl.Left, rcControl.Top, rcControl.Right - rcControl.Left, rcControl.Bottom - rcControl.Top, hMemDC, 0, 0, bmp.bmWidth, bmp.bmHeight, vbSrcCopy)
    Else
        Call BitBlt(hdc, rcControl.Left, rcControl.Top, bmp.bmWidth, bmp.bmHeight, hMemDC, 0, 0, vbSrcCopy)
    End If
    
    Call SelectObject(hMemDC, hOldBitmap)
    Call DeleteDC(hMemDC)
    Call ReleaseDC(0, hTmpDC)
End Sub

Private Sub drawShape(ByRef ctl As Shape, ByVal hdc As Long, ByRef rcPaint As RECT)
    If (ctl.Visible = False) Then Exit Sub
    Dim rcControl As RECT
    Dim rcIntersect As RECT
    rcControl.Left = ctl.Left / Screen.TwipsPerPixelX
    rcControl.Top = ctl.Top / Screen.TwipsPerPixelY
    rcControl.Right = rcControl.Left + ctl.Width / Screen.TwipsPerPixelX
    rcControl.Bottom = rcControl.Top + ctl.Height / Screen.TwipsPerPixelY
    
    Call IntersectRect(rcIntersect, rcPaint, rcControl)
    If IsRectEmpty(rcIntersect) Then Exit Sub
    
End Sub

Private Sub drawLine(ByRef ctl As Line, ByVal hdc As Long, ByRef rcPaint As RECT)
    If (ctl.Visible = False) Then Exit Sub
    Dim rcControl As RECT
    Dim rcIntersect As RECT
    Dim hPen As Long, hOldPen As Long
    Dim dwPenStyle As Long
    Dim pt1 As POINTAPI, pt2 As POINTAPI, ptOld As POINTAPI
    
    pt1.x = ctl.X1 / Screen.TwipsPerPixelX
    pt1.y = ctl.Y1 / Screen.TwipsPerPixelY
    pt2.x = ctl.X2 / Screen.TwipsPerPixelX
    pt2.y = ctl.Y2 / Screen.TwipsPerPixelY
    
    rcControl.Left = IIf(pt1.x < pt2.x, pt1.x, pt2.x)
    rcControl.Top = IIf(pt1.y < pt2.y, pt1.y, pt2.y)
    rcControl.Right = IIf(pt1.x > pt2.x, pt1.x, pt2.x)
    rcControl.Bottom = IIf(pt1.y > pt2.y, pt1.y, pt2.y)
    
    Call IntersectRect(rcIntersect, rcPaint, rcControl)
    If IsRectEmpty(rcIntersect) Then Exit Sub
    
    dwPenStyle = PS_SOLID
    If (ctl.BorderStyle = 0) Then dwPenStyle = PS_NULL
    If (ctl.BorderStyle = 1) Then dwPenStyle = PS_SOLID
    If (ctl.BorderStyle = 2) Then dwPenStyle = PS_DASH
    If (ctl.BorderStyle = 3) Then dwPenStyle = PS_DOT
    If (ctl.BorderStyle = 4) Then dwPenStyle = PS_DASHDOT
    If (ctl.BorderStyle = 5) Then dwPenStyle = PS_DASHDOTDOT
    If (ctl.BorderStyle = 6) Then dwPenStyle = PS_INSIDEFRAME
    hPen = CreatePen(dwPenStyle, ctl.BorderWidth, ctl.BorderColor)
    hOldPen = SelectObject(hdc, hPen)
    Call MoveToEx(hdc, pt1.x, pt1.y, ptOld)
    Call LineTo(hdc, pt2.x, pt2.y)
    Call SelectObject(hdc, hOldPen)
    Call DeleteObject(hPen)
End Sub

Private Function createMemDC(ByVal cx As Long, ByVal cy As Long) As Boolean
    createMemDC = False
    
    ' 创建缓冲HDC
    Call destroyMemDC
    
    Dim hTmpDC As Long
    hTmpDC = GetDC(0)
    
    m_hMemDC = CreateCompatibleDC(hTmpDC)
    If m_hMemDC = 0 Then Exit Function
    
    m_hMemBitmap = CreateCompatibleBitmap(hTmpDC, cx, cy)
    If m_hMemBitmap = 0 Then Exit Function
    
    Call ReleaseDC(0, hTmpDC)
    
    m_nMemDCState = SaveDC(m_hMemDC)
    Call SelectObject(m_hMemDC, m_hMemBitmap)
    
    createMemDC = True
End Function

Private Function destroyMemDC() As Boolean
    destroyMemDC = False
    
    If m_hMemDC <> 0 Then
        Call RestoreDC(m_hMemDC, m_nMemDCState)
        Call DeleteDC(m_hMemDC)
        
        m_hMemDC = 0
        m_nMemDCState = 0
        
        destroyMemDC = True
    End If
    
    If m_hMemBitmap <> 0 Then
        Call DeleteObject(m_hMemBitmap)
        m_hMemBitmap = 0
        
        destroyMemDC = True
    End If
End Function

Private Function onPaint(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef bHandled As Boolean) As Long
    Dim ps As PAINTSTRUCT
    Dim rcClient As RECT
    'Dim rcPaint As RECT
    
    Call GetClientRect(hWnd, rcClient)
    Call BeginPaint(hWnd, ps)
    
    ' 获取脏区域
    'Call GetBoundsRect(ps.hDC, rcPaint, 0)
    ' 设置缓冲DC脏区域边界
    Call SetBoundsRect(m_hMemDC, ps.rcPaint, DCB_SET)
    ' 擦除缓冲DC上的脏区域
    Call FillRect(m_hMemDC, ps.rcPaint, GetStockObject(BLACK_BRUSH))
    
    ' 绘画背景
    Call doPaintBackground(m_hMemDC, rcClient, ps.rcPaint)
    
    ' 画标签
    Call doPaintTabs(m_hMemDC, rcClient, ps.rcPaint)
    
    ' 绘画无窗口的控件
    Call enumWindowLessControlsToDraw(m_hMemDC, ps.rcPaint)
    
    ' 绘画UpDown控件
    If UpDownVisible() = True Then
        Call doPaintUpDown(m_hMemDC, rcClient, ps.rcPaint)
    End If
    
    ' 复制缓冲DC内容到窗口DC
    Call BitBlt(ps.hdc, 0, 0, rcClient.Right - rcClient.Left, rcClient.Bottom - rcClient.Top, m_hMemDC, 0, 0, vbSrcCopy)
    Call EndPaint(hWnd, ps)
    
    bHandled = True
    onPaint = 0
End Function

Private Sub onSize(ByVal nType As Long, ByVal cx As Long, ByVal cy As Long)
    ' Private Const SIZE_MINIMIZED = 1
    If (nType <> 1) Then
        Call createMemDC(cx, cy)
        If (UpDownVisible = True) Then
            Call adjustTab(m_SSTab.Tab)
        Else
            m_TabOffsetX = 0
        End If
        Call UpdateAll
    End If
End Sub

Private Sub onMouseMove(ByVal Button As Long, ByRef Point As POINTAPI)
    Dim nTab As Long, nOldTab As Long
    Dim rcUpButton As RECT, rcDownButton As RECT
    
    nTab = TAB_NULL
    nOldTab = m_HoverTab
    
    Call getTabRect(TAB_UP, rcUpButton)
    Call getTabRect(TAB_DOWN, rcDownButton)
    If (UpDownVisible() = True And pointInRect(rcUpButton, Point)) Then
        nTab = TAB_UP
    ElseIf (UpDownVisible() = True And pointInRect(rcDownButton, Point)) Then
        nTab = TAB_DOWN
    ElseIf Point.y <= m_TabHeight Then
        nTab = Fix((Point.x - m_TabOffsetX) / m_TabWidth)
        If nTab < 0 Or nTab > m_SSTab.Tabs - 1 Then
            nTab = TAB_NULL
        End If
    End If
    
    If nTab <> m_HoverTab Then
        m_HoverTab = nTab
        Call updateTab(nOldTab)
        Call updateTab(m_HoverTab)
    End If
End Sub

Private Sub onMouseDBClick(ByVal Button As Long, ByRef Point As POINTAPI)
    '
End Sub

Private Sub onMouseDown(ByVal Button As Long, ByRef Point As POINTAPI)
    If (Button = 1) Then
        Dim rcUpButton As RECT, rcDownButton As RECT
        
        Call getTabRect(TAB_UP, rcUpButton)
        Call getTabRect(TAB_DOWN, rcDownButton)
        If (UpDownVisible() = True And pointInRect(rcUpButton, Point) = True) Then
            Call onUpButtonClick
        ElseIf (UpDownVisible() = True And pointInRect(rcDownButton, Point) = True) Then
            Call onDownButtonClick
        ElseIf Point.y <= m_TabHeight Then
            Dim nTab As Long
            nTab = Fix((Point.x - m_TabOffsetX) / m_TabWidth)
            If nTab >= 0 And nTab <= m_SSTab.Tabs - 1 Then
                If (m_SSTab.TabEnabled(nTab) = True) Then
                    m_SSTab.Tab = nTab
                    If (UpDownVisible = True) Then
                        Call adjustTab(m_SSTab.Tab)
                    End If
                End If
            End If
        End If
        m_SSTab.SetFocus
    End If
End Sub

Private Sub onMouseUp(ByVal Button As Long, ByRef Point As POINTAPI)
    '
End Sub

Private Sub onMouseHover()
    '
End Sub

Private Sub onMouseLeave()
    If m_HoverTab <> TAB_NULL Then
        Dim nOldTab As Long
        nOldTab = m_HoverTab
        m_HoverTab = TAB_NULL
        Call updateTab(nOldTab)
    End If
End Sub

Private Sub onUpButtonClick()
    Dim i As Integer
    If (m_SSTab.Tab > 0) Then
        For i = m_SSTab.Tab - 1 To 0 Step -1
            If (m_SSTab.TabEnabled(i) = True) Then
                m_SSTab.Tab = i
                Exit Sub
            End If
        Next
    End If
End Sub

Private Sub onDownButtonClick()
    Dim i As Integer
    If (m_SSTab.Tab < m_SSTab.Tabs - 1) Then
        For i = m_SSTab.Tab + 1 To m_SSTab.Tabs - 1
            If (m_SSTab.TabEnabled(i) = True) Then
                m_SSTab.Tab = i
                Exit Sub
            End If
        Next
    End If
End Sub

Private Sub doPaintBackground(ByVal hdc As Long, ByRef rcBackground As RECT, ByRef rcPaint As RECT)
    RaiseEvent DrawBackGround(hdc, rcBackground.Right - rcBackground.Left, rcBackground.Bottom - rcBackground.Top)
End Sub

Private Sub doPaintTabs(ByVal hdc As Long, ByRef rcBackground As RECT, ByRef rcPaint As RECT)
    Dim nTabCount As Integer
    Dim rcIntersect As RECT
    Dim rcTab As RECT
    Dim nState As Long
    Dim i As Integer
    
    nTabCount = m_SSTab.Tabs
    
    For i = 0 To nTabCount - 1
        ' 判断标签是否处于脏区域中，不在脏区域不调用重绘
        Call getTabRect(i, rcTab)
        
        Call IntersectRect(rcIntersect, rcPaint, rcTab)
        If (IsRectEmpty(rcIntersect) = 0) Then
            nState = STATE_NORMAL
            If (i = m_SSTab.Tab) Then nState = nState Or STATE_SELECTED
            If (Me.Focus And i = m_SSTab.Tab) Then nState = nState Or STATE_FOCUSED
            If (i = m_HoverTab) Then nState = nState Or STATE_HOVER
            If (m_SSTab.TabEnabled(i) = False) Then nState = nState Or STATE_DISABLED
            
            ' 设置裁剪区域
            Call SetBoundsRect(hdc, rcIntersect, DCB_SET)
            RaiseEvent DrawTab(i, nState, hdc, rcTab.Left, rcTab.Top, m_TabWidth, m_TabHeight)
        End If
    Next
    
    ' 还原裁剪区域
    Call SetBoundsRect(hdc, rcPaint, DCB_SET)
End Sub

Private Sub doPaintUpDown(ByVal hdc As Long, ByRef rcBackground As RECT, ByRef rcPaint As RECT)
    Dim rcUpDown As RECT
    Dim rcIntersect As RECT
    Dim nState As Long
    
    ' 绘画UpButton
    Call getTabRect(TAB_UP, rcUpDown)
    Call IntersectRect(rcIntersect, rcPaint, rcUpDown)
    If (IsRectEmpty(rcIntersect) = 0) Then
        nState = STATE_NORMAL
        If (m_HoverTab = TAB_UP) Then nState = nState Or STATE_HOVER
        If (m_SSTab.Tab <= 0) Then nState = nState Or STATE_DISABLED
        
        ' 设置裁剪区域
        Call SetBoundsRect(hdc, rcIntersect, DCB_SET)
        RaiseEvent DrawUpDown(True, nState, hdc, rcUpDown.Left, rcUpDown.Top, rcUpDown.Right - rcUpDown.Left, rcUpDown.Bottom - rcUpDown.Top)
    End If
    
    ' 绘画DownButton
    Call getTabRect(TAB_DOWN, rcUpDown)
    Call IntersectRect(rcIntersect, rcPaint, rcUpDown)
    If (IsRectEmpty(rcIntersect) = 0) Then
        nState = STATE_NORMAL
        If (m_HoverTab = TAB_DOWN) Then nState = nState Or STATE_HOVER
        If (m_SSTab.Tab >= m_SSTab.Tabs - 1) Then nState = nState Or STATE_DISABLED
        
        ' 设置裁剪区域
        Call SetBoundsRect(hdc, rcIntersect, DCB_SET)
        RaiseEvent DrawUpDown(False, nState, hdc, rcUpDown.Left, rcUpDown.Top, rcUpDown.Right - rcUpDown.Left, rcUpDown.Bottom - rcUpDown.Top)
    End If
    
    ' 还原裁剪区域
    Call SetBoundsRect(hdc, rcPaint, DCB_SET)
End Sub

Private Sub updateTab(ByVal nTab As Integer)
    Dim rcTab As RECT
    If getTabRect(nTab, rcTab) = True Then
        Call InvalidateRect(m_hSSTabWnd, rcTab, 1)
    End If
End Sub

' 调整活动标签页的位置使其左右都在可视视图内
Private Sub adjustTab(ByVal nTab As Integer)
    Dim rcTab As RECT, rcClient As RECT
    Dim bNeedUpdate As Boolean
    Call GetClientRect(m_hSSTabWnd, rcClient)
    Call getTabRect(nTab, rcTab)
    
    If (rcTab.Left < 0) Then
        m_TabOffsetX = -m_TabWidth * nTab
        bNeedUpdate = True
    End If
    
    If (rcTab.Right > rcClient.Right - rcClient.Left - m_UpDownWidth * 2) Then
        m_TabOffsetX = -(m_TabWidth * (nTab + 1) - (rcClient.Right - rcClient.Left - m_UpDownWidth * 2))
        bNeedUpdate = True
    End If
    
    If (bNeedUpdate = True) Then
        Call UpdateAll
    End If
End Sub

Private Function getTabRect(ByVal nTab As Integer, ByRef rcTab As RECT) As Boolean
    Dim rcClient As RECT
    getTabRect = False
    
    If (nTab = TAB_UP) Then
        Call GetClientRect(m_hSSTabWnd, rcClient)
        rcTab.Left = rcClient.Right - m_UpDownWidth * 2
        rcTab.Top = rcClient.Top + (m_TabHeight - m_UpDownHeight) / 2
        rcTab.Right = rcTab.Left + m_UpDownWidth
        rcTab.Bottom = rcTab.Top + m_UpDownHeight
        
        getTabRect = True
    ElseIf (nTab = TAB_DOWN) Then
        Call GetClientRect(m_hSSTabWnd, rcClient)
        rcTab.Left = rcClient.Right - m_UpDownWidth
        rcTab.Top = rcClient.Top + (m_TabHeight - m_UpDownHeight) / 2
        rcTab.Right = rcTab.Left + m_UpDownWidth
        rcTab.Bottom = rcTab.Top + m_UpDownHeight
        
        getTabRect = True
    ElseIf (nTab >= 0 And nTab <= m_SSTab.Tabs - 1) Then
        rcTab.Left = m_TabOffsetX + m_TabWidth * nTab
        rcTab.Top = 0
        rcTab.Right = rcTab.Left + m_TabWidth
        rcTab.Bottom = rcTab.Top + m_TabHeight
        
        getTabRect = True
    End If
End Function

Public Function Attach(SSTabCtrl As Object) As Boolean
    Attach = False
    If (SSTabCtrl Is Nothing) Then Exit Function
    If Not (m_hSSTabWnd = 0) Then Exit Function
    
    m_HoverTab = TAB_NULL
    m_bIsTracked = False
    
    Set m_SSTab = SSTabCtrl
    m_hSSTabWnd = m_SSTab.hWnd
    '' 确保SSTab只显示一行
    'm_SSTab.TabsPerRow = m_SSTab.Tabs
    m_SSTab.TabOrientation = 0
    
    ' 创建缓冲DC
    Dim rcClient As RECT
    Call GetClientRect(m_hSSTabWnd, rcClient)
    Call createMemDC(rcClient.Right - rcClient.Left, rcClient.Bottom - rcClient.Top)
    
    If (UpDownVisible = True) Then
        Call adjustTab(m_SSTab.Tab)
    End If
    
    ' 子类化
    Call subclass_Start(m_hSSTabWnd)
    Call subclass_AddMsg(m_hSSTabWnd, WM_PAINT, MSG_BEFORE)
    Call subclass_AddMsg(m_hSSTabWnd, WM_SIZE, MSG_AFTER)
    Call subclass_AddMsg(m_hSSTabWnd, WM_SETFOCUS, MSG_AFTER)
    Call subclass_AddMsg(m_hSSTabWnd, WM_KILLFOCUS, MSG_AFTER)
    Call subclass_AddMsg(m_hSSTabWnd, WM_ERASEBKGND, MSG_BEFORE)
    Call subclass_AddMsg(m_hSSTabWnd, WM_NCDESTROY, MSG_BEFORE)
    Call subclass_AddMsg(m_hSSTabWnd, WM_MOUSEMOVE, MSG_BEFORE)
    Call subclass_AddMsg(m_hSSTabWnd, WM_LBUTTONDBLCLK, MSG_BEFORE)
    Call subclass_AddMsg(m_hSSTabWnd, WM_LBUTTONDOWN, MSG_BEFORE)
    Call subclass_AddMsg(m_hSSTabWnd, WM_LBUTTONUP, MSG_BEFORE)
    Call subclass_AddMsg(m_hSSTabWnd, WM_MBUTTONDBLCLK, MSG_BEFORE)
    Call subclass_AddMsg(m_hSSTabWnd, WM_MBUTTONDOWN, MSG_BEFORE)
    Call subclass_AddMsg(m_hSSTabWnd, WM_MBUTTONUP, MSG_BEFORE)
    Call subclass_AddMsg(m_hSSTabWnd, WM_RBUTTONDBLCLK, MSG_BEFORE)
    Call subclass_AddMsg(m_hSSTabWnd, WM_RBUTTONDOWN, MSG_BEFORE)
    Call subclass_AddMsg(m_hSSTabWnd, WM_RBUTTONUP, MSG_BEFORE)
    Call subclass_AddMsg(m_hSSTabWnd, WM_MOUSEHOVER, MSG_AFTER)
    Call subclass_AddMsg(m_hSSTabWnd, WM_MOUSELEAVE, MSG_AFTER)
    
    Call InvalidateRectNull(m_hSSTabWnd, 0, 1)
    Attach = True
End Function

Public Function Detach() As Boolean
    Detach = False
    If (m_SSTab Is Nothing) Then Exit Function
    
    Call subclass_StopAll
    
    Call InvalidateRectNull(m_hSSTabWnd, 0, 1)
    m_hSSTabWnd = 0
    Set m_SSTab = Nothing
    
    Call destroyMemDC
    
    Detach = True
End Function

Public Sub UpdateAll()
    If (m_hSSTabWnd = 0) Then Exit Sub
    Call InvalidateRectNull(m_hSSTabWnd, 0, 1)
End Sub

Public Sub Update(ByVal nleft As Long, ByVal nTop As Long, ByVal nWidth As Long, ByVal nHeight As Long)
    If (m_hSSTabWnd = 0) Then Exit Sub
    
    Dim rcUpdate As RECT
    rcUpdate.Left = nleft
    rcUpdate.Top = nTop
    rcUpdate.Right = nleft + nWidth
    rcUpdate.Bottom = nTop + nHeight
    Call InvalidateRect(m_hSSTabWnd, rcUpdate, 1)
End Sub

Public Function IsAttached() As Boolean
    IsAttached = Not (m_SSTab Is Nothing)
End Function

Public Property Get hWnd() As Long
    hWnd = m_hSSTabWnd
End Property

Public Property Get TabWidth() As Long
    TabWidth = m_TabWidth
End Property

Public Property Let TabWidth(ByVal NewValue As Long)
    If m_TabWidth = NewValue Then Exit Property
    m_TabWidth = NewValue
    Call UpdateAll
End Property

Public Property Get TabHeight() As Long
    TabHeight = m_TabHeight
End Property

Public Property Let TabHeight(ByVal NewValue As Long)
    If m_TabHeight = NewValue Then Exit Property
    m_TabHeight = NewValue
    Call UpdateAll
End Property

Public Property Get UpDownVisible() As Boolean
    UpDownVisible = False
    If (m_SSTab Is Nothing) Then Exit Property
    Dim rcClient As RECT
    Call GetClientRect(m_hSSTabWnd, rcClient)
    UpDownVisible = ((m_SSTab.Tabs * m_TabWidth) > rcClient.Right - rcClient.Left - m_UpDownWidth * 2)
End Property

Public Property Get UpDownWidth() As Long
    UpDownWidth = m_UpDownWidth
End Property

Public Property Let UpDownWidth(ByVal NewValue As Long)
    If m_UpDownWidth = NewValue Then Exit Property
    m_UpDownWidth = NewValue
    Call UpdateAll
End Property

Public Property Get UpDownHeight() As Long
    UpDownHeight = m_UpDownHeight
End Property

Public Property Let UpDownHeight(ByVal NewValue As Long)
    If m_UpDownHeight = NewValue Then Exit Property
    m_UpDownHeight = NewValue
    Call UpdateAll
End Property

Public Property Get Focus() As Boolean
    Focus = False
    If m_hSSTabWnd = 0 Then Exit Property
    
    Focus = (GetFocus = m_hSSTabWnd)
End Property

Private Sub m_SSTab_Click(PreviousTab As Integer)
    Call adjustTab(m_SSTab.Tab)
End Sub
