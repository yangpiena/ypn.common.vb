Attribute VB_Name = "ModQRCode"
'---------------------------------------------------------------------------------------
' Module    : ModQRCode
' Author    : YPN
' Date      : 2018-04-08 22:17
' Purpose   : 二维码
'---------------------------------------------------------------------------------------

Option Explicit
Private Declare Function StretchDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, ByRef lpBits As Any, ByRef lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Guid, ByVal fPictureOwnsHandle As Long, ipic As IUnknown) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpWideCharStr As Any, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByRef lpDefaultChar As Any, ByRef lpUsedDefaultChar As Any) As Long

Private Const m_UTF8 As Long = 65001

Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


'---------------------------------------------------------------------------------------
' Procedure : MQRCode
' Author    : YPN
' Date      : 2017-09-25 15:41
' Purpose   : 生成QR Code码制的二维码
' Param     : i_QRText    二维码内容
'             i_Version  （可选参数）生成版本，支持40种，从1到40，默认自动，即0
'             i_ECLevel  （可选参数）容错等级，支持4种：L-7%、M-15%、Q-25%、H-30%，默认M（传入首字母L、M、Q、H即可）
'             i_MaskType （可选参数）模糊类型，支持8种，从0到7，默认自动，即-1
'             i_Encoding （可选参数）字符编码，支持2种：UTF-8 和 ANSI，默认UTF-8
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MQRCode(ByVal i_QRText As String, Optional ByVal i_Version As Long = 0, Optional ByVal i_ECLevel As String = "M", Optional ByVal i_MaskType As Long = -1, Optional ByVal i_encoding As String = "UTF-8") As StdPicture
    
    Dim v_QRC     As New ClsQRCode
    Dim v_Input() As Byte
    Dim v_QRText  As String
    Dim v_ECLevel As Long
        
    '确定容错能力等级
    Select Case UCase(i_encoding)
    Case "L"
        v_ECLevel = 1
        
    Case "M"
        v_ECLevel = 2
        
    Case "Q"
        v_ECLevel = 3
        
    Case "H"
        v_ECLevel = 4
    
    Case Else
        v_ECLevel = 2
    End Select
    
    '确定字符编码
    Select Case UCase(i_encoding)
    Case "UTF-8"
        v_QRText = i_QRText
        j = Len(v_QRText)
        i = j * 3 + 64
        ReDim v_Input(i)
        j = WideCharToMultiByte(m_UTF8, 0, ByVal StrPtr(v_QRText), j, v_Input(0), i, ByVal 0, ByVal 0)
        
    Case Else
        v_QRText = StrConv(i_QRText, vbFromUnicode)
        v_Input = v_QRText
        j = LenB(v_QRText)
    End Select
    
    '生成二维码
    Set MQRCode = v_QRC.Encode(v_Input, j, i_Version, v_ECLevel, i_MaskType)
    
End Function

'color depth: 8 bits, 0=white, all other=black
'width must be a multiple of 4
Public Function MByteArrayToPicture(ByVal lp As Long, ByVal nWidth As Long, ByVal nHeight As Long, Optional ByVal nLeftPadding As Long, Optional ByVal nTopPadding As Long, Optional ByVal nRightPadding As Long, Optional ByVal nBottomPadding As Long) As StdPicture
    
    Dim tBMI As BITMAPINFO
    Dim h As Long, hdc As Long, hBmp As Long
    Dim hbr As Long
    Dim r As RECT
    
    '///
    With tBMI.bmiHeader
        .biSize = 40&
        .biWidth = nWidth
        .biHeight = -nHeight
        .biPlanes = 1
        .biBitCount = 8
        .biSizeImage = nWidth * nHeight
        .biClrUsed = 256
    End With
    tBMI.bmiColors(0) = &HFFFFFF
    tBMI.bmiColors(2) = &H808080 'debug only
    '///
    h = GetDC(0)
    hdc = CreateCompatibleDC(h)
    r.Right = nWidth + nLeftPadding + nRightPadding
    r.Bottom = nHeight + nTopPadding + nBottomPadding
    hBmp = CreateCompatibleBitmap(h, r.Right, r.Bottom)
    hBmp = SelectObject(hdc, hBmp)
    '///
    hbr = CreateSolidBrush(vbWhite)
    FillRect hdc, r, hbr
    DeleteObject hbr
    StretchDIBits hdc, nLeftPadding, nTopPadding, nWidth, nHeight, 0, 0, nWidth, nHeight, ByVal lp, tBMI, 0, vbSrcCopy
    '///
    hBmp = SelectObject(hdc, hBmp)
    DeleteDC hdc
    ReleaseDC 0, h
    '///
    Set MByteArrayToPicture = bitmapToPicture(hBmp, 1)
    
End Function

Private Function bitmapToPicture(ByVal hBmp As Long, ByVal fPictureOwnsHandle As Long) As StdPicture
    
    If (hBmp = 0) Then Exit Function
    
    Dim oNewPic As IUnknown, tPicConv As PictDesc, IGuid As Guid
    
    ' Fill PictDesc structure with necessary parts:
    With tPicConv
        .cbSizeofStruct = Len(tPicConv)
        .picType = vbPicTypeBitmap
        .hImage = hBmp
    End With
    
    ' Fill in IUnknown Interface ID
    With IGuid
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    
    ' Create a picture object:
    OleCreatePictureIndirect tPicConv, IGuid, fPictureOwnsHandle, oNewPic
    
    ' Return it:
    Set bitmapToPicture = oNewPic
    
End Function

