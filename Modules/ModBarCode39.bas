Attribute VB_Name = "ModBarCode39"
'---------------------------------------------------------------------------------------
' Module    : ModBarCode39
' Author    : YPN
' Date      : 2017-07-18 16:01
' Purpose   : Code39是Intermec公司于1971年发明的条码码制，是目前应用最广泛的编码之一。可表示包含数字、英文字母的44个字符。
'             Code93和Code39编码相同，唯一的区别是Code93占用的尺寸更小。
'             Code128可表示ASCII码表的0-127，故命名为Code 128.
'---------------------------------------------------------------------------------------

Option Explicit
Private Const M_CHKCHAR = 43

Private m_Code_A       As String
Private m_Code_B()     As Variant
Private m_BarH         As Long
Private m_BarText      As String
Private m_Obj          As Object
Private m_HasCaption   As Boolean
Private m_Pos          As Long
Private m_Top          As Long
Private m_Start        As Integer
Private m_PosCtr       As Integer
Private m_Total        As Long
Private m_ChkSum       As Long
Private m_WithCheckSum As Boolean


Public Function MBarCode39(i_BarText As String, i_BarHeight As Integer, Optional i_WithCheckSum As Boolean = False, Optional ByVal i_HasCaption As Boolean = False) As StdPicture
    
    Set m_Obj = FrmPublic.Picture1
    m_WithCheckSum = i_WithCheckSum
    init_Table
    m_BarText = i_BarText
    m_HasCaption = i_HasCaption
    m_Obj.Picture = Nothing
    
    If Not checkCode Then Exit Function
    
    m_BarH = i_BarHeight * 10
    m_Top = 10
    
    m_Obj.BackColor = vbWhite
    m_Obj.AutoRedraw = True
    m_Obj.ScaleMode = 3
    
    If i_HasCaption Then
        m_Obj.Height = (m_Obj.TextHeight(m_BarText) + m_BarH + 25) * Screen.TwipsPerPixelY
    Else
        m_Obj.Height = (m_BarH + 20) * Screen.TwipsPerPixelY
    End If
    'm_Obj.Height = (m_Obj.TextHeight(m_BarText) + m_BarH + 25) * Screen.TwipsPerPixelY
    m_Obj.Width = (((Len(m_BarText) + IIf(m_WithCheckSum, 3, 2)) * 12 + 20)) * 16
    
    Call paint_Bar(m_BarText)
    
    Set MBarCode39 = FrmPublic.Picture1.Image
    
End Function

Private Function checkCode() As Boolean

    Dim ii As Integer
    
    m_BarText = UCase(Replace(m_BarText, "*", ""))
    For ii = 1 To Len(m_BarText)
        If InStr(m_Code_A, Mid(m_BarText, ii, 1)) = 0 Then
            GoTo Err_Found
        End If
    Next
    checkCode = True
    Exit Function
    
Err_Found:
    Err.Raise vbObjectError + 513, "Bar 39", _
    "An Invalid Character Found in Bar Text"
    checkCode = False
    
End Function

Private Sub paint_Bar(i_str As String)

    Dim ii As Long, jj As Integer, ctr As Integer
    
    m_Total = 0
    m_Pos = 1
    m_PosCtr = 0
    
    Call draw_Bar(CStr(m_Code_B(M_CHKCHAR)))
    
    For ii = 1 To Len(i_str)
        m_PosCtr = InStr(m_Code_A, Mid(i_str, ii, 1)) - 1
        m_Total = m_Total + m_PosCtr
        Call draw_Bar(CStr(m_Code_B(m_PosCtr)))
    Next
    m_ChkSum = m_Total Mod 43
    
    If m_WithCheckSum Then Call draw_Bar(CStr(m_Code_B(m_ChkSum)))
    
    Call draw_Bar(CStr(m_Code_B(M_CHKCHAR)))
    
    If m_HasCaption Then
        m_Obj.CurrentX = ((m_Pos + 20) / 2) - m_Obj.TextWidth(i_str) / 2 '水平坐标
        m_Obj.CurrentY = 15 + m_BarH    ' 垂直坐标
        m_Obj.Print i_str   ' 大印信息
    End If
    
End Sub

Private Sub draw_Bar(i_encoding As String)

    Dim ii As Integer
    
    For ii = 1 To Len(i_encoding)
        m_Pos = m_Pos + 1
        m_Obj.Line (m_Pos + 10, m_Top)-(m_Pos + 10, m_Top + m_BarH), IIf(Mid(i_encoding, ii, 1), vbBlack, vbWhite)
    Next
    
    m_Pos = m_Pos + 1
    m_Obj.Line (m_Pos + 10, m_Top)-(m_Pos + 10, m_Top + m_BarH), vbWhite
    
End Sub

Private Sub init_Table()

    m_Code_A = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%*"
    m_Code_B = Array( _
    "101001101101", "110100101011", "101100101011", "110110010101", "101001101011", "110100110101", _
    "101100110101", "101001011011", "110100101101", "101100101101", "110101001011", "101101001011", _
    "110110100101", "101011001011", "110101100101", "101101100101", "101010011011", "110101001101", _
    "101101001101", "101011001101", "110101010011", "101101010011", "110110101001", "101011010011", _
    "110101101001", "101101101001", "101010110011", "110101011001", "101101011001", "101011011001", _
    "110010101011", "100110101011", "110011010101", "100101101011", "110010110101", "100110110101", _
    "100101011011", "110010101101", "100110101101", "100100100101", "100100101001", "100101001001", _
    "101001001001", "100101101101" _
    )
    
End Sub
