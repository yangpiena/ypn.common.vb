Attribute VB_Name = "ModBarCode128"
'---------------------------------------------------------------------------------------
' Module    : ModBarCode128
' Author    : YPN
' Date      : 2017-07-18 15:54
' Purpose   : Code39是Intermec公司于1971年发明的条码码制，是目前应用最广泛的编码之一。可表示包含数字、英文字母的44个字符。
'             Code93和Code39编码相同，唯一的区别是Code93占用的尺寸更小。
'             Code128可表示ASCII码表的0-127，故命名为Code 128.
'---------------------------------------------------------------------------------------

Option Explicit
Private Const M_CODEC = 99
Private Const M_CODEB = 100
Private Const M_CODEA = 101
Private Const M_FNC1 = 102
Private Const M_STARTA = 103
Private Const M_STARTB = 104
Private Const M_STARTC = 105

Private m_Code_A     As String
Private m_Code_B()   As Variant
Private m_BarH       As Long
Private m_BarText    As String
Private m_Obj        As Object
Private m_HasCaption As Boolean
Private m_Pos        As Long
Private m_Top        As Long
Private m_Cnt        As Integer
Private m_Start      As Integer
Private m_PosCtr     As Integer
Private m_Total      As Long
Private m_ChkSum     As Long


Public Function MBarCode128(i_BarText As String, i_BarHeight As Integer, Optional ByVal i_HasCaption As Boolean = False) As StdPicture
    
    On Error GoTo MBarCode128_Error
    
    Set m_Obj = FrmPublic.Picture1
    
    Call init_Table
    
    m_Top = 10
    m_BarH = i_BarHeight * 10
    m_BarText = Replace(Trim(i_BarText), Chr(13) + Chr(10), "")  '去空格、回车符CHAR(13)、换行符CHAR(10)
    m_HasCaption = i_HasCaption
    m_Obj.Picture = Nothing
    m_Obj.BackColor = vbWhite
    m_Obj.AutoRedraw = True
    m_Obj.ScaleMode = 3
    
    If i_HasCaption Then
        m_Obj.Height = (m_Obj.TextHeight(m_BarText) + m_BarH + 25) * Screen.TwipsPerPixelY
    Else
        m_Obj.Height = (m_BarH + 20) * Screen.TwipsPerPixelY
    End If
    'm_Obj.Height = (m_Obj.TextHeight(m_BarText) + m_BarH + 25) * Screen.TwipsPerPixelY
    m_Obj.Width = ((test_String(m_BarText) + 3) * 11 + 25) * Screen.TwipsPerPixelX
    
    Call paint_Code(m_BarText)
    
    Set MBarCode128 = FrmPublic.Picture1.Image
    
    On Error GoTo 0
    Exit Function
    
MBarCode128_Error:
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MBarCode128 of Module ModBarCode128"
    
End Function

Private Function test_String(i_str As String)
    
    Dim ii As Long, jj As Integer, ctr As Integer
    
    ctr = 0
    jj = 0
    
    For ii = 1 To Len(i_str)
        If InStr("0123456789", Mid(i_str, ii, 1)) > 0 Then
            ctr = ctr + 1
        Else
            jj = jj + IIf(ctr = 0, 1, ctr)
            ctr = 0
        End If
    Next
    If (ctr >= 4 And ii >= Len(i_str)) Then
        If jj <> 0 Then jj = jj + 1
        If ctr Mod 2 <> 0 Then
            ctr = ctr - 1
            jj = jj + 2
            
        End If
        jj = jj + (ctr / 2)
    End If
    
    test_String = jj
    
End Function

Private Sub paint_Code(i_str As String)
    
    Dim ii As Long, jj As Integer, ctr As Integer
    
    m_Total = 0
    m_Pos = 1
    m_Start = 0
    ctr = 0
    m_PosCtr = 0
    m_Cnt = 0
    
    For ii = 1 To Len(i_str)
        If InStr("0123456789", Mid(i_str, ii, 1)) > 0 Then
            ctr = ctr + 1
        Else
            For jj = ii - ctr To ii
                Call printB(Mid(i_str, jj, 1))
                m_Cnt = m_Cnt + 1
            Next
            ctr = 0
        End If
    Next
    If (ctr >= 4 And ii >= Len(i_str)) Then
        If ctr Mod 2 <> 0 Then
            m_Cnt = m_Cnt + 1
            Call printB(Mid(i_str, ii - ctr, 1))
            ctr = ctr - 1
        End If
        Call printC(Mid(i_str, ii - ctr, ctr))
    End If
    m_ChkSum = m_Total Mod 103
    Call draw_Bar(CStr(m_Code_B(m_ChkSum)))
    Call draw_Bar("1100011101011")
    
    If m_HasCaption Then
        m_Obj.CurrentX = ((m_Pos + 20) / 2) - m_Obj.TextWidth(i_str) / 2   ' 水平坐标
        m_Obj.CurrentY = 15 + m_BarH    ' 垂直坐标
        m_Obj.Print i_str   '　打印信息
    End If
    'Picture = Me.Image
    
End Sub

Private Sub printB(i_str As String)
    
    m_PosCtr = m_PosCtr + 1
    m_Total = m_Total + ((InStr(m_Code_A, i_str) - 1) * m_PosCtr)
    
    If m_Start <> M_STARTB Then
        If m_Start = 0 Then
            m_Total = m_Total + M_STARTB
            m_Start = M_STARTB
            Call draw_Bar(CStr(m_Code_B(M_STARTB)))
        Else
            m_Start = M_CODEB
            Call draw_Bar(CStr(m_Code_B(M_CODEB)))
            m_PosCtr = m_PosCtr + 1
            m_Total = m_Total + (M_CODEB * m_PosCtr)
        End If
    End If
    
    Call draw_Bar(CStr(m_Code_B(InStr(m_Code_A, i_str) - 1)))
    
End Sub

Private Sub printC(i_str As String)
    
    Dim jj As Integer
    
    If m_Start <> M_STARTC Then
        If m_Start = 0 Then
            m_Total = m_Total + M_STARTC
            m_Start = M_STARTC
            draw_Bar CStr(m_Code_B(M_STARTC))
        Else
            m_Start = M_CODEC
            draw_Bar CStr(m_Code_B(M_CODEC))
            m_PosCtr = m_PosCtr + 1
            m_Total = m_Total + (M_CODEC * m_PosCtr)
        End If
    End If
    
    Call setC(i_str)
    
    For jj = 1 To Len(i_str) Step 2
        m_PosCtr = m_PosCtr + 1
        m_Total = m_Total + CInt(Mid(i_str, jj, 2)) * m_PosCtr
    Next
    
End Sub

Private Sub setC(i_str As String)
    
    Dim ii As Integer
    
    For ii = 1 To Len(i_str) Step 2
        draw_Bar CStr(m_Code_B(CInt(Mid(i_str, ii, 2))))
        m_Cnt = m_Cnt + 1
    Next
    
End Sub

Private Sub draw_Bar(i_encoding As String)
    
    Dim ii As Integer
    
    For ii = 1 To Len(i_encoding)
        m_Pos = m_Pos + 1
        m_Obj.Line (m_Pos + 10, m_Top)-(m_Pos + 10, m_Top + m_BarH), IIf(Mid(i_encoding, ii, 1), vbBlack, vbWhite)
    Next
    
    ii = 0
    
End Sub

Private Sub init_Table()
    
    m_Code_A = " !""#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~"
    m_Code_B = Array( _
    "11011001100", "11001101100", "11001100110", "10010011000", "10010001100", "10001001100", _
    "10011001000", "10011000100", "10001100100", "11001001000", "11001000100", "11000100100", _
    "10110011100", "10011011100", "10011001110", "10111001100", "10011101100", "10011100110", _
    "11001110010", "11001011100", "11001001110", "11011100100", "11001110100", "11101101110", _
    "11101001100", "11100101100", "11100100110", "11101100100", "11100110100", "11100110010", _
    "11011011000", "11011000110", "11000110110", "10100011000", "10001011000", "10001000110", _
    "10110001000", "10001101000", "10001100010", "11010001000", "11000101000", "11000100010", _
    "10110111000", "10110001110", "10001101110", "10111011000", "10111000110", "10001110110", _
    "11101110110", "11010001110", "11000101110", "11011101000", "11011100010", "11011101110", _
    "11101011000", "11101000110", "11100010110", "11101101000", "11101100010", "11100011010", _
    "11101111010", "11001000010", "11110001010", "10100110000", "10100001100", "10010110000", _
    "10010000110", "10000101100", "10000100110", "10110010000", "10110000100", "10011010000", _
    "10011000010", "10000110100", "10000110010", "11000010010", "11001010000", "11110111010", _
    "11000010100", "10001111010", "10100111100", "10010111100", "10010011110", "10111100100", _
    "10011110100", "10011110010", "11110100100", "11110010100", "11110010010", "11011011110", _
    "11011110110", "11110110110", "10101111000", "10100011110", "10001011110", "10111101000", _
    "10111100010", "11110101000", "11110100010", "10111011110", "10111101110", "11101011110", _
    "11110101110", "11010000100", "11010010000", "11010011100" _
    )
End Sub
