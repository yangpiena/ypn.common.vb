Attribute VB_Name = "ModBarEAN8And13"
'---------------------------------------------------------------------------------------
' Module    : ModBarEAN8And13
' Author    : YPN
' Date      : 2017-07-18 16:08
' Purpose   : EAN是欧洲的标准，有EAN8、EAN13、EAN128标准。
'             EAN13共13个数字。前3位为国家码；接着的4位是厂商码；再5位是产品码，最后1位是校验码。
'             EAN8是EAN13的精简版。前2位为国家码；5位产品码；1位校验码；
'---------------------------------------------------------------------------------------

Option Explicit
Private Const M_CHKCHAR = 43

Private m_LeftHand_Odd()  As Variant
Private m_LeftHand_Even() As Variant
Private m_Right_Hand()    As Variant
Private m_Parity()        As Variant
Private m_BarH            As Long
Private m_BarText         As String
Private m_Obj             As Object
Private m_Pos             As Long
Private m_Top             As Long
Private m_HasCaption      As Boolean
Private m_Start           As Integer
Private m_PosCtr          As Integer
Private m_Total           As Long
Private m_ChkSum          As Long


Public Function MBarEAN13(i_BarText As String, i_BarHeight As Integer, Optional ByVal i_HasCaption As Boolean = False) As StdPicture
    
    Set m_Obj = FrmPublic.Picture1
    Call init_Table
    m_BarText = i_BarText
    m_HasCaption = i_HasCaption
    m_Obj.Picture = Nothing
    
    If Not checkCode13 Then Exit Function
    
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
    m_Obj.Width = (((Len(m_BarText)) * 8)) * 20
    
    Call paint_Bar13(m_BarText)
    
    Set MBarEAN13 = FrmPublic.Picture1.Image
    
End Function

Public Function MBarEAN8(i_BarText As String, i_BarHeight As Integer, Optional ByVal i_HasCaption As Boolean = False) As StdPicture
    
    Set m_Obj = FrmPublic.Picture1
    init_Table
    m_BarText = i_BarText
    m_HasCaption = i_HasCaption
    m_Obj.Picture = Nothing
    
    If Not checkCode8 Then Exit Function
    
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
    m_Obj.Width = (((Len(m_BarText)) * 8) + 20) * 20 'Screen.TwipsPerPixelX
    
    Call paint_Bar8(m_BarText)
    
    Set MBarEAN8 = FrmPublic.Picture1.Image
    
End Function

Private Function checkCode13() As Boolean
    
    Dim ii As Integer
    
    If Len(m_BarText) <> 12 Then
        Err.Raise vbObjectError + 513, "EAN-13", _
        "Should be 12 Digit Numbers"
        GoTo Err_Found
    End If
    For ii = 1 To Len(m_BarText)
        If InStr("0123456789", Mid(m_BarText, ii, 1)) = 0 Then
            Err.Raise vbObjectError + 513, "EAN-13", _
            "An Invalid Character Found in Bar Text"
            GoTo Err_Found
        End If
    Next
    checkCode13 = True
    Exit Function
    
Err_Found:
    checkCode13 = False
End Function

Private Function checkCode8() As Boolean
    
    Dim ii As Integer
    
    If Len(m_BarText) <> 7 Then
        Err.Raise vbObjectError + 513, "EAN-8", _
        "Should be 7 Digit Numbers"
        GoTo Err_Found
    End If
    For ii = 1 To Len(m_BarText)
        If InStr("0123456789", Mid(m_BarText, ii, 1)) = 0 Then
            Err.Raise vbObjectError + 513, "EAN-8", _
            "An Invalid Character Found in Bar Text"
            GoTo Err_Found
        End If
    Next
    checkCode8 = True
    Exit Function
    
Err_Found:
    checkCode8 = False
End Function

Private Sub paint_Bar13(ByVal i_str As String)
    
    Dim ii As Long, jj As Integer, ctr As Integer, xEven As Boolean, xParity As String
    
    m_Total = 0
    m_Pos = 5
    
    If m_HasCaption Then
        m_Obj.CurrentX = m_Pos
        m_Obj.CurrentY = 5 + m_BarH
        
        m_Obj.Print Mid(i_str, 1, 1)
    End If
    Call draw_Bar("101", True)
    
    m_Obj.CurrentY = 15 + m_BarH
    xParity = m_Parity(CInt(Mid(i_str, 1, 1)))
    
    For ii = 1 To Len(i_str)
        If ((Len(i_str) + 1) - ii) Mod 2 = 0 Then
            m_Total = m_Total + (CInt(Mid(i_str, ii, 1)))
        Else
            m_Total = m_Total + CInt(Mid(i_str, ii, 1) * 3)
        End If
        If ii = 8 Then
            Call draw_Bar("01010", True)
        End If
        jj = CInt(Mid(i_str, ii, 1))
        If ii > 1 And ii < 8 Then
            Call draw_Bar(CStr(IIf(Mid(xParity, ii - 1, 1) = "E", m_LeftHand_Even(jj), m_LeftHand_Odd(jj))), False)
        ElseIf ii > 1 And ii >= 8 Then
            Call draw_Bar(CStr(m_Right_Hand(jj)), False)
        End If
    Next
    m_ChkSum = 0
    jj = m_Total Mod 10
    If jj <> 0 Then
        m_ChkSum = 10 - jj
    End If
    Call draw_Bar(CStr(m_Right_Hand(m_ChkSum)), False)
    Call draw_Bar("101", True)
    
    If m_HasCaption Then
        m_Obj.CurrentX = 23
        m_Obj.CurrentY = 10 + m_BarH
        m_Obj.Print Mid(i_str, 2, 6)
        
        m_Obj.CurrentX = 68
        m_Obj.CurrentY = 10 + m_BarH
        m_Obj.Print Mid(i_str, 8, 6) & m_ChkSum
    End If
    
End Sub

Private Sub paint_Bar8(ByVal i_str As String)
    
    Dim ii As Long, jj As Integer, ctr As Integer, xEven As Boolean, xParity As String
    
    m_Total = 0
    m_Pos = 5
    
    Call draw_Bar("101", True)
    
    m_Obj.CurrentX = m_Pos
    m_Obj.CurrentY = 15 + m_BarH
    xParity = m_Parity(7) 'CInt(Mid(i_str, 1, 1)))
    
    For ii = 1 To Len(i_str)
        If ((Len(i_str) + 1) - ii) Mod 2 = 0 Then
            m_Total = m_Total + (CInt(Mid(i_str, ii, 1)))
        Else
            m_Total = m_Total + CInt(Mid(i_str, ii, 1) * 3)
        End If
        If ii = 5 Then
            Call draw_Bar("01010", True)
        End If
        jj = CInt(Mid(i_str, ii, 1))
        If ii < 5 Then
            Call draw_Bar(CStr(m_LeftHand_Odd(jj)), False)
        ElseIf ii >= 5 Then
            Call draw_Bar(CStr(m_Right_Hand(jj)), False)
        End If
    Next
    m_ChkSum = 0
    jj = m_Total Mod 10
    If jj <> 0 Then
        m_ChkSum = 10 - jj
    End If
    Call draw_Bar(CStr(m_Right_Hand(m_ChkSum)), False)
    Call draw_Bar("101", True)
    
    If m_HasCaption Then
        m_Obj.CurrentX = 23
        m_Obj.CurrentY = 10 + m_BarH
        m_Obj.Print Mid(i_str, 1, 4)
        
        m_Obj.CurrentX = 53
        m_Obj.CurrentY = 10 + m_BarH
        m_Obj.Print Mid(i_str, 5, 4) & m_ChkSum
    End If
    
End Sub

Private Sub draw_Bar(i_encoding As String, i_guard As Boolean)
    
    Dim ii As Integer
    
    For ii = 1 To Len(i_encoding)
        m_Pos = m_Pos + 1
        m_Obj.Line (m_Pos + 10, m_Top)-(m_Pos + 10, m_Top + m_BarH + IIf(i_guard, 5, 0)), IIf(Mid(i_encoding, ii, 1), vbBlack, vbWhite)
    Next
    
End Sub

Private Sub init_Table()
    
    m_LeftHand_Odd = Array("0001101", "0011001", "0010011", "0111101", "0100011", "0110001", "0101111", "0111011", "0110111", "0001011")
    m_LeftHand_Even = Array("0100111", "0110011", "0011011", "0100001", "0011101", "0111001", "0000101", "0010001", "0001001", "0010111")
    m_Right_Hand = Array("1110010", "1100110", "1101100", "1000010", "1011100", "1001110", "1010000", "1000100", "1001000", "1110100")
    m_Parity = Array("OOOOOO", "OOEOEE", "OOEEOE", "OOEEEO", "OEOOEE", "OEEOOE", "OEEEOO", "OEOEOE", "OEOEEO", "OEEOEO")
    
End Sub


