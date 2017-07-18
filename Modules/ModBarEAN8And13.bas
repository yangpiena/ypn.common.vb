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

Private Const ChkChar = 43

Private LeftHand_Odd()  As Variant
Private LeftHand_Even() As Variant
Private Right_Hand()    As Variant
Private Parity()        As Variant

Private BarH     As Long
Private zBarText As String
Private xObj     As Object

Private xPos   As Long, xtop      As Long, zHasCaption As Boolean
Private xStart As Integer, posCtr As Integer, xTotal   As Long, chkSum As Long



Public Sub MBarEAN13(zObj As Object, zBarH As Integer, BarText As String, Optional ByVal HasCaption As Boolean = False)

    Set xObj = zObj
    init_Table
    zBarText = BarText
    zHasCaption = HasCaption
    xObj.Picture = Nothing
    
    If Not CheckCode13 Then Exit Sub
    
    BarH = zBarH * 10
    xtop = 10
    
    xObj.BackColor = vbWhite
    xObj.AutoRedraw = True
    xObj.ScaleMode = 3
    If HasCaption Then
        xObj.Height = (xObj.TextHeight(zBarText) + BarH + 25) * Screen.TwipsPerPixelY
    Else
        xObj.Height = (BarH + 20) * Screen.TwipsPerPixelY
    End If
    xObj.Width = (((Len(zBarText)) * 8)) * 20
    
    paint_Bar13 zBarText
    zObj.Picture = zObj.Image
    
End Sub

Public Sub MBarEAN8(zObj As Object, zBarH As Integer, BarText As String, Optional ByVal HasCaption As Boolean = False)

    Set xObj = zObj
    init_Table
    zBarText = BarText
    zHasCaption = HasCaption
    xObj.Picture = Nothing
    
    If Not checkCode8 Then Exit Sub
    
    BarH = zBarH * 10
    xtop = 10
    
    xObj.BackColor = vbWhite
    xObj.AutoRedraw = True
    xObj.ScaleMode = 3
    
    If HasCaption Then
        xObj.Height = (xObj.TextHeight(zBarText) + BarH + 25) * Screen.TwipsPerPixelY
    Else
        xObj.Height = (BarH + 20) * Screen.TwipsPerPixelY
    End If
    'xObj.Height = (xObj.TextHeight(zBarText) + BarH + 25) * Screen.TwipsPerPixelY
    xObj.Width = (((Len(zBarText)) * 8) + 20) * 20 'Screen.TwipsPerPixelX
    
    paint_Bar8 zBarText
    zObj.Picture = zObj.Image
    
End Sub

Private Function CheckCode13() As Boolean

    Dim ii As Integer
    
    If Len(zBarText) <> 12 Then
        Err.Raise vbObjectError + 513, "EAN-13", _
        "Should be 12 Digit Numbers"
        GoTo Err_Found
    End If
    For ii = 1 To Len(zBarText)
        If InStr("0123456789", Mid(zBarText, ii, 1)) = 0 Then
            Err.Raise vbObjectError + 513, "EAN-13", _
            "An Invalid Character Found in Bar Text"
            GoTo Err_Found
        End If
    Next
    CheckCode13 = True
    Exit Function
    
Err_Found:
    CheckCode13 = False
End Function

Private Function checkCode8() As Boolean

    Dim ii As Integer
    
    If Len(zBarText) <> 7 Then
        Err.Raise vbObjectError + 513, "EAN-8", _
        "Should be 7 Digit Numbers"
        GoTo Err_Found
    End If
    For ii = 1 To Len(zBarText)
        If InStr("0123456789", Mid(zBarText, ii, 1)) = 0 Then
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

Private Sub paint_Bar13(ByVal xstr As String)

    Dim ii As Long, jj As Integer, ctr As Integer, xEven As Boolean, xParity As String
    
    xTotal = 0
    xPos = 5
    
    If zHasCaption Then
        xObj.CurrentX = xPos
        xObj.CurrentY = 5 + BarH
        
        xObj.Print Mid(xstr, 1, 1)
    End If
    draw_Bar "101", True
    
    xObj.CurrentY = 15 + BarH
    xParity = Parity(CInt(Mid(xstr, 1, 1)))
    
    
    For ii = 1 To Len(xstr)
        If ((Len(xstr) + 1) - ii) Mod 2 = 0 Then
            xTotal = xTotal + (CInt(Mid(xstr, ii, 1)))
        Else
            xTotal = xTotal + CInt(Mid(xstr, ii, 1) * 3)
        End If
        If ii = 8 Then
            draw_Bar "01010", True
        End If
        jj = CInt(Mid(xstr, ii, 1))
        If ii > 1 And ii < 8 Then
            draw_Bar CStr(IIf(Mid(xParity, ii - 1, 1) = "E", LeftHand_Even(jj), LeftHand_Odd(jj))), False
        ElseIf ii > 1 And ii >= 8 Then
            draw_Bar CStr(Right_Hand(jj)), False
        End If
    Next
    chkSum = 0
    jj = xTotal Mod 10
    If jj <> 0 Then
        chkSum = 10 - jj
    End If
    draw_Bar CStr(Right_Hand(chkSum)), False
    draw_Bar "101", True
    
    If zHasCaption Then
        xObj.CurrentX = 23
        xObj.CurrentY = 10 + BarH
        xObj.Print Mid(xstr, 2, 6)
        
        xObj.CurrentX = 68
        xObj.CurrentY = 10 + BarH
        xObj.Print Mid(xstr, 8, 6) & chkSum
    End If
    
End Sub

Private Sub paint_Bar8(ByVal xstr As String)

    Dim ii As Long, jj As Integer, ctr As Integer, xEven As Boolean, xParity As String
    
    xTotal = 0
    xPos = 5
    
    
    draw_Bar "101", True
    
    xObj.CurrentX = xPos
    xObj.CurrentY = 15 + BarH
    xParity = Parity(7) 'CInt(Mid(xstr, 1, 1)))
    
    
    For ii = 1 To Len(xstr)
        If ((Len(xstr) + 1) - ii) Mod 2 = 0 Then
            xTotal = xTotal + (CInt(Mid(xstr, ii, 1)))
        Else
            xTotal = xTotal + CInt(Mid(xstr, ii, 1) * 3)
        End If
        If ii = 5 Then
            draw_Bar "01010", True
        End If
        jj = CInt(Mid(xstr, ii, 1))
        If ii < 5 Then
            draw_Bar CStr(LeftHand_Odd(jj)), False
        ElseIf ii >= 5 Then
            draw_Bar CStr(Right_Hand(jj)), False
        End If
    Next
    chkSum = 0
    jj = xTotal Mod 10
    If jj <> 0 Then
        chkSum = 10 - jj
    End If
    draw_Bar CStr(Right_Hand(chkSum)), False
    draw_Bar "101", True
    
    If zHasCaption Then
        xObj.CurrentX = 23
        xObj.CurrentY = 10 + BarH
        xObj.Print Mid(xstr, 1, 4)
        
        xObj.CurrentX = 53
        xObj.CurrentY = 10 + BarH
        xObj.Print Mid(xstr, 5, 4) & chkSum
    End If
    
End Sub

Private Sub draw_Bar(Encoding As String, Guard As Boolean)

    Dim ii As Integer
    For ii = 1 To Len(Encoding)
        xPos = xPos + 1
        xObj.Line (xPos + 10, xtop)-(xPos + 10, xtop + BarH + IIf(Guard, 5, 0)), IIf(Mid(Encoding, ii, 1), vbBlack, vbWhite)
    Next
    
End Sub

Private Sub init_Table()

    LeftHand_Odd = Array("0001101", "0011001", "0010011", "0111101", "0100011", "0110001", "0101111", "0111011", "0110111", "0001011")
    LeftHand_Even = Array("0100111", "0110011", "0011011", "0100001", "0011101", "0111001", "0000101", "0010001", "0001001", "0010111")
    Right_Hand = Array("1110010", "1100110", "1101100", "1000010", "1011100", "1001110", "1010000", "1000100", "1001000", "1110100")
    Parity = Array("OOOOOO", "OOEOEE", "OOEEOE", "OOEEEO", "OEOOEE", "OEEOOE", "OEEEOO", "OEOEOE", "OEOEEO", "OEEOEO")

End Sub


