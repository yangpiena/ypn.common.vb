Attribute VB_Name = "ModGetInitial"
'---------------------------------------------------------------------------------------
' Module    : ModGetInitial
' Author    : YPN
' Date      : 2017-06-28 16:57
' Purpose   : 获取首字母
'---------------------------------------------------------------------------------------

Option Explicit



'---------------------------------------------------------------------------------------
' Procedure : MGetInitial
' Author    : YPN
' Date      : 2017-06-28 17:07
' Purpose   : 获取第一个汉字的首字母
' Param     : i_Str 汉字字符串
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MGetInitialFirst(ByVal i_Str As String) As String

    If i_Str = "" Then Exit Function
    
    MGetInitialFirst = GetPinyin(Left(i_Str, 1))
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : MGetInitialAll
' Author    : YPN
' Date      : 2017-06-28 17:04
' Purpose   : 获取所有汉字的首字母
' Param     : i_Str 汉字字符串
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MGetInitialAll(ByVal i_Str As String) As String

    If i_Str = "" Then Exit Function
        
    For i = 1 To Len(i_Str)
        MGetInitialAll = MGetInitialAll & GetPinyin(Mid(i_Str, i, 1))
    Next i
    
End Function



'---------------------------------------------------------------------------------------
' Procedure : GetPinyin
' Author    : YPN
' Date      : 2017-06-28 17:04
' Purpose   : 获取汉字的拼音
' Param     : i_Str 汉字
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Private Function GetPinyin(ByVal i_Str As String) As String

    Dim v_Pinyin As String
    
    i_Str = Hex(Asc(i_Str))  ' 将汉字转换为其内码的十六进制字符串
    
    Select Case i_Str
    Case "B0A1" To "B0C4": v_Pinyin = "A"
    Case "B0C5" To "B2C0": v_Pinyin = "B"
    Case "B2C1" To "B4ED": v_Pinyin = "C"
    Case "B4EE" To "B6E9": v_Pinyin = "D"
    Case "B6EA" To "B7A1": v_Pinyin = "E"
    Case "B7A2" To "B8C0": v_Pinyin = "F"
    Case "B8C1" To "B9FD": v_Pinyin = "G"
    Case "B9FE" To "BBF6": v_Pinyin = "H"
    Case "BBF7" To "BFA5": v_Pinyin = "J"
    Case "BFA6" To "C0AB": v_Pinyin = "K"
    Case "C0AC" To "C2E7": v_Pinyin = "L"
    Case "C2E8" To "C4C2": v_Pinyin = "M"
    Case "C4C3" To "C5B5": v_Pinyin = "N"
    Case "C5B6" To "C5BD": v_Pinyin = "O"
    Case "C5BE" To "C6D9": v_Pinyin = "P"
    Case "C6DA" To "C8BA": v_Pinyin = "Q"
    Case "C8BB" To "C8F5": v_Pinyin = "R"
    Case "C8F6" To "CBF9": v_Pinyin = "S"
    Case "CBFA" To "CDD9": v_Pinyin = "T"
    Case "CDDA" To "CEF3": v_Pinyin = "W"
    Case "CEF4" To "D188": v_Pinyin = "X"
    Case "D189" To "D4D0": v_Pinyin = "Y"
    Case "D4D1" To "D7F9": v_Pinyin = "Z"
    Case Else
        v_Pinyin = ""
    End Select
    
    GetPinyin = v_Pinyin
    
End Function
