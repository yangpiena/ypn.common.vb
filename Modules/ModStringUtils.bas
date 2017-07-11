Attribute VB_Name = "ModStringUtils"
'---------------------------------------------------------------------------------------
' Module    : ModStringUtils
' Author    : YPN
' Date      : 2017-06-29 14:46
' Purpose   : �ַ���������
'---------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : MIsNull
' Author    : YPN
' Date      : 2017-06-29 14:51
' Purpose   : �жϱ����Ƿ�Ϊ��
' Param     : i_Var ����
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MIsNull(ByVal i_Var As Variant) As Boolean

    If isNull(i_Var) Then
        MIsNull = True
        Exit Function
    End If

    If Trim(i_Var) = "" Then
        MIsNull = True
        Exit Function
    End If
    
    MIsNull = False
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : MGetFileNameInPath
' Author    : YPN
' Date      : 2017-06-28 17:45
' Purpose   : ��ָ��ȫ·���л�ȡ�ļ���
' Param     : i_Path      ָ��ȫ·��
'             i_HasSuffix �Ƿ������׺����Ĭ�ϰ���
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MGetFileNameInPath(ByVal i_Path As String, Optional ByVal i_NeedSuffix As Boolean = False) As String

    Dim v_FileName As String, v_FileNameNoSuffix As String
    
    i_Path = Trim(i_Path)
    i = InStrRev(i_Path, "\")
    j = Len(i_Path)
    If i = 0 Then Exit Function
    
    v_FileName = Mid(i_Path, i + 1, j - i)
    
    i = InStrRev(v_FileName, ".")
    j = Len(v_FileName)
    If i = 0 Then Exit Function
    
    v_FileNameNoSuffix = Mid(v_FileName, 1, i - 1)
    
    If i_NeedSuffix Then
        MGetFileNameInPath = v_FileName
    Else
        MGetFileNameInPath = v_FileNameNoSuffix
    End If
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : MGetSuffixInFileName
' Author    : YPN
' Date      : 2017-06-28 17:50
' Purpose   : ���ļ����л�ȡ��׺��
' Param     : i_FileName �ļ���
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MGetSuffixInFileName(ByVal i_FileName As String) As String

    MGetSuffixInFileName = IIf(InStr(i_FileName, "."), Right(i_FileName, Len(i_FileName) - InStrRev(i_FileName, ".")), vbNullString)

End Function

'---------------------------------------------------------------------------------------
' Procedure : MGetInitial
' Author    : YPN
' Date      : 2017-06-28 17:07
' Purpose   : ��ȡ��һ�����ֵ�����ĸ
' Param     : i_Str �����ַ���
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MGetInitialFirst(ByVal i_Str As String) As String

    If i_Str = "" Then Exit Function
    
    MGetInitialFirst = getPinyin(Left(i_Str, 1))
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : MGetInitialAll
' Author    : YPN
' Date      : 2017-06-28 17:04
' Purpose   : ��ȡ���к��ֵ�����ĸ
' Param     : i_Str �����ַ���
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MGetInitialAll(ByVal i_Str As String) As String

    If i_Str = "" Then Exit Function
        
    For i = 1 To Len(i_Str)
        MGetInitialAll = MGetInitialAll & getPinyin(Mid(i_Str, i, 1))
    Next i
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : MGetYear2
' Author    : YPN
' Date      : 2017-07-10 17:06
' Purpose   : ��ȡ�����е���ݺ�2λ
' Param     : i_Date ����
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MGetYear2(ByVal i_Date As String) As Integer

    MGetYear2 = Right(CStr(Year(i_Date)), 2)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : MHexToText
' Author    : YPN
' Date      : 2017-07-05 15:55
' Purpose   : ��16���Ʊ��봮ת��Ϊ�ı���û��д�쳣������ֻҪ���� TextToHex() ת���Ľ����û����
' Param     : i_Code Ҫת����16���Ʊ���
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MHexToText(i_Code As String) As String

    Dim aBuffer() As Byte
    Dim i As Long, n As Long
    
    n = Len(i_Code) \ 2 - 1
    ReDim aBuffer(n)
    For i = 0 To n
        aBuffer(i) = CByte("&H" & Mid$(i_Code, i + i + 1, 2))
    Next
    MHexToText = StrConv(aBuffer, vbUnicode)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : MTextToHex
' Author    : YPN
' Date      : 2017-07-05 15:54
' Purpose   : ���ı�ת��Ϊ16���Ʊ��봮
' Param     : i_Text Ҫת�����ı�
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MTextToHex(i_Text As String) As String

    Dim aBuffer() As Byte
    Dim strOut As String
    Dim i As Long, p As Long
    
    aBuffer = StrConv(i_Text, vbFromUnicode)
    i = UBound(aBuffer)
    strOut = Space$(i + i + 2)
    p = 1
    For i = 0 To i
        Mid$(strOut, p) = Right$("0" & Hex$(aBuffer(i)), 2)
        p = p + 2
    Next
    MTextToHex = strOut
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetPinyin
' Author    : YPN
' Date      : 2017-06-28 17:04
' Purpose   : ��ȡ���ֵ�ƴ��
' Param     : i_Str ����
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Private Function getPinyin(ByVal i_Str As String) As String

    Dim v_Pinyin As String
    
    i_Str = Hex(Asc(i_Str))  ' ������ת��Ϊ�������ʮ�������ַ���
    
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
    
    getPinyin = v_Pinyin
    
End Function

