Attribute VB_Name = "ModStringUtils"
'---------------------------------------------------------------------------------------
' Module    : ModStringUtils
' Author    : YPN
' Date      : 2017-06-29 14:46
' Purpose   : �ַ���������
'---------------------------------------------------------------------------------------

Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : MBase64ToText
' Author    : YPN
' Date      : 2017-12-13 11:29
' Purpose   : ��Base64λ�����ַ���ת��Ϊ�ı�
' Param     : i_Text Ҫת�����ı�
' Return    :
' Remark    : Base64 ����
'---------------------------------------------------------------------------------------
'
Public Function MBase64ToText(ByVal i_Base64 As String) As String
    
    Dim v_OutStr() As Byte
    Dim v_Length   As Long, v_Mods As Long
    Const B64_CHAR_DICT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
    
    On Error GoTo Base64Decode_Error
    
    If InStr(1, i_Base64, "=") <> 0 Then i_Base64 = Left(i_Base64, InStr(1, i_Base64, "=") - 1)     '�ж�Base64��ʵ����,��ȥ��λ
    v_Mods = Len(i_Base64) Mod 4
    v_Length = Len(i_Base64) - v_Mods
    ReDim v_OutStr(v_Length / 4 * 3 - 1 + Switch(v_Mods = 0, 0, v_Mods = 2, 1, v_Mods = 3, 2))
    
    For i = 1 To v_Length Step 4
        Dim buf(3) As Byte
        For j = 0 To 3
            buf(j) = InStr(1, B64_CHAR_DICT, Mid(i_Base64, i + j, 1)) - 1                           '�����ַ���λ��ȡ������ֵ
        Next
        v_OutStr((i - 1) / 4 * 3) = buf(0) * &H4 + (buf(1) And &H30) / &H10
        v_OutStr((i - 1) / 4 * 3 + 1) = (buf(1) And &HF) * &H10 + (buf(2) And &H3C) / &H4
        v_OutStr((i - 1) / 4 * 3 + 2) = (buf(2) And &H3) * &H40 + buf(3)
    Next
    If v_Mods = 2 Then
        v_OutStr(v_Length / 4 * 3) = (InStr(1, B64_CHAR_DICT, Mid(i_Base64, v_Length + 1, 1)) - 1) * &H4 + ((InStr(1, B64_CHAR_DICT, Mid(i_Base64, v_Length + 2, 1)) - 1) And &H30) / 16
    ElseIf v_Mods = 3 Then
        v_OutStr(v_Length / 4 * 3) = (InStr(1, B64_CHAR_DICT, Mid(i_Base64, v_Length + 1, 1)) - 1) * &H4 + ((InStr(1, B64_CHAR_DICT, Mid(i_Base64, v_Length + 2, 1)) - 1) And &H30) / 16
        v_OutStr(v_Length / 4 * 3 + 1) = ((InStr(1, B64_CHAR_DICT, Mid(i_Base64, v_Length + 2, 1)) - 1) And &HF) * &H10 + ((InStr(1, B64_CHAR_DICT, Mid(i_Base64, v_Length + 3, 1)) - 1) And &H3C) / &H4
    End If
    MBase64ToText = StrConv(v_OutStr, vbUnicode)                                                    '��ȡ������
    
    On Error GoTo 0
    Exit Function
    
Base64Decode_Error:
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Base64Decode of Module ModStringUtils"
    
End Function

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
    
    If IsNull(i_Var) Then
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
' Procedure : MGetLocalComputerName
' Author    : YPN
' Date      : 2018-04-08 16:00
' Purpose   : ��ȡ������������
' Param     :
' Return    : String
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MGetLocalComputerName() As String
    
    On Error GoTo MGetLocalComputerName_Error
    
    MGetLocalComputerName = Environ("computername")
    
    On Error GoTo 0
    Exit Function
    
MGetLocalComputerName_Error:
    
    ' MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MGetLocalComputerName of Module ModStringUtils"
    MGetLocalComputerName = "��ȡ����"
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : MGetLocalIP
' Author    : YPN
' Date      : 2018-04-08 16:00
' Purpose   : ��ȡ����IP��ַ
' Param     :
' Return    : String
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MGetLocalIP() As String
    
    Dim v_winIP As Object
    
    On Error GoTo MGetLocalIP_Error
    
    Set v_winIP = CreateObject("MSWinsock.Winsock")
    MGetLocalIP = v_winIP.localip
    
    On Error GoTo 0
    Exit Function
    
MGetLocalIP_Error:
    
    ' MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MGetLocalIP of Module ModStringUtils"
    MGetLocalIP = "��ȡ����"
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : MGetLocalUserName
' Author    : YPN
' Date      : 2018-04-08 16:00
' Purpose   : ��ȡ�����û�����
' Param     :
' Return    : String
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MGetLocalUserName() As String
    
    On Error GoTo MGetLocalUserName_Error
    
    MGetLocalUserName = Environ("username")
    
    On Error GoTo 0
    Exit Function
    
MGetLocalUserName_Error:
    
    ' MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MGetLocalUserName of Module ModStringUtils"
    MGetLocalUserName = "��ȡ����"
    
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
' Procedure : MGetMonthBegin
' Author    : YPN
' Date      : 2018/03/08 16:51
' Purpose   : ��ȡָ�����ڵ��³�����
' Param     :
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MGetMonthBegin(ByVal i_Date As String) As String
    
    MGetMonthBegin = DateSerial(Year(i_Date), Month(i_Date), 1)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : MGetMonthEnd
' Author    : YPN
' Date      : 2018/03/08 16:52
' Purpose   : ��ȡָ�����ڵ���ĩ����
' Param     :
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MGetMonthEnd(ByVal i_Date As String) As String
    
    MGetMonthEnd = DateAdd("d", -1, DateSerial(Year(i_Date), Month(i_Date) + 1, 1))
    
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
' Procedure : MRndInt
' Author    : YPN
' Date      : 2018/03/05 16:55
' Purpose   : ����ָ����Χ�ڵ��������
' Param     : i_Min ��С��
'             i_Max �����
' Return    : �������
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MRndInt(ByVal i_Min As Long, ByVal i_Max As Long) As Integer
    
    Randomize
    MRndInt = Int(Rnd() * (i_Max - i_Min + 1)) + i_Min
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : MTextToBase64
' Author    : YPN
' Date      : 2017-12-13 11:19
' Purpose   : ���ı�ת��ΪBase64λ�����ַ���
' Param     : i_Text Ҫת�����ı�
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MTextToBase64(ByVal i_text As String) As String
    
    Dim v_Str()  As Byte, v_Buf() As Byte
    Dim v_Length As Long, v_Mods  As Long
    Const B64_CHAR_DICT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
    
    On Error GoTo MTextToBase64_Error
    
    v_Str() = StrConv(i_text, vbFromUnicode)
    v_Mods = (UBound(v_Str) + 1) Mod 3    '����3������
    v_Length = UBound(v_Str) + 1 - v_Mods
    ReDim v_Buf(v_Length / 3 * 4 + IIf(v_Mods <> 0, 4, 0) - 1)
    
    For i = 0 To v_Length - 1 Step 3
        v_Buf(i / 3 * 4) = (v_Str(i) And &HFC) / &H4
        v_Buf(i / 3 * 4 + 1) = (v_Str(i) And &H3) * &H10 + (v_Str(i + 1) And &HF0) / &H10
        v_Buf(i / 3 * 4 + 2) = (v_Str(i + 1) And &HF) * &H4 + (v_Str(i + 2) And &HC0) / &H40
        v_Buf(i / 3 * 4 + 3) = v_Str(i + 2) And &H3F
    Next
    If v_Mods = 1 Then
        v_Buf(v_Length / 3 * 4) = (v_Str(v_Length) And &HFC) / &H4
        v_Buf(v_Length / 3 * 4 + 1) = (v_Str(v_Length) And &H3) * &H10
        v_Buf(v_Length / 3 * 4 + 2) = 64
        v_Buf(v_Length / 3 * 4 + 3) = 64
    ElseIf v_Mods = 2 Then
        v_Buf(v_Length / 3 * 4) = (v_Str(v_Length) And &HFC) / &H4
        v_Buf(v_Length / 3 * 4 + 1) = (v_Str(v_Length) And &H3) * &H10 + (v_Str(v_Length + 1) And &HF0) / &H10
        v_Buf(v_Length / 3 * 4 + 2) = (v_Str(v_Length + 1) And &HF) * &H4
        v_Buf(v_Length / 3 * 4 + 3) = 64
    End If
    For i = 0 To UBound(v_Buf)
        MTextToBase64 = MTextToBase64 + Mid(B64_CHAR_DICT, v_Buf(i) + 1, 1)
    Next
    
    On Error GoTo 0
    Exit Function
    
MTextToBase64_Error:
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MTextToBase64 of Module ModStringUtils"
    
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
Public Function MTextToHex(i_text As String) As String
    
    Dim aBuffer() As Byte
    Dim strOut As String
    Dim i As Long, p As Long
    
    aBuffer = StrConv(i_text, vbFromUnicode)
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
' Procedure : MTextToURL_GB2312
' Author    : YPN
' Date      : 2018-04-13 16:35
' Purpose   : �ı�����ת����URL��ʽ�ַ���������GB2312
' Param     :
' Return    : String
' Remark    : �˷������ڷ��ͷ��POST�У�������ת�������ķ��͡�
'---------------------------------------------------------------------------------------
'
Public Function MTextToURL_GB2312(i_text) As String
    
    Dim v_txt, v_numTxt, v_Str
    
    MTextToURL_GB2312 = ""
    
    For i = 1 To Len(i_text)
        v_txt = ""
        v_txt = Mid(i_text, i, 1)
        v_numTxt = "-,.0123456789/"
        If InStr(v_numTxt, v_txt) > 0 Then
            MTextToURL_GB2312 = MTextToURL_GB2312 & v_txt
        Else
            If Asc(v_txt) < 0 Then
                v_Str = "%" & Right(CStr(Hex(Asc(v_txt))), 2)
                v_Str = "%" & Left(CStr(Hex(Asc(v_txt))), Len(CStr(Hex(Asc(v_txt)))) - 2) & v_Str
                MTextToURL_GB2312 = MTextToURL_GB2312 & v_Str
            ElseIf Asc(v_txt) >= 65 And Asc(v_txt) <= 90 Then
                MTextToURL_GB2312 = MTextToURL_GB2312 & v_txt
            ElseIf Asc(v_txt) >= 97 And Asc(v_txt) <= 122 Then
                MTextToURL_GB2312 = MTextToURL_GB2312 & v_txt
            Else
                MTextToURL_GB2312 = MTextToURL_GB2312 & "%" & CStr(Hex(Asc(v_txt)))
            End If
        End If
    Next
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : MTextToURL_UTF8
' Author    : YPN
' Date      : 2018-04-13 16:36
' Purpose   : �ı�����ת����URL��ʽ�ַ���������UTF-8
' Param     :
' Return    : String
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MTextToURL_UTF8(i_text) As String
    
    Dim v_wch, v_uch, v_szRet
    Dim v_asc, v_asc2, v_asc3
    
    If i_text = "" Then
        MTextToURL_UTF8 = i_text
        Exit Function
    End If
    
    For i = 1 To Len(i_text)
        v_wch = Mid(i_text, i, 1)
        v_asc = AscW(v_wch)
        If v_asc < 0 Then v_asc = v_asc + 65536
        If (v_asc And &HFF80) = 0 Then
            v_szRet = v_szRet & v_wch
        Else
            If (v_asc And &HF000) = 0 Then
                v_uch = "%" & CStr(Hex(((v_asc \ 2 ^ 6)) Or &HC0)) & CStr(Hex(v_asc And &H3F Or &H80))
                v_szRet = v_szRet & v_uch
            Else
                v_uch = "%" & CStr(Hex((v_asc \ 2 ^ 12) Or &HE0)) & "%" & CStr(Hex((v_asc \ 2 ^ 6) And &H3F Or &H80)) & "%" & CStr(Hex(v_asc And &H3F Or &H80))
                v_szRet = v_szRet & v_uch
            End If
        End If
    Next
    
    MTextToURL_UTF8 = v_szRet
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : MTrimText
' Author    : YPN
' Date      : 2018/01/29 11:37
' Purpose   : ȥ���ַ������߿ո�֧��ȥ�����߿յĻ��з�
' Param     :
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MTrimText(i_text As String) As String
    
    'ȥ�����߻��з�
    i_text = Replace(i_text, " ", "YPN_TmpValue")
    i_text = Replace(i_text, vbCrLf, " ")
    i_text = Trim(i_text)
    i_text = Replace(i_text, " ", vbCrLf)
    i_text = Replace(i_text, "YPN_TmpValue", " ")
    'ȥ�����߿ո�
    i_text = Trim(i_text)
    
    MTrimText = i_text
    
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

