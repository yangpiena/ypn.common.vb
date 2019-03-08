Attribute VB_Name = "ModJSON"
'---------------------------------------------------------------------------------------
' Module    : ModJSON
' Author    : YPN
' Date      : 2018-4-5
' Purpose   : JSON工具类
'---------------------------------------------------------------------------------------

Option Explicit
Public Const INVALID_JSON      As Long = 1
Public Const INVALID_OBJECT    As Long = 2
Public Const INVALID_ARRAY     As Long = 3
Public Const INVALID_BOOLEAN   As Long = 4
Public Const INVALID_NULL      As Long = 5
Public Const INVALID_KEY       As Long = 6
Public Const INVALID_RPC_CALL  As Long = 7
Private m_Errors               As String


'---------------------------------------------------------------------------------------
' Procedure : MJSONParse
' Author    : YPN
' Date      : 2019/03/08 11:39
' Purpose   : JSON解析
' Param     : i_JSONString JSON格式源数据
'             i_JSONPath   数据访问路径
' Return    : Variant
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MJSONParse(ByVal i_JSONString As String, ByVal i_JSONPath As String) As Variant
    Dim v_JSON As Object
    On Error GoTo MJSONParse_Error
    
    Set v_JSON = CreateObject("MSScriptControl.ScriptControl")
    v_JSON.Language = "JScript"
    MJSONParse = v_JSON.eval("JSON=" & i_JSONString & ";JSON." & i_JSONPath & ";")
    Set v_JSON = Nothing
    
    On Error GoTo 0
    Exit Function
    
MJSONParse_Error:
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MJSONParse of Module ModJSON"
End Function

'---------------------------------------------------------------------------------------
' Procedure : MJSONAnalyze
' Author    : YPN
' Date      : 2018-4-5
' Purpose   : JSON解析
' Param     : i_JSONString 待解析的JSON字符串
'           : i_JSONKey    解析的关键字
' Return    : String
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MJSONAnalyze(ByVal i_JSONString As String, ByVal i_JSONKey As String) As String
    
    Dim v_JSON     As Object
    Dim v_JsonData
    Dim v_JsonTmp  As Object
    
    On Error GoTo MJSONAnalyze_Error
    
    Set v_JSON = parse(i_JSONString)
    v_JsonData = Split(i_JSONKey, ".")
    
    If IsArray(v_JsonData) And Not v_JSON Is Nothing Then
        Set v_JsonTmp = v_JSON
        
        For i = 0 To UBound(v_JsonData) - 1
            Set v_JsonTmp = v_JsonTmp.Item(CStr(v_JsonData(i)))
        Next
        
        MJSONAnalyze = v_JsonTmp.Item(CStr(v_JsonData(UBound(v_JsonData))))
    Else
        MJSONAnalyze = m_Errors
    End If
    
    On Error GoTo 0
    Exit Function
    
MJSONAnalyze_Error:
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MJSONAnalyze of Module ModJSON. " & m_Errors
    
End Function

Private Sub Class_Initialize()
    
    m_Errors = ""
    
End Sub


'   parse string and create JSON object
Private Function parse(ByRef i_str As String) As Object
    
    Dim index As Long
    
    index = 1
    m_Errors = ""
    
    On Error Resume Next
    
    Call skipChar(i_str, index)
    Select Case Mid(i_str, index, 1)
    Case "{"
        Set parse = parseObject(i_str, index)
    Case "["
        Set parse = parseArray(i_str, index)
    Case Else
        m_Errors = "Invalid JSON"
    End Select
    
End Function

'   parse collection of key/value
Private Function parseObject(ByRef i_str As String, ByRef i_Index As Long) As Dictionary
    
    Set parseObject = New Dictionary
    Dim sKey As String
    
    ' "{"
    Call skipChar(i_str, i_Index)
    If Mid(i_str, i_Index, 1) <> "{" Then
        m_Errors = m_Errors & "Invalid Object at position " & i_Index & " : " & Mid(i_str, i_Index) & vbCrLf
        Exit Function
    End If
    
    i_Index = i_Index + 1
    
    Do
        Call skipChar(i_str, i_Index)
        If "}" = Mid(i_str, i_Index, 1) Then
            i_Index = i_Index + 1
            Exit Do
        ElseIf "," = Mid(i_str, i_Index, 1) Then
            i_Index = i_Index + 1
            Call skipChar(i_str, i_Index)
        ElseIf i_Index > Len(i_str) Then
            m_Errors = m_Errors & "Missing '}': " & Right(i_str, 20) & vbCrLf
            Exit Do
        End If
        
        
        ' add key/value pair
        sKey = parseKey(i_str, i_Index)
        On Error Resume Next
        
        parseObject.Add sKey, parseValue(i_str, i_Index)
        If Err.Number <> 0 Then
            m_Errors = m_Errors & Err.Description & ": " & sKey & vbCrLf
            Exit Do
        End If
    Loop
eh:
    
End Function

'   parse list
Private Function parseArray(ByRef i_str As String, ByRef i_Index As Long) As Collection
    
    Set parseArray = New Collection
    
    ' "["
    Call skipChar(i_str, i_Index)
    If Mid(i_str, i_Index, 1) <> "[" Then
        m_Errors = m_Errors & "Invalid Array at position " & i_Index & " : " + Mid(i_str, i_Index, 20) & vbCrLf
        Exit Function
    End If
    
    i_Index = i_Index + 1
    
    Do
        Call skipChar(i_str, i_Index)
        If "]" = Mid(i_str, i_Index, 1) Then
            i_Index = i_Index + 1
            Exit Do
        ElseIf "," = Mid(i_str, i_Index, 1) Then
            i_Index = i_Index + 1
            Call skipChar(i_str, i_Index)
        ElseIf i_Index > Len(i_str) Then
            m_Errors = m_Errors & "Missing ']': " & Right(i_str, 20) & vbCrLf
            Exit Do
        End If
        
        ' add value
        On Error Resume Next
        parseArray.Add parseValue(i_str, i_Index)
        If Err.Number <> 0 Then
            m_Errors = m_Errors & Err.Description & ": " & Mid(i_str, i_Index, 20) & vbCrLf
            Exit Do
        End If
    Loop
    
End Function

'   parse string / number / object / array / true / false / null
Private Function parseValue(ByRef i_str As String, ByRef i_Index As Long)
    
    Call skipChar(i_str, i_Index)
    
    Select Case Mid(i_str, i_Index, 1)
    Case "{"
        Set parseValue = parseObject(i_str, i_Index)
    Case "["
        Set parseValue = parseArray(i_str, i_Index)
    Case """", "'"
        parseValue = parseString(i_str, i_Index)
    Case "t", "f"
        parseValue = parseBoolean(i_str, i_Index)
    Case "n"
        parseValue = parseNull(i_str, i_Index)
    Case Else
        parseValue = parseNumber(i_str, i_Index)
    End Select
    
End Function

'
'   parse string
'
Private Function parseString(ByRef i_str As String, ByRef i_Index As Long) As String
    
    Dim quote   As String
    Dim Char    As String
    Dim Code    As String
    
    Dim SB As New ClsStringBuilder
    
    Call skipChar(i_str, i_Index)
    quote = Mid(i_str, i_Index, 1)
    i_Index = i_Index + 1
    
    Do While i_Index > 0 And i_Index <= Len(i_str)
        Char = Mid(i_str, i_Index, 1)
        Select Case (Char)
        Case "\"
            i_Index = i_Index + 1
            Char = Mid(i_str, i_Index, 1)
            Select Case (Char)
            Case """", "\", "/", "'"
                SB.Append Char
                i_Index = i_Index + 1
            Case "b"
                SB.Append vbBack
                i_Index = i_Index + 1
            Case "f"
                SB.Append vbFormFeed
                i_Index = i_Index + 1
            Case "n"
                SB.Append vbLf
                i_Index = i_Index + 1
            Case "r"
                SB.Append vbCr
                i_Index = i_Index + 1
            Case "t"
                SB.Append vbTab
                i_Index = i_Index + 1
            Case "u"
                i_Index = i_Index + 1
                Code = Mid(i_str, i_Index, 4)
                SB.Append ChrW(Val("&h" + Code))
                i_Index = i_Index + 4
            End Select
        Case quote
            i_Index = i_Index + 1
            
            parseString = SB.toString
            Set SB = Nothing
            
            Exit Function
            
        Case Else
            SB.Append Char
            i_Index = i_Index + 1
        End Select
    Loop
    
    parseString = SB.toString
    Set SB = Nothing
    
End Function

'
'   parse number
'
Private Function parseNumber(ByRef i_str As String, ByRef i_Index As Long)
    
    Dim Value   As String
    Dim Char    As String
    
    Call skipChar(i_str, i_Index)
    Do While i_Index > 0 And i_Index <= Len(i_str)
        Char = Mid(i_str, i_Index, 1)
        If InStr("+-0123456789.eE", Char) Then
            Value = Value & Char
            i_Index = i_Index + 1
        Else
            parseNumber = CDec(Value)
            Exit Function
        End If
    Loop
    
End Function

'
'   parse true / false
'
Private Function parseBoolean(ByRef i_str As String, ByRef i_Index As Long) As Boolean
    
    Call skipChar(i_str, i_Index)
    If Mid(i_str, i_Index, 4) = "true" Then
        parseBoolean = True
        i_Index = i_Index + 4
    ElseIf Mid(i_str, i_Index, 5) = "false" Then
        parseBoolean = False
        i_Index = i_Index + 5
    Else
        m_Errors = m_Errors & "Invalid Boolean at position " & i_Index & " : " & Mid(i_str, i_Index) & vbCrLf
    End If
    
End Function

'
'   parse null
'
Private Function parseNull(ByRef i_str As String, ByRef i_Index As Long)
    
    Call skipChar(i_str, i_Index)
    If Mid(i_str, i_Index, 4) = "null" Then
        parseNull = Null
        i_Index = i_Index + 4
    Else
        m_Errors = m_Errors & "Invalid null value at position " & i_Index & " : " & Mid(i_str, i_Index) & vbCrLf
    End If
    
End Function

Private Function parseKey(ByRef i_str As String, ByRef i_Index As Long) As String
    
    Dim dquote  As Boolean
    Dim squote  As Boolean
    Dim Char    As String
    
    Call skipChar(i_str, i_Index)
    Do While i_Index > 0 And i_Index <= Len(i_str)
        Char = Mid(i_str, i_Index, 1)
        Select Case (Char)
        Case """"
            dquote = Not dquote
            i_Index = i_Index + 1
            If Not dquote Then
                Call skipChar(i_str, i_Index)
                If Mid(i_str, i_Index, 1) <> ":" Then
                    m_Errors = m_Errors & "Invalid Key at position " & i_Index & " : " & parseKey & vbCrLf
                    Exit Do
                End If
            End If
        Case "'"
            squote = Not squote
            i_Index = i_Index + 1
            If Not squote Then
                Call skipChar(i_str, i_Index)
                If Mid(i_str, i_Index, 1) <> ":" Then
                    m_Errors = m_Errors & "Invalid Key at position " & i_Index & " : " & parseKey & vbCrLf
                    Exit Do
                End If
            End If
        Case ":"
            i_Index = i_Index + 1
            If Not dquote And Not squote Then
                Exit Do
            Else
                parseKey = parseKey & Char
            End If
        Case Else
            If InStr(vbCrLf & vbCr & vbLf & vbTab & " ", Char) Then
            Else
                parseKey = parseKey & Char
            End If
            i_Index = i_Index + 1
        End Select
    Loop
    
End Function

'
'   skip special character
'
Private Sub skipChar(ByRef i_str As String, ByRef i_Index As Long)
    
    Dim bComment As Boolean
    Dim bStartComment As Boolean
    Dim bLongComment As Boolean
    
    Do While i_Index > 0 And i_Index <= Len(i_str)
        Select Case Mid(i_str, i_Index, 1)
        Case vbCr, vbLf
            If Not bLongComment Then
                bStartComment = False
                bComment = False
            End If
            
        Case vbTab, " ", "(", ")"
            
        Case "/"
            If Not bLongComment Then
                If bStartComment Then
                    bStartComment = False
                    bComment = True
                Else
                    bStartComment = True
                    bComment = False
                    bLongComment = False
                End If
            Else
                If bStartComment Then
                    bLongComment = False
                    bStartComment = False
                    bComment = False
                End If
            End If
            
        Case "*"
            If bStartComment Then
                bStartComment = False
                bComment = True
                bLongComment = True
            Else
                bStartComment = True
            End If
            
        Case Else
            If Not bComment Then
                Exit Do
            End If
        End Select
        
        i_Index = i_Index + 1
    Loop
    
End Sub

Public Function toString(ByRef i_Obj As Variant) As String
    
    Dim SB As New ClsStringBuilder
    
    Select Case VarType(i_Obj)
    Case vbNull
        SB.Append "null"
    Case vbDate
        SB.Append """" & CStr(i_Obj) & """"
    Case vbString
        SB.Append """" & Encode(i_Obj) & """"
    Case vbObject
        
        Dim bFI As Boolean
        Dim i As Long
        
        bFI = True
        If TypeName(i_Obj) = "Dictionary" Then
            
            SB.Append "{"
            Dim keys
            keys = i_Obj.keys
            For i = 0 To i_Obj.Count - 1
                If bFI Then bFI = False Else SB.Append ","
                Dim key
                key = keys(i)
                SB.Append """" & key & """:" & toString(i_Obj.Item(key))
            Next i
            SB.Append "}"
            
        ElseIf TypeName(i_Obj) = "Collection" Then
            
            SB.Append "["
            Dim Value
            For Each Value In i_Obj
                If bFI Then bFI = False Else SB.Append ","
                SB.Append toString(Value)
            Next Value
            SB.Append "]"
            
        End If
    Case vbBoolean
        If i_Obj Then SB.Append "true" Else SB.Append "false"
    Case vbVariant, vbArray, vbArray + vbVariant
        Dim sEB
        SB.Append multiArray(i_Obj, 1, "", sEB)
    Case Else
        SB.Append Replace(i_Obj, ",", ".")
    End Select
    
    toString = SB.toString
    Set SB = Nothing
    
End Function

Private Function Encode(i_str) As String
    
    Dim SB As New ClsStringBuilder
    Dim i As Long
    Dim j As Long
    Dim aL1 As Variant
    Dim aL2 As Variant
    Dim c As String
    Dim p As Boolean
    
    aL1 = Array(&H22, &H5C, &H2F, &H8, &HC, &HA, &HD, &H9)
    aL2 = Array(&H22, &H5C, &H2F, &H62, &H66, &H6E, &H72, &H74)
    For i = 1 To Len(i_str)
        p = True
        c = Mid(i_str, i, 1)
        For j = 0 To 7
            If c = Chr(aL1(j)) Then
                SB.Append "\" & Chr(aL2(j))
                p = False
                Exit For
            End If
        Next
        
        If p Then
            Dim a
            a = AscW(c)
            If a > 31 And a < 127 Then
                SB.Append c
            ElseIf a > -1 Or a < 65535 Then
                SB.Append "\u" & String(4 - Len(Hex(a)), "0") & Hex(a)
            End If
        End If
    Next
    
    Encode = SB.toString
    Set SB = Nothing
    
End Function

Private Function multiArray(i_ArrayBody, i_BaseCount, i_PoSition, ByRef i_PT)   ' Array BoDy, Integer BaseCount, String PoSition
    
    Dim iDU As Long
    Dim iDL As Long
    Dim i As Long
    
    On Error Resume Next
    iDL = LBound(i_ArrayBody, i_BaseCount)
    iDU = UBound(i_ArrayBody, i_BaseCount)
    
    Dim SB As New ClsStringBuilder
    
    Dim sPB1, sPB2  ' String PointBuffer1, String PointBuffer2
    If Err.Number = 9 Then
        sPB1 = i_PT & i_PoSition
        For i = 1 To Len(sPB1)
            If i <> 1 Then sPB2 = sPB2 & ","
            sPB2 = sPB2 & Mid(sPB1, i, 1)
        Next
        '        multiArray = multiArray & toString(Eval("i_ArrayBody(" & sPB2 & ")"))
        SB.Append toString(i_ArrayBody(sPB2))
    Else
        i_PT = i_PT & i_PoSition
        SB.Append "["
        For i = iDL To iDU
            SB.Append multiArray(i_ArrayBody, i_BaseCount + 1, i, i_PT)
            If i < iDU Then SB.Append ","
        Next
        SB.Append "]"
        i_PT = Left(i_PT, i_BaseCount - 2)
    End If
    Err.Clear
    multiArray = SB.toString
    
    Set SB = Nothing
    
End Function
