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

Private psErrors               As String


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
    Dim v_Json     As Object
    Dim v_JsonData
    Dim v_JsonTmp  As Object
    
   On Error GoTo MJSONAnalyze_Error

    Set v_Json = parse(i_JSONString)
    v_JsonData = Split(i_JSONKey, ".")
    
    If IsArray(v_JsonData) Then
        Set v_JsonTmp = v_Json
        
        For i = 0 To UBound(v_JsonData) - 1
            Set v_JsonTmp = v_JsonTmp.Item(CStr(v_JsonData(i)))
        Next
        
        MJSONAnalyze = v_JsonTmp.Item(CStr(v_JsonData(UBound(v_JsonData))))
    End If

   On Error GoTo 0
   Exit Function

MJSONAnalyze_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MJSONAnalyze of Module ModJSON"
End Function

Private Sub Class_Initialize()
    psErrors = ""
End Sub


'   parse string and create JSON object
Private Function parse(ByRef str As String) As Object
    Dim index As Long
    
    index = 1
    psErrors = ""
    
    On Error Resume Next
    
    Call skipChar(str, index)
    Select Case Mid(str, index, 1)
    Case "{"
        Set parse = parseObject(str, index)
    Case "["
        Set parse = parseArray(str, index)
    Case Else
        psErrors = "Invalid JSON"
    End Select
End Function

'   parse collection of key/value
Private Function parseObject(ByRef str As String, ByRef index As Long) As Dictionary
    Set parseObject = New Dictionary
    Dim sKey As String
    
    ' "{"
    Call skipChar(str, index)
    If Mid(str, index, 1) <> "{" Then
        psErrors = psErrors & "Invalid Object at position " & index & " : " & Mid(str, index) & vbCrLf
        Exit Function
    End If
    
    index = index + 1
    
    Do
        Call skipChar(str, index)
        If "}" = Mid(str, index, 1) Then
            index = index + 1
            Exit Do
        ElseIf "," = Mid(str, index, 1) Then
            index = index + 1
            Call skipChar(str, index)
        ElseIf index > Len(str) Then
            psErrors = psErrors & "Missing '}': " & Right(str, 20) & vbCrLf
            Exit Do
        End If
        
        
        ' add key/value pair
        sKey = parseKey(str, index)
        On Error Resume Next
        
        parseObject.Add sKey, parseValue(str, index)
        If Err.Number <> 0 Then
            psErrors = psErrors & Err.Description & ": " & sKey & vbCrLf
            Exit Do
        End If
    Loop
eh:
    
End Function

'   parse list
Private Function parseArray(ByRef str As String, ByRef index As Long) As Collection
    Set parseArray = New Collection
    
    ' "["
    Call skipChar(str, index)
    If Mid(str, index, 1) <> "[" Then
        psErrors = psErrors & "Invalid Array at position " & index & " : " + Mid(str, index, 20) & vbCrLf
        Exit Function
    End If
    
    index = index + 1
    
    Do
        
        Call skipChar(str, index)
        If "]" = Mid(str, index, 1) Then
            index = index + 1
            Exit Do
        ElseIf "," = Mid(str, index, 1) Then
            index = index + 1
            Call skipChar(str, index)
        ElseIf index > Len(str) Then
            psErrors = psErrors & "Missing ']': " & Right(str, 20) & vbCrLf
            Exit Do
        End If
        
        ' add value
        On Error Resume Next
        parseArray.Add parseValue(str, index)
        If Err.Number <> 0 Then
            psErrors = psErrors & Err.Description & ": " & Mid(str, index, 20) & vbCrLf
            Exit Do
        End If
    Loop
End Function

'   parse string / number / object / array / true / false / null
Private Function parseValue(ByRef str As String, ByRef index As Long)
    Call skipChar(str, index)
    
    Select Case Mid(str, index, 1)
    Case "{"
        Set parseValue = parseObject(str, index)
    Case "["
        Set parseValue = parseArray(str, index)
    Case """", "'"
        parseValue = parseString(str, index)
    Case "t", "f"
        parseValue = parseBoolean(str, index)
    Case "n"
        parseValue = parseNull(str, index)
    Case Else
        parseValue = parseNumber(str, index)
    End Select
End Function

'
'   parse string
'
Private Function parseString(ByRef str As String, ByRef index As Long) As String
    Dim quote   As String
    Dim Char    As String
    Dim Code    As String
    
    Dim SB As New ClsStringBuilder
    
    Call skipChar(str, index)
    quote = Mid(str, index, 1)
    index = index + 1
    
    Do While index > 0 And index <= Len(str)
        Char = Mid(str, index, 1)
        Select Case (Char)
        Case "\"
            index = index + 1
            Char = Mid(str, index, 1)
            Select Case (Char)
            Case """", "\", "/", "'"
                SB.Append Char
                index = index + 1
            Case "b"
                SB.Append vbBack
                index = index + 1
            Case "f"
                SB.Append vbFormFeed
                index = index + 1
            Case "n"
                SB.Append vbLf
                index = index + 1
            Case "r"
                SB.Append vbCr
                index = index + 1
            Case "t"
                SB.Append vbTab
                index = index + 1
            Case "u"
                index = index + 1
                Code = Mid(str, index, 4)
                SB.Append ChrW(Val("&h" + Code))
                index = index + 4
            End Select
        Case quote
            index = index + 1
            
            parseString = SB.toString
            Set SB = Nothing
            
            Exit Function
            
        Case Else
            SB.Append Char
            index = index + 1
        End Select
    Loop
    
    parseString = SB.toString
    Set SB = Nothing
End Function

'
'   parse number
'
Private Function parseNumber(ByRef str As String, ByRef index As Long)
    Dim Value   As String
    Dim Char    As String
    
    Call skipChar(str, index)
    Do While index > 0 And index <= Len(str)
        Char = Mid(str, index, 1)
        If InStr("+-0123456789.eE", Char) Then
            Value = Value & Char
            index = index + 1
        Else
            parseNumber = CDec(Value)
            Exit Function
        End If
    Loop
End Function

'
'   parse true / false
'
Private Function parseBoolean(ByRef str As String, ByRef index As Long) As Boolean
    Call skipChar(str, index)
    If Mid(str, index, 4) = "true" Then
        parseBoolean = True
        index = index + 4
    ElseIf Mid(str, index, 5) = "false" Then
        parseBoolean = False
        index = index + 5
    Else
        psErrors = psErrors & "Invalid Boolean at position " & index & " : " & Mid(str, index) & vbCrLf
    End If
End Function

'
'   parse null
'
Private Function parseNull(ByRef str As String, ByRef index As Long)
    Call skipChar(str, index)
    If Mid(str, index, 4) = "null" Then
        parseNull = Null
        index = index + 4
    Else
        psErrors = psErrors & "Invalid null value at position " & index & " : " & Mid(str, index) & vbCrLf
    End If
End Function

Private Function parseKey(ByRef str As String, ByRef index As Long) As String
    Dim dquote  As Boolean
    Dim squote  As Boolean
    Dim Char    As String
    
    Call skipChar(str, index)
    Do While index > 0 And index <= Len(str)
        Char = Mid(str, index, 1)
        Select Case (Char)
        Case """"
            dquote = Not dquote
            index = index + 1
            If Not dquote Then
                Call skipChar(str, index)
                If Mid(str, index, 1) <> ":" Then
                    psErrors = psErrors & "Invalid Key at position " & index & " : " & parseKey & vbCrLf
                    Exit Do
                End If
            End If
        Case "'"
            squote = Not squote
            index = index + 1
            If Not squote Then
                Call skipChar(str, index)
                If Mid(str, index, 1) <> ":" Then
                    psErrors = psErrors & "Invalid Key at position " & index & " : " & parseKey & vbCrLf
                    Exit Do
                End If
            End If
        Case ":"
            index = index + 1
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
            index = index + 1
        End Select
    Loop
End Function

'
'   skip special character
'
Private Sub skipChar(ByRef str As String, ByRef index As Long)
    Dim bComment As Boolean
    Dim bStartComment As Boolean
    Dim bLongComment As Boolean
    
    Do While index > 0 And index <= Len(str)
        Select Case Mid(str, index, 1)
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
        
        index = index + 1
    Loop
End Sub

Public Function toString(ByRef obj As Variant) As String
    Dim SB As New ClsStringBuilder
    
    Select Case VarType(obj)
    Case vbNull
        SB.Append "null"
    Case vbDate
        SB.Append """" & CStr(obj) & """"
    Case vbString
        SB.Append """" & Encode(obj) & """"
    Case vbObject
        
        Dim bFI As Boolean
        Dim i As Long
        
        bFI = True
        If TypeName(obj) = "Dictionary" Then
            
            SB.Append "{"
            Dim keys
            keys = obj.keys
            For i = 0 To obj.Count - 1
                If bFI Then bFI = False Else SB.Append ","
                Dim key
                key = keys(i)
                SB.Append """" & key & """:" & toString(obj.Item(key))
            Next i
            SB.Append "}"
            
        ElseIf TypeName(obj) = "Collection" Then
            
            SB.Append "["
            Dim Value
            For Each Value In obj
                If bFI Then bFI = False Else SB.Append ","
                SB.Append toString(Value)
            Next Value
            SB.Append "]"
            
        End If
    Case vbBoolean
        If obj Then SB.Append "true" Else SB.Append "false"
    Case vbVariant, vbArray, vbArray + vbVariant
        Dim sEB
        SB.Append multiArray(obj, 1, "", sEB)
    Case Else
        SB.Append Replace(obj, ",", ".")
    End Select
    
    toString = SB.toString
    Set SB = Nothing
End Function

Private Function Encode(str) As String
    Dim SB As New ClsStringBuilder
    Dim i As Long
    Dim j As Long
    Dim aL1 As Variant
    Dim aL2 As Variant
    Dim c As String
    Dim p As Boolean
    
    aL1 = Array(&H22, &H5C, &H2F, &H8, &HC, &HA, &HD, &H9)
    aL2 = Array(&H22, &H5C, &H2F, &H62, &H66, &H6E, &H72, &H74)
    For i = 1 To Len(str)
        p = True
        c = Mid(str, i, 1)
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

Private Function multiArray(aBD, iBC, sPS, ByRef sPT)   ' Array BoDy, Integer BaseCount, String PoSition
    Dim iDU As Long
    Dim iDL As Long
    Dim i As Long
    
    On Error Resume Next
    iDL = LBound(aBD, iBC)
    iDU = UBound(aBD, iBC)
    
    Dim SB As New ClsStringBuilder
    
    Dim sPB1, sPB2  ' String PointBuffer1, String PointBuffer2
    If Err.Number = 9 Then
        sPB1 = sPT & sPS
        For i = 1 To Len(sPB1)
            If i <> 1 Then sPB2 = sPB2 & ","
            sPB2 = sPB2 & Mid(sPB1, i, 1)
        Next
        '        multiArray = multiArray & toString(Eval("aBD(" & sPB2 & ")"))
        SB.Append toString(aBD(sPB2))
    Else
        sPT = sPT & sPS
        SB.Append "["
        For i = iDL To iDU
            SB.Append multiArray(aBD, iBC + 1, i, sPT)
            If i < iDU Then SB.Append ","
        Next
        SB.Append "]"
        sPT = Left(sPT, iBC - 2)
    End If
    Err.Clear
    multiArray = SB.toString
    
    Set SB = Nothing
End Function
