Attribute VB_Name = "ModFileUtils"
'---------------------------------------------------------------------------------------
' Module    : ModFileUtils
' Author    : YPN
' Date      : 2017-12-12 16:17
' Purpose   : 文件工具类
'---------------------------------------------------------------------------------------

Option Explicit

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


'---------------------------------------------------------------------------------------
' Procedure : MApendText
' Author    : YPN
' Date      : 2017-12-12 16:20
' Purpose   : 追加内容到指定文件
' Param     : i_TextFile     指定文件
'             i_ApendContent 追加内容
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Sub MApendText(ByVal i_TextFile As String, ByVal i_ApendContent As String)
    
    On Error GoTo MApendText_Error
    
    If Dir(i_TextFile) <> "" Then          ' 如果文件存在
        Open i_TextFile For Append As #1   ' 以追加方式打开文件
        'Print #1                          ' 为了防止原文件末尾没有换行，而加入的换行
        Print #1, i_ApendContent
        Close #1
    Else
        MsgBox "指定文件不存在：" & i_TextFile, vbExclamation, TS
        Exit Sub
    End If
    
    On Error GoTo 0
    Exit Sub
    
MApendText_Error:
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MApendText of Module ModFileUtils"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : MGetIniValue
' Author    : YPN
' Date      : 2017-12-15 11:23
' Purpose   : 获取初始化文件（.ini）指定键（Key）的值（Value）
' Param     : i_Section    节
'             i_Key        键
'             i_FileName   完整的INI文件名
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MGetIniValue(ByVal i_Section As String, ByVal i_Key As String, ByVal i_FileName As String) As String
    
    Dim v_Buff As String * 128
    
    On Error GoTo MGetIniValue_Error
    
    X = GetPrivateProfileString(i_Section, i_Key, "", v_Buff, 128, i_FileName)
    I = InStr(v_Buff, Chr(0))
    
    MGetIniValue = Trim(Left(v_Buff, I - 1))
    
    On Error GoTo 0
    Exit Function
    
MGetIniValue_Error:
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MGetIniValue of Module ModFileUtils"
End Function

'---------------------------------------------------------------------------------------
' Procedure : MGetTextMaxLine
' Author    : YPN
' Date      : 2017-12-12 17:17
' Purpose   : 获取text文件最大行数（二进制方式）
' Param     : i_TextFile     指定文件
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MGetTextMaxLine(ByVal i_TextFile As String) As Long
    
    Dim v_Str()       As Byte
    Dim v_TextContent As String
    
    On Error GoTo MGetTextMaxLine_Error
    
    If Dir(i_TextFile) <> "" Then          ' 如果文件存在
        Open i_TextFile For Binary As #1
        ReDim v_Str(LOF(1) - 1) As Byte
        Get #1, , v_Str
        Close #1
        v_TextContent = StrConv(v_Str(), vbUnicode)
        MGetTextMaxLine = UBound(Split(v_TextContent, vbCrLf))
    Else
        MsgBox "指定文件不存在：" & i_TextFile, vbExclamation, TS
        Exit Function
    End If
    
    
    On Error GoTo 0
    Exit Function
    
MGetTextMaxLine_Error:
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MGetTextMaxLine of Module ModFileUtils"
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : MSetIniValue
' Author    : YPN
' Date      : 2017-12-15 12:19
' Purpose   : 写入初始化文件（.ini）指定键（Key）和值（Value）
' Param     : i_Section    节
'             i_Key        键
'             i_Value      值
'             i_FileName   完整的INI文件名
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MSetIniValue(ByVal i_Section As String, ByVal i_Key As String, ByVal i_Value As String, ByVal i_FileName As String) As Boolean
    
    
    Dim v_Buff As String * 128
    
    On Error GoTo MSetIniValue_Error
    
    v_Buff = i_Value + Chr(0)
    X = WritePrivateProfileString(i_Section, i_Key, v_Buff, i_FileName)
    
    MSetIniValue = X
    
    On Error GoTo 0
    Exit Function
    
MSetIniValue_Error:
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MSetIniValue of Module ModFileUtils"
    
End Function
