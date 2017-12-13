Attribute VB_Name = "ModFileUtils"
'---------------------------------------------------------------------------------------
' Module    : ModFileUtils
' Author    : YPN
' Date      : 2017-12-12 16:17
' Purpose   : 文件工具类
'---------------------------------------------------------------------------------------

Option Explicit

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
