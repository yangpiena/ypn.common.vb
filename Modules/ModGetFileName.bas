Attribute VB_Name = "ModGetFileName"
'---------------------------------------------------------------------------------------
' Module    : ModGetFileName
' Author    : YPN
' Date      : 2017-06-28 17:44
' Purpose   : 获取文件名
'---------------------------------------------------------------------------------------

Option Explicit



'---------------------------------------------------------------------------------------
' Procedure : MGetFileNameInPath
' Author    : YPN
' Date      : 2017-06-28 17:45
' Purpose   : 从指定全路径中获取文件名
' Param     : i_Path      指定全路径
'             i_HasSuffix 是否包含后缀名，默认包含
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
' Purpose   : 从文件名中获取后缀名
' Param     : i_FileName 文件名
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MGetSuffixInFileName(ByVal i_FileName As String) As String

    MGetSuffixInFileName = IIf(InStr(i_FileName, "."), Right(i_FileName, Len(i_FileName) - InStrRev(i_FileName, ".")), vbNullString)

End Function
