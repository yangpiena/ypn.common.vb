Attribute VB_Name = "ModFileUtils"
'---------------------------------------------------------------------------------------
' Module    : ModFileUtils
' Author    : YPN
' Date      : 2017-12-12 16:17
' Purpose   : �ļ�������
'---------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : MApendText
' Author    : YPN
' Date      : 2017-12-12 16:20
' Purpose   : ׷�����ݵ�ָ���ļ�
' Param     : i_TextFile     ָ���ļ�
'             i_ApendContent ׷������
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Sub MApendText(ByVal i_TextFile As String, ByVal i_ApendContent As String)
    
    On Error GoTo MApendText_Error
    
    If Dir(i_TextFile) <> "" Then          ' ����ļ�����
        Open i_TextFile For Append As #1   ' ��׷�ӷ�ʽ���ļ�
        'Print #1                          ' Ϊ�˷�ֹԭ�ļ�ĩβû�л��У�������Ļ���
        Print #1, i_ApendContent
        Close #1
    Else
        MsgBox "ָ���ļ������ڣ�" & i_TextFile, vbExclamation, TS
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
' Purpose   : ��ȡtext�ļ���������������Ʒ�ʽ��
' Param     : i_TextFile     ָ���ļ�
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MGetTextMaxLine(ByVal i_TextFile As String) As Long
    
    Dim v_Str()       As Byte
    Dim v_TextContent As String
    
    On Error GoTo MGetTextMaxLine_Error
    
    If Dir(i_TextFile) <> "" Then          ' ����ļ�����
        Open i_TextFile For Binary As #1
        ReDim v_Str(LOF(1) - 1) As Byte
        Get #1, , v_Str
        Close #1
        v_TextContent = StrConv(v_Str(), vbUnicode)
        MGetTextMaxLine = UBound(Split(v_TextContent, vbCrLf))
    Else
        MsgBox "ָ���ļ������ڣ�" & i_TextFile, vbExclamation, TS
        Exit Function
    End If
    
    
    On Error GoTo 0
    Exit Function
    
MGetTextMaxLine_Error:
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MGetTextMaxLine of Module ModFileUtils"
    
End Function
