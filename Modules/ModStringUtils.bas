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
' Purpose   : �ж��ַ����Ƿ�Ϊ��
' Param     : i_Str �ַ���
' Return    :
' Remark    :
'---------------------------------------------------------------------------------------
'
Public Function MIsNull(ByVal i_Str As Variant) As Boolean

    If isNull(i_Str) Then
        MIsNull = True
        Exit Function
    End If

    If Trim(i_Str) = "" Then
        MIsNull = True
        Exit Function
    End If
    
    MIsNull = False
    
End Function
