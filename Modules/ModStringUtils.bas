Attribute VB_Name = "ModStringUtils"
'---------------------------------------------------------------------------------------
' Module    : ModStringUtils
' Author    : YPN
' Date      : 2017-06-29 14:46
' Purpose   : 字符串工具类
'---------------------------------------------------------------------------------------

Option Explicit



'---------------------------------------------------------------------------------------
' Procedure : MIsNull
' Author    : YPN
' Date      : 2017-06-29 14:51
' Purpose   : 判断变量是否为空
' Param     : i_Var 变量
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
